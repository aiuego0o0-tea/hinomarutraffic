import Anthropic from '@anthropic-ai/sdk';
import { messagingApi } from '@line/bot-sdk';
import crypto from 'crypto';

// Vercel のデフォルトbodyパーサーを無効化（LINE署名検証のためraw bodyが必要）
export const config = {
  api: { bodyParser: false },
};

const lineClient = new messagingApi.MessagingApiClient({
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
});

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const SYSTEM_PROMPT = `あなたは「tea」です。ソーシャル交通ほめ丸会社チームのAIアシスタントです。

【チーム情報】
- メンバー：ハーブ、杉原、一柳、オーナー
- 事業：日の丸交通株式会社のSNS採用マーケティング運用支援
- 主なプラットフォーム：TikTok・Instagram・X・LINE
- 目標：SNS→LINE登録→採用説明会→入社の一気通貫モデル構築

【teaの行動指針】
- 結論から、短く答える
- わからないことは「確認が必要です」と正直に伝える
- フレンドリーだがビジネスに責任感のあるトーン
- 日本語で答える
- LINE画面を意識して読みやすく書く（長文は箇条書き・改行を活用）
- 企画・文章・絵コンテ・メールなど業務全般を手伝う`;

// リクエストのraw bodyを取得
function getRawBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', (chunk) => chunks.push(chunk));
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

// LINE署名を検証
function validateSignature(rawBody, signature, secret) {
  const hash = crypto
    .createHmac('SHA256', secret)
    .update(rawBody)
    .digest('base64');
  return hash === signature;
}

// メッセージをClaudeに送って返答を得る
async function askTea(userMessage) {
  const response = await anthropic.messages.create({
    model: 'claude-sonnet-4-5',
    max_tokens: 1000,
    system: SYSTEM_PROMPT,
    messages: [{ role: 'user', content: userMessage }],
  });
  return response.content[0].text;
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).send('Method Not Allowed');
  }

  // raw body 取得
  const rawBody = await getRawBody(req);
  const signature = req.headers['x-line-signature'];

  // 署名検証（セキュリティ）
  if (!validateSignature(rawBody, signature, process.env.LINE_CHANNEL_SECRET)) {
    console.error('Invalid signature');
    return res.status(401).send('Invalid signature');
  }

  const { events } = JSON.parse(rawBody.toString());

  // 各イベントを並列処理
  await Promise.all(
    events.map(async (event) => {
      // テキストメッセージ以外は無視
      if (event.type !== 'message' || event.message.type !== 'text') return;

      try {
        const replyText = await askTea(event.message.text);

        await lineClient.replyMessage({
          replyToken: event.replyToken,
          messages: [
            {
              type: 'text',
              // LINEの文字数上限は5000文字
              text: replyText.slice(0, 4999),
            },
          ],
        });
      } catch (error) {
        console.error('tea error:', error);
        // エラー時もユーザーに通知
        await lineClient
          .replyMessage({
            replyToken: event.replyToken,
            messages: [
              {
                type: 'text',
                text: '⚠️ エラーが発生しました。もう一度試してみてください。',
              },
            ],
          })
          .catch(() => {}); // reply失敗しても握りつぶす
      }
    })
  );

  res.status(200).json({ ok: true });
}
