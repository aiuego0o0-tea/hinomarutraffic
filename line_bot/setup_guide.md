# tea LINE Bot — セットアップガイド
*所要時間：約30分 / オーナー作業用*

---

## 全体の流れ

```
STEP 1: LINE公式アカウント作成（LINE Developers）
STEP 2: Vercelにデプロイ
STEP 3: 環境変数を設定
STEP 4: LINE側にWebhook URLを登録
STEP 5: メンバーを招待して動作確認
```

---

## STEP 1 — LINE公式アカウント作成

### 1-1. LINE Developersにログイン
1. https://developers.line.biz/ja/ を開く
2. 右上「ログイン」→ 自分のLINEアカウントでログイン

### 1-2. プロバイダーを作成
1. 「新規プロバイダー作成」をクリック
2. 名前：`ほめ丸チーム`（なんでもOK）

### 1-3. チャンネルを作成
1. 「Messaging API」を選択
2. 以下を入力：

| 項目 | 入力内容 |
|---|---|
| チャンネル名 | `tea` |
| チャンネル説明 | `チームのAIアシスタント` |
| 大業種 | サービス業 |
| 小業種 | 任意 |

3. 「作成」→ 同意にチェックして確認

### 1-4. キーを取得（後で使う）
チャンネル作成後、以下2つをメモ帳にコピーしておく：

**① Channel Secret**
- 「Messaging API設定」→「チャンネルシークレット」

**② Channel Access Token**
- 「Messaging API設定」→ 一番下「チャンネルアクセストークン（長期）」→「発行」ボタン

---

## STEP 2 — Vercelにデプロイ

### 2-1. GitHubにリポジトリを作る
1. https://github.com を開く
2. 右上「+」→「New repository」
3. 名前：`tea-line-bot`
4. Private に設定して「Create repository」

### 2-2. コードをGitHubにアップ
ターミナルで以下を実行：

```bash
cd /Users/aiuego/Desktop/AI\ CEO_tea/clients/hinomarutraffic/line_bot
git init
git add .
git commit -m "first commit"
git branch -M main
git remote add origin https://github.com/[あなたのGitHubユーザー名]/tea-line-bot.git
git push -u origin main
```

### 2-3. Vercelにデプロイ
1. https://vercel.com を開く
2. 「Sign Up」→「Continue with GitHub」
3. 「Add New Project」→ `tea-line-bot` を選択
4. **「Deploy」を押す前に** → STEP 3へ（環境変数を先に設定）

---

## STEP 3 — 環境変数を設定（Vercel）

「Configure Project」画面の「Environment Variables」に以下を追加：

| Name | Value |
|---|---|
| `LINE_CHANNEL_SECRET` | STEP 1でメモしたChannel Secret |
| `LINE_CHANNEL_ACCESS_TOKEN` | STEP 1でメモしたChannel Access Token |
| `ANTHROPIC_API_KEY` | .envファイルに入っているAPIキー |

全部入力したら「Deploy」を押す。

デプロイが完了すると URL が発行される（例：`https://tea-line-bot-xxxx.vercel.app`）  
→ このURLをメモしておく

---

## STEP 4 — LINE側にWebhook URLを登録

1. LINE Developers → 作ったチャンネル → 「Messaging API設定」
2. 「Webhook URL」に以下を入力：
   ```
   https://[VercelのURL]/api/webhook
   ```
   例：`https://tea-line-bot-xxxx.vercel.app/api/webhook`

3. 「検証」ボタンを押す → 「成功」と表示されればOK
4. 「Webhookの利用」を**ON**にする

### 自動返信を止める設定
1. 「Messaging API設定」→「LINE公式アカウント機能」
2. 「自動応答メッセージ」→「編集」→**オフ**にする
3. 「あいさつメッセージ」→ 任意で設定

---

## STEP 5 — メンバーを招待して動作確認

### LINE公式アカウントを友だち追加
「Messaging API設定」→「QRコード」を表示  
→ ハーブ・杉原・一柳に共有して友だち追加してもらう

### 動作確認
LINEで話しかけてみる：
```
「企画案のテンプレート作って」
「今週の週報を作って」
「メール文章考えて」
```

---

## トラブル対応

| 症状 | 確認箇所 |
|---|---|
| 返信が来ない | Vercelのデプロイが成功しているか確認 |
| 「Invalid signature」エラー | LINE_CHANNEL_SECRETが正しいか確認 |
| 「Unauthorized」エラー | LINE_CHANNEL_ACCESS_TOKENが正しいか確認 |
| Vercelのログを確認 | Vercelダッシュボード → 該当プロジェクト → 「Logs」 |

---

## 完了後のURL構造

```
https://[VercelのURL]/api/webhook  ← LINEからのメッセージを受け取る口
```

---

*tea LINE Bot / setup_guide.md / 2026-05-20*
