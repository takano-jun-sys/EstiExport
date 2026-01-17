# 自動デプロイ設定ガイド

このガイドでは、GitHubからGoogle Apps Scriptへの自動デプロイを設定する方法を説明します。

## 概要

**clasp (Command Line Apps Script Projects)** と **GitHub Actions** を使用して、GitHubにコードをpushすると自動的にGASにデプロイされる仕組みを構築します。

```
ローカル開発 → Git push → GitHub Actions → clasp → GAS自動デプロイ
```

## メリット

- ✅ コードをGitで管理できる
- ✅ ローカルエディタ（VSCode等）で開発可能
- ✅ プルリクエストでコードレビュー可能
- ✅ mainブランチにマージすると自動デプロイ
- ✅ バージョン管理とロールバックが容易

---

## セットアップ手順

### ステップ1: claspのローカルインストール

#### 1-1. Node.jsのインストール

まだインストールしていない場合：
- https://nodejs.org/ から最新LTS版をダウンロード
- インストール後、ターミナルで確認：

```bash
node --version
npm --version
```

#### 1-2. claspのインストール

```bash
npm install -g @google/clasp
```

確認：

```bash
clasp --version
```

### ステップ2: claspのログイン

#### 2-1. Googleアカウントでログイン

```bash
clasp login
```

- ブラウザが開き、Googleアカウントの選択画面が表示されます
- GASプロジェクトを持っているアカウントを選択
- 権限を許可

#### 2-2. 認証情報の確認

ログインすると、以下のファイルが作成されます：

- **macOS/Linux**: `~/.clasprc.json`
- **Windows**: `%USERPROFILE%\.clasprc.json`

このファイルの内容を後で使用するので、確認しておきます：

```bash
# macOS/Linux
cat ~/.clasprc.json

# Windows (PowerShell)
Get-Content $env:USERPROFILE\.clasprc.json
```

内容例：
```json
{
  "token": {
    "access_token": "...",
    "refresh_token": "...",
    "scope": "...",
    "token_type": "Bearer",
    "expiry_date": ...
  },
  "oauth2ClientSettings": {
    "clientId": "...",
    "clientSecret": "...",
    "redirectUri": "..."
  },
  "isLocalCreds": false
}
```

**重要**: この内容をメモ帳などにコピーしておきます（後でGitHub Secretsに設定）

### ステップ3: GASプロジェクトとclaspを連携

#### 3-1. Script IDの取得

1. **Apps Scriptエディタを開く**
   - 作成したスプレッドシート → 拡張機能 → Apps Script

2. **プロジェクトの設定を開く**
   - 左側の歯車アイコン（プロジェクトの設定）をクリック

3. **Script IDをコピー**
   - 「スクリプト ID」の値をコピー
   - 例: `1a2b3c4d5e6f7g8h9i0j1k2l3m4n5o6p7q8r9s0t`

#### 3-2. .clasp.jsonの作成

このリポジトリをローカルにクローン済みの場合：

```bash
cd EstiExport
```

`.clasp.json`ファイルを作成：

```bash
cp .clasp.json.template .clasp.json
```

`.clasp.json`を編集して、Script IDを設定：

```json
{
  "scriptId": "YOUR_SCRIPT_ID_HERE",
  "rootDir": "."
}
```

`YOUR_SCRIPT_ID_HERE`を、先ほどコピーしたScript IDに置き換えます。

#### 3-3. 動作確認（ローカルからpush）

```bash
clasp push
```

成功すると：
```
└─ Code.gs
└─ Config.gs
└─ PDFExporter.gs
└─ QuoteGenerator.gs
└─ Setup.gs
└─ appsscript.json
Pushed 6 files.
```

Apps Scriptエディタを開いて、コードが更新されていることを確認します。

### ステップ4: GitHub Secretsの設定

GitHubリポジトリでSecretsを設定し、GitHub Actionsがclaspを使えるようにします。

#### 4-1. GitHubリポジトリの設定ページを開く

1. GitHubでリポジトリを開く
2. 「Settings」タブをクリック
3. 左側のメニューから「Secrets and variables」→「Actions」をクリック

#### 4-2. CLASPRC_JSONの追加

1. **「New repository secret」をクリック**

2. **Secretを入力**：
   - **Name**: `CLASPRC_JSON`
   - **Value**: ステップ2-2でコピーした`~/.clasprc.json`の内容をそのまま貼り付け

3. **「Add secret」をクリック**

#### 4-3. CLASP_JSONの追加

1. **「New repository secret」をクリック**

2. **Secretを入力**：
   - **Name**: `CLASP_JSON`
   - **Value**: 先ほど作成した`.clasp.json`の内容を貼り付け

```json
{
  "scriptId": "1a2b3c4d5e6f7g8h9i0j1k2l3m4n5o6p7q8r9s0t",
  "rootDir": "."
}
```

3. **「Add secret」をクリック**

### ステップ5: 自動デプロイのテスト

#### 5-1. コードを修正

例えば、`Code.gs`にコメントを追加：

```javascript
/**
 * EstiExport - AppSheet連携見積書作成システム
 * メインエントリーポイント
 * バージョン: 1.0.1
 */
```

#### 5-2. GitHubにpush

```bash
git add Code.gs
git commit -m "Update version comment"
git push origin main
```

**注意**: ブランチ名が`master`の場合は`main`を`master`に変更してください

#### 5-3. GitHub Actionsの実行を確認

1. GitHubリポジトリページを開く
2. 「Actions」タブをクリック
3. 最新のワークフロー実行を確認
4. 緑色のチェックマークが表示されれば成功

#### 5-4. GASで確認

1. Apps Scriptエディタを開く
2. ページを更新（F5）
3. `Code.gs`のコメントが更新されていることを確認

---

## 開発ワークフロー

### 通常の開発フロー

```bash
# 1. ローカルで開発
# VSCodeなどでコードを編集

# 2. ローカルでテスト（任意）
clasp push
# Apps Scriptエディタで動作確認

# 3. Gitにコミット
git add .
git commit -m "Add new feature"

# 4. GitHubにpush
git push origin main

# 5. 自動デプロイが実行される
# GitHub Actions が clasp push を実行
# GASに自動的にデプロイされる
```

### プルリクエストを使った開発フロー（推奨）

```bash
# 1. 新しいブランチを作成
git checkout -b feature/new-function

# 2. ローカルで開発
# コードを編集

# 3. ローカルでテスト
clasp push
# 動作確認

# 4. GitHubにpush
git add .
git commit -m "Add new function"
git push origin feature/new-function

# 5. プルリクエストを作成
# GitHubでPRを作成

# 6. レビュー後、mainにマージ
# マージすると自動的にGASにデプロイ
```

---

## トラブルシューティング

### エラー: "Exceeded rate limit"

**原因**: claspのAPI呼び出し制限に達した

**解決方法**:
- 数分待ってから再実行
- 短時間に何度もpushしない

### エラー: "Permission denied"

**原因**: Script IDが間違っているか、アクセス権限がない

**解決方法**:
1. `.clasp.json`のScript IDを確認
2. claspでログインしたGoogleアカウントがGASプロジェクトのオーナーか確認
3. 再ログイン: `clasp login --creds ~/.clasprc.json`

### GitHub Actionsが失敗する

**原因**: Secretsが正しく設定されていない

**解決方法**:
1. GitHubの「Settings」→「Secrets and variables」→「Actions」を確認
2. `CLASPRC_JSON`と`CLASP_JSON`が正しく設定されているか確認
3. JSONの形式が正しいか確認（余分な改行やスペースがないか）

### ローカルとGASでコードが同期されない

**原因**: .claspignoreの設定またはファイル構造の問題

**解決方法**:
1. `clasp status`で状態を確認
2. `clasp pull`でGASからコードを取得して比較
3. `clasp push --force`で強制的にpush

---

## 高度な設定

### デプロイメントの作成

バージョン管理されたデプロイメントを作成：

```bash
# 新しいデプロイメントを作成
clasp deploy -d "Version 1.0.0"

# デプロイメント一覧を確認
clasp deployments
```

### 特定のバージョンをデプロイ

GitHub Actionsでタグベースのデプロイを設定する場合、`.github/workflows/deploy.yml`を修正：

```yaml
on:
  push:
    tags:
      - 'v*'
```

タグをpushすると自動デプロイ：

```bash
git tag v1.0.0
git push origin v1.0.0
```

### 複数環境へのデプロイ

開発環境と本番環境を分ける場合：

**開発環境用**: `.clasp.dev.json`
```json
{
  "scriptId": "DEV_SCRIPT_ID",
  "rootDir": "."
}
```

**本番環境用**: `.clasp.prod.json`
```json
{
  "scriptId": "PROD_SCRIPT_ID",
  "rootDir": "."
}
```

デプロイコマンド：

```bash
# 開発環境
clasp push --project .clasp.dev.json

# 本番環境
clasp push --project .clasp.prod.json
```

---

## ベストプラクティス

### 1. ローカル開発環境を整える

- **エディタ**: VSCode + Apps Script拡張機能
- **Linter**: ESLint設定
- **フォーマッター**: Prettier設定

### 2. ブランチ戦略

- `main` / `master`: 本番環境
- `develop`: 開発環境
- `feature/*`: 機能開発ブランチ

### 3. コミットメッセージ

わかりやすいコミットメッセージを使用：

```
feat: Add email notification feature
fix: Fix PDF export bug
docs: Update README
refactor: Improve quote generation logic
```

### 4. テストの追加

GASUnit などを使用して自動テストを追加し、GitHub Actionsで実行

### 5. 環境変数の管理

機密情報は`.env`ファイルやGitHub Secretsで管理し、コードにハードコードしない

---

## まとめ

これで、GitHubからGASへの自動デプロイ環境が構築できました！

**開発フロー**:
1. ローカルでコード編集（VSCodeなど）
2. `git push`でGitHubにpush
3. GitHub Actionsが自動的にGASにデプロイ
4. スプレッドシートで即座に反映

**次のステップ**:
- ローカル開発環境を整える
- プルリクエストベースの開発フローを導入
- 自動テストの追加
- 複数環境（dev/prod）の構築

質問があれば、GitHubのIssuesで報告してください！
