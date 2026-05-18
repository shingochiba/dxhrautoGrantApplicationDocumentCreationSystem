# Streamlit Community Cloud デプロイ手順書

人材開発支援助成金 書類自動作成ツールを **無料で5人が使える状態** にするための手順です。

---

## 全体の流れ

1. GitHub アカウント作成（1回だけ）
2. コードを GitHub にアップロード
3. Streamlit Community Cloud でデプロイ
4. シークレット（パスワード）の設定
5. 利用者にURLを共有

所要時間：**約30分**

---

## Step 1: GitHub アカウント作成

1. [https://github.com/signup](https://github.com/signup) にアクセス
2. メールアドレス、パスワードを入力して登録
3. メール認証を完了

すでに GitHub アカウントをお持ちの場合はスキップ。

---

## Step 2: GitHub Desktop のインストール（コマンド操作不要に）

コマンドライン操作が苦手な場合、GUI ツール **GitHub Desktop** を使うと簡単です。

1. [https://desktop.github.com/](https://desktop.github.com/) から GitHub Desktop をダウンロード&インストール
2. 起動して GitHub アカウントでサインイン

---

## Step 3: プライベートリポジトリ作成

1. GitHub Desktop の **File → New Repository** をクリック
2. 以下を入力:
   - **Name**: `document-generator`（任意）
   - **Local path**: 任意のフォルダ（デスクトップなど）
   - **Initialize this repository with a README**: チェックON
   - **Git Ignore**: `None` のまま（既にプロジェクトに `.gitignore` があるため）
3. **Create Repository** をクリック
4. **Publish repository** ボタンをクリック
5. チェック: **Keep this code private**（必ずONに！）
6. **Publish repository** をクリック

---

## Step 4: コードをリポジトリに配置

1. エクスプローラーで、作成したリポジトリのローカルフォルダを開く
2. `document_generator_new/` の中身を**すべて**コピーして、リポジトリフォルダに貼り付け
   - ただし `.streamlit/secrets.toml` と `data/` と `output/` は**コピーしない**（`.gitignore` で除外済みなので結果的に同じですが）
3. GitHub Desktop に戻ると変更が検出される
4. 左下の **Summary** に `Initial upload` など入力
5. **Commit to main** をクリック
6. **Push origin** をクリック

これで GitHub にコードがアップロードされます。

### 確認ポイント（重要）

GitHub のリポジトリページを開いて、**`.streamlit/secrets.toml` ファイルが無い**ことを確認してください。あった場合はパスワードが公開されてしまいます。`.streamlit/` フォルダ内は `config.toml` と `secrets.toml.example` のみのはずです。

---

## Step 5: Streamlit Community Cloud でデプロイ

1. [https://streamlit.io/cloud](https://streamlit.io/cloud) にアクセス
2. **Continue with GitHub** でサインイン
3. 初回は GitHub への認証許可（**Install & Authorize**）
4. 右上の **Create app** → **Deploy a public app from GitHub** を選択
5. 以下を入力:
   - **Repository**: `あなたのユーザー名/document-generator`
   - **Branch**: `main`
   - **Main file path**: `app.py`
   - **App URL (optional)**: 好きな名前（例: `dxhr-docgen`）
6. **Advanced settings** をクリック
7. **Python version**: `3.12` を選択（または 3.11）
8. **Secrets** 欄に以下をコピー&ペースト:

```toml
[passwords]
"hashimoto@dxhr.inc" = "Dxhr2026#"
"shimizu.a@dxhr.inc" = "Dxhr2026#"
"arimura@dxhr.inc"   = "Dxhr2026#"
"ohira@dxhr.inc"     = "Dxhr2026#"
"chiba@dxhr.inc"     = "Dxhr2026#"
```

9. **Save** → **Deploy!** をクリック
10. デプロイ完了まで 2〜5分ほど待つ

---

## Step 6: 動作確認

1. 発行された URL（例: `https://dxhr-docgen.streamlit.app/`）にアクセス
2. ログイン画面が表示される
3. メールアドレスとパスワードを入力してログイン
4. 会社情報入力画面が表示されれば成功

---

## Step 7: 利用者にURLを共有

5人それぞれにメールなどで以下を伝える:

```
【書類作成ツールご案内】

URL: https://dxhr-docgen.streamlit.app/
ログインID: あなたのメールアドレス（例: hashimoto@dxhr.inc）
パスワード: Dxhr2026#

※ 初回はパスワードを変更することをおすすめします（下記「パスワード変更」参照）
※ ブラウザは Chrome / Edge / Safari 推奨
```

---

## 運用上の注意

### データの永続性について
- Streamlit Cloud のファイル保存は **再デプロイ時に消える** 可能性があります
- よって、会社情報は **毎回 Excel インポート** する運用が安全です
- 同じセッション中は session_state に保存されるため、連続操作では再入力不要

### パスワード変更
1. GitHub リポジトリの `.streamlit/secrets.toml.example` は変更**せず**
2. Streamlit Cloud のアプリ設定 → **Secrets** 欄を書き換え → **Save**
3. アプリが自動で再起動されて反映される

### 利用人数の追加
`[passwords]` セクションに追加するだけ:
```toml
[passwords]
"hashimoto@dxhr.inc" = "Dxhr2026#"
"newuser@dxhr.inc" = "NewPassword#"
...
```

### コードの更新
1. ローカルの `document_generator_new/` でコードを編集
2. GitHub Desktop で Commit & Push
3. Streamlit Cloud が自動検知して再デプロイ（1〜2分）

---

## トラブルシューティング

### 「ログインボタンが反応しない」
→ ブラウザで Ctrl+Shift+R で強制リロード

### 「認証情報が設定されていません」エラー
→ Streamlit Cloud の Secrets が正しく設定されているか確認。`[passwords]` セクションが必須。

### 「アプリが Over Capacity」と表示される
→ Streamlit Cloud の一時的な混雑。数分後に再アクセスで解決。

### デプロイが失敗する
→ Streamlit Cloud の管理画面で **Logs** を確認。依存パッケージエラーなら `requirements.txt` を見直し。

---

## セキュリティ上の推奨事項

- パスワードは定期的に変更する
- 5人のパスワードは全員同じにせず、各自別のパスワードにする方が安全
- 退職者が出たら即座に該当ユーザーの行を削除
- secrets.toml は **絶対に GitHub にコミットしない**（`.gitignore` で防止済み）

---

## BTS (バグ・タスク管理) 連携の設定

BTS スプレッドシートを非公開のまま Streamlit アプリから読み込むため、Google Cloud のサービスアカウントを使用します。

### Step A: サービスアカウント作成

1. [Google Cloud Console](https://console.cloud.google.com/) にログイン
2. 新規プロジェクト作成 (既存でもOK / 例: `dxhr-docgen`)
3. **APIとサービス → ライブラリ** で「Google Sheets API」を検索→**有効にする**
4. **APIとサービス → 認証情報** → 「**認証情報を作成**」→「**サービスアカウント**」
5. 任意の名前 (例: `bts-reader`) を入力→作成→「完了」
6. 作成したサービスアカウントをクリック→「**鍵**」タブ→「鍵を追加」→「**新しい鍵を作成**」→形式「**JSON**」→作成
7. JSONファイルが自動ダウンロードされる (例: `dxhr-docgen-xxxxx.json`)

### Step B: スプレッドシートを共有

1. JSON ファイルを開いて `client_email` の値 (`xxx@xxx.iam.gserviceaccount.com`) をコピー
2. BTS スプレッドシートを開き、右上「**共有**」をクリック
3. コピーしたメールアドレスを貼り付け、権限を「**閲覧者**」のままで追加
4. 「送信」をクリック (通知は不要)

### Step C: Streamlit Cloud に登録

#### 🛠 一発変換 (おすすめ)

ダウンロードした JSON を Streamlit Secrets 用 TOML に自動変換するスクリプトを同梱しています。

```powershell
# ローカルでダウンロードした JSON ファイルを指定
python tools/convert_sa_json_to_toml.py "C:\path\to\dxhr-docgen-xxxxx.json"
```

出力された `[gcp_service_account]` セクションを、Streamlit Cloud の **Settings → Secrets** で既存の `[passwords]` セクションの下にそのまま貼り付ければ完了。`private_key` の改行も三連クォートで正しくエスケープされた状態で出力されます。

#### 手動で TOML を書く場合

1. Streamlit Cloud のアプリ管理画面 → **Settings** → **Secrets**
2. 既存の `[passwords]` セクションの**下に**以下を追記 (各値は **JSON の対応するフィールドの中身そのもの** に置換):

```toml
[gcp_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "abc123def456..."
private_key = """-----BEGIN PRIVATE KEY-----
MIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQ...
...
...vQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQ==
-----END PRIVATE KEY-----
"""
client_email = "bts-reader@your-project-id.iam.gserviceaccount.com"
client_id = "123456789012345678901"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/bts-reader%40your-project-id.iam.gserviceaccount.com"
```

3. **Save** をクリックすると、アプリが自動再起動

#### ⚠️ よくある TOML エラーと対処

**「Invalid format: please enter valid TOML.」** が出る場合、ほぼ `private_key` の書き方が原因です。

**正しい方法 (推奨): 三連クォート `"""` で囲み、改行はそのまま**

JSON ファイル内では `private_key` は1行に潰されて `\n` 文字列として書かれていますが、TOML では以下のように **三連クォートで囲んで実改行に展開して** 貼り付けます:

```toml
private_key = """-----BEGIN PRIVATE KEY-----
MIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQ...
（複数行つづく）
-----END PRIVATE KEY-----
"""
```

JSON の `"private_key": "-----BEGIN ...\nXXX\n...\n-----END PRIVATE KEY-----\n"` から TOML に変換する手順:
1. 先頭の `"private_key": "` と末尾の `"` を削除
2. 文字列中の `\n` を実際の改行に置換 (テキストエディタの「\n を改行に置換」機能を使う)
3. `private_key = """` と `"""` で囲む

**代替方法: 1行に書いて `\n` のまま (こちらでも動く)**

```toml
private_key = "-----BEGIN PRIVATE KEY-----\nMIIEvQ...\n-----END PRIVATE KEY-----\n"
```

`\n` をバックスラッシュ + n の2文字のまま (改行を入れず) 1行で書く方法。JSON の値をそのままダブルクォート間に貼ればよい。

**やってはいけないこと:**
- ダブルクォート `"..."` で囲んだ中に **実改行を入れる** → これがエラーの主因。基本文字列は単一行のみ許可されます。
- 値を囲むクォートを忘れる
- 文字列の前後に余計な空白や `,` が残っている

#### 他フィールドの注意点
- `project_id`, `client_email`, `private_key_id`, `client_id`, `client_x509_cert_url` も JSON の値をそのままダブルクォートで囲む
- `client_x509_cert_url` の `@` 記号は **URLエンコードで `%40`** になっている (JSON の表記そのままでOK)

### Step D: ローカル開発時の設定 (任意)

ローカルで動かす場合は以下のいずれかで認証情報を渡す:

- **方法1**: ダウンロードした JSON を `config/google_service_account.json` にリネームして配置 (`.gitignore` で除外済)
- **方法2**: `.streamlit/secrets.toml` に上記 `[gcp_service_account]` セクションを追加
- **方法3**: 環境変数 `GOOGLE_APPLICATION_CREDENTIALS` に JSON のフルパスを設定

### 動作確認

アプリ起動後、サイドバーの「🐛 BTS (バグ・タスク管理)」をクリックして一覧が表示されればOK。
権限エラーが出た場合は Step B のスプレッドシート共有を見直してください。
