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
