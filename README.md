# freee API × Google Apps Script（GAS）連携

## 概要
freee会計APIから取引一覧データを取得し、Googleスプレッドシートに転記するGoogle Apps Script（GAS）のサンプルです。

---

## 構成
- GAS（main.js）
- freee会計API（OAuth2認証）
- Googleスプレッドシート

---

## セットアップ手順

1. **freee開発者アカウント作成・アプリ登録**
   - [freee開発者コンソール](https://developer.freee.co.jp/)でアプリを作成し、クライアントID・シークレット・リダイレクトURIを取得
   - リダイレクトURI例：
     `https://script.google.com/macros/d/【スクリプトID】/usercallback`
   - アプリを「公開」または「テスト」状態にする

2. **Google Apps Scriptプロジェクト作成**
   - スクリプトエディタでmain.jsを作成し、本リポジトリの内容を貼り付け
   - `CLIENT_ID`、`CLIENT_SECRET`、`REDIRECT_URI`をfreeeの値に書き換え

3. **OAuth2ライブラリ導入**
   - スクリプトエディタの「ライブラリ」からID `1B7Jc7wG2QbFJi7dUo6g9gK2kT9gkQ4v1b1Jg1j1Jg1Jg1Jg1Jg1Jg1Jg1` を追加

4. **スプレッドシートと紐付け**
   - スクリプトを紐付けたGoogleスプレッドシートを開く

---

## 使い方

1. **main関数を実行**
   - 初回はログに表示される認証URLにアクセスし、freeeの認証画面で許可
   - 認証後は自動で会社IDを取得し、取引一覧をスプレッドシートに出力

2. **2回目以降はmain関数を実行するだけで自動取得**

---

## 注意点
- 初回認証時は必ずユーザー操作が必要です（OAuth2仕様）
- freeeアプリのリダイレクトURIとGASのREDIRECT_URIは完全一致させてください
- 会社IDはAPIから自動取得します
- 取引一覧は最大50件取得（必要に応じてパラメータ調整可）
- 取得データやAPIレスポンスはLogger.logでコンソール出力されます

---

## 参考リンク
- [freee会計APIリファレンス](https://developer.freee.co.jp/reference/accounting/reference)
- [GAS OAuth2ライブラリ導入手順](https://github.com/googleworkspace/apps-script-oauth2)

---

何か問題があればissueまたはPRでご連絡ください。 