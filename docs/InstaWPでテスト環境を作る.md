# InstaWPで無料のWordPressテスト環境を作る

GASはGoogleのサーバー上で動くため、接続先のWordPressはインターネットに公開されたURLである必要がある(自分のPC上の`localhost`には接続できない)。InstaWPを使うと、公開URL付きのWordPressを数分で無料で用意できる。

## 1. サイトを作る

1. [instawp.com](https://instawp.com) にアクセスし、アカウントを作成してログインする
2. ダッシュボードの「Sites」ページで「Add Sites」をクリック
3. 「From Scratch」を選ぶ
4. WordPressのバージョンなどはデフォルトのままでよい
5. 無料の一時ドメイン(`https://(ランダムな文字列).instawp.site` のような形)を選ぶ(独自ドメインは有料)
6. テーマ・プラグインの選択はスキップしてよい
7. 「Create Site」をクリックして作成する

作成が終わると、サイトの公開URL(例: `https://xxxxx.instawp.site`)と管理画面へのログイン情報が表示される。「Magic Login」ボタンからワンクリックで管理画面に入れる。

> **注意**: 無料サイトは一定期間で自動的に使えなくなることがある(pay-as-you-goの課金体系)。テストが終わったら早めに連携確認を済ませておくと安心。

## 2. アプリケーションパスワードを発行する

WordPress 5.6以降は、プラグインなしで「アプリケーションパスワード」機能が使える(ただしサイトがHTTPSであることが条件。InstaWPのサイトは標準でHTTPSなので問題ない)。

1. 作成したサイトの管理画面にログインする(Magic Loginでよい)
2. 左メニュー「ユーザー → プロフィール」を開く
3. ページ下部の「アプリケーションパスワード」セクションまでスクロール
4. 新しいアプリケーションパスワード名(例: `gas-blog-automation`)を入力して「新規追加」
5. 表示されたパスワード(スペース区切りの英数字)を**その場でコピー**する(閉じると二度と表示されない)

もし「アプリケーションパスワード」のセクションが見当たらない場合は、サイトがHTTPSになっているか確認する(InstaWPの標準ドメインなら通常は問題ない)。

## 3. GAS側に設定する

1. スプレッドシートの「設定」シートに、WordPress URL(例: `https://xxxxx.instawp.site`)とユーザー名を入力する
2. メニュー「記事生成 → 初期設定(APIキーなど)」から、先ほど発行したアプリケーションパスワードを登録する

これで `README.md` の「5. 動かしてみる」の手順に進める。

## うまく繋がらないときは

- `docs/トラブルシューティング.md` の「WordPress投稿エラー」の項目を参照
- ブラウザで `(InstaWPのURL)/wp-json/wp/v2/posts` を開いて、JSON形式のレスポンスが返ってくるか確認する(エラーページが出る場合はREST APIが無効化されている可能性がある)

---

Sources:
- [Create Site - InstaWP Docs](https://docs.instawp.com/sites/create-site)
- [Application Passwords - InstaWP](https://instawp.com/plugin/application-passwords/)
