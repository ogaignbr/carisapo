# carisapo（キャリサポ LP）

キャリア支援サービス「キャリサポ」のランディングページ（静的 HTML）です。

## 公開URL（GitHub Pages）

リポジトリで Pages を有効化すると、次のような URL で閲覧できます。

`https://ogaignbr.github.io/carisapo/`

## GitHub での設定（初回のみ）

1. このリポジトリを GitHub で開く: [ogaignbr/carisapo](https://github.com/ogaignbr/carisapo)
2. **Settings**（設定）→ 左メニュー **Pages**
3. **Build and deployment** の **Source** で **GitHub Actions** を選択する  
   （ブランチ指定ではなく、ワークフローでデプロイする方式です）
4. `main` へプッシュすると、`.github/workflows/pages.yml` が自動実行され、サイトが更新されます

### うまく表示されないとき

- **Actions** タブでワークフローが成功（緑）になっているか確認してください
- 初回デプロイ直後は数分かかることがあります
- 画像はリポジトリ直下の `1.jpg`〜`9.jpg` を参照しています。同じフォルダに置いてプッシュしてください

## ローカルでの確認

`index.html` をブラウザで開いてください（画像と同じディレクトリに配置した状態）。
