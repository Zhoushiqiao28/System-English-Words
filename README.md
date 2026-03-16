# System English Words

`system.xlsx` をもとにした、オフライン学習向けの英単語アプリです。

## いちばん簡単な使い方

自分でサーバーを立てたくない場合は、`GitHub Pages` に公開して iPhone の Safari から開く方法がおすすめです。

1. このフォルダを GitHub のリポジトリにアップロードします。
2. GitHub の `Settings > Pages` を開きます。
3. `Build and deployment` を `GitHub Actions` にします。
4. `main` ブランチに push すると、自動で公開されます。
5. 公開された `https://...github.io/...` の URL を iPhone の Safari で開きます。
6. 共有メニューから `ホーム画面に追加` を選びます。

最初の読み込みと更新時はネット接続が必要です。  
一度読み込んだあとは、通常の学習はオフラインでも使えます。

## GitHub Pages 用の補足

- このプロジェクトには GitHub Pages 用の workflow を同梱しています。
- 静的サイトなので、Node.js サーバーや常時起動のPCは不要です。
- PWA として使うため、公開先は `HTTPS` である必要があります。GitHub Pages はこの条件を満たします。

## 主な機能

- `Study Setup -> Quiz -> Session Score` の3画面フロー
- `4択`、`入力`、`フラッシュ`
- `English -> Japanese`、`Japanese -> English`、`Mixed`
- 開始番号と終了番号による範囲指定
- 正答率、連続正解、苦手語、混同しやすい選択肢の記録
- 履歴表示と学習データの書き出し
- CSV 読み込み
- テーマ変更と背景画像設定
- iPhone のホーム画面追加に対応した PWA 構成

## ローカルで確認したい場合

公開前に PC で確認したいときだけ、簡易サーバーを使えます。

1. `start-local.bat` を実行します。
2. ブラウザで `http://127.0.0.1:4173` を開きます。

## 既定データの再生成

`system.xlsx` を更新したあとに既定データを作り直す場合:

```powershell
python .\tools\extract_system_xlsx.py
```
