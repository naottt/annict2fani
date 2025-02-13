# annict2fanicsv.ps1
アニメ視聴記録、感想シェアサイトAnnict.comに自分が書いた感想を全て取得し、Fani通調査票形式のCSVを出力します。

# Features
アニメ視聴時にAnnictに感想や評価を残しておけば、Fani通レビュー調査票の元ネタを簡単作成できます。
取得するのは記録(感想)、評価、各回の記録(感想)、番組メモです。
([Fani通](https://x.com/fanitu)は年2回発行されるアニメ総合感想同人誌です。)

Fani通調査票(Excel)には付属 AnnictCSV2FaniExcel.xlsm を使ってコピーします。

# Requirement
* Windows Powershell (5.1以降)

Windows標準機能で動きます。片仮名/平仮名変換のためVB.NETのStrConvを内部で利用しています。動作確認はWindows11(24H2)で行っています。

# Installation
Annictの個人用アクセストークンを以下から新規作成します。
[Annict.com/settings/apps](https://annict.com/settings/apps)

Windows上でユーザー環境変数 ANNICT_ACCESS_TOKEN を作成し、上記で作成したアクセストークンをセットします。

# Usage
お好みのフォルダでPowershellターミナルを開きスクリプトを実行すればOK。デスクトップ上に "annict_personal_review_YYYYMMDD_HHMM.csv" が出力されます。(スクリプト実行権限は設定済の前提)
```
.\annict2fanicsv.ps1
```
もしくは、annict2fanicsv.ps1のショートカットを作成し、ショートカットを右クリックしプロパティ(R)＞リンク(T)先を以下に変更します。その後ショートカットを実行して下さい。
```
powershell -ExecutionPolicy RemoteSigned -File annict2fanicsv.ps1
```
"-dump"を付けて実行するとAnnict APIから取得したJSONファイルをデスクトップ上へファイル保存します。(エラー調査用)
```
.\annict2fanicsv.ps1 -dump
```

# Note
* 各話コメントはレビュー本文末尾に"第1話:コメント"形式で追記します
* レビューが複数回がある場合、評価が上書きされない様に別行として出力しています
* レビューが無く、各回コメントもしくは番組メモのみの場合は追記します。その場合評価等は入りません。備考欄に注釈をいれています
* 記入日時が取得できた場合U列に出力します(Fani通調査票に項目はありません)
* V列に番組名(かな)を出力します(Fani通調査票に項目はありません)
* CSVはシーズン、番組名かな順にソートしています
* Annictの視聴状況 "見たい/見てる/見た/一時中断/視聴中止" をFani通の"(見たい)/視聴途中/視聴済/視聴途中/途中で切"に割当ています。"見たい"はFani通の視聴状況にはありません。またFani通の"繰り返し/初回切"ステータスは無いので割り当てていません
* 評価はAnnictが4段階、Fani通は5段階なのでFani通の5～2に割り当てています。1は設定していません
* Fani通調査票(Excel)にはAnnictCSV2FaniExcel.xlsmを使って転記します

# Author
* 作成者 @naottt

# License
"annict2fanicsv.ps1" is under [MIT license](https://en.wikipedia.org/wiki/MIT_License).
