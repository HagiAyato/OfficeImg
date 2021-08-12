# OfficeImg
Officeのファイル(docx/xlsx/pptx)から画像を抜き出すコンソールアプリ

<h2>インストール方法</h2>
<ol type="1">
  <li><a href="https://github.com/HagiAyato/OfficeImg/releases/">リリースページ</a>
  から"Release_OfficeImg.zip"をダウンロードする</li>
  <li>ダウンロードした"Release_OfficeImg.zip"を、任意のフォルダに解凍する</li>
</ol>
<h4>必須環境</h4>
<ul>
  <li>Framework: .NET Framework 4.6.1</li>
</ul>
<h2>使い方・仕様</h2>

- ドラッグアンドドロップをしたOfficeファイルについて、使用画像を取り出す
- 複数ファイルのドラッグアンドドロップにも対応
- 取り出した画像は、ドラッグアンドドロップしたファイルと同じ階層に置く
- 対応拡張子は下記の通り
 - Word:".docx",".docm",".dotx",".dotm"
 - Power Point:".pptx",".pptm",".ppsx",".ppsm",".potx",".potm"
 - Excel:".xlsx",".xlsm",".xltx",".xltm"
 - Open Document:".ods",".odt",".odp"

<h2>処理の流れ</h2>

1. ファイルの存在チェック
1. ファイルの拡張子チェック
1. ファイルを解凍
1. 解凍したフォルダから画像ファイルを取り出す
1. 解凍したフォルダ削除

<h2>補足情報</h2>
<h4>開発環境</h4>
<ul>
  <li>OS: Windows10</li>
  <li>ソフト: Microsoft Visual Studio Community 2017</li>
</ul>
<h4>使用ライブラリ</h4>
特になし
<h4>著作権</h4>
Copyright (C) 2021 HA Works
<h4>問い合わせ・規約</h4>
商用以外の利用は無制限・無料で可能です。<br />
商用利用について、もしくは問題発生時などは下記アドレスにお問い合わせください。
haworks.eng☆gmail.com※☆→@
