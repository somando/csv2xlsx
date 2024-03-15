# CSV to XLSX

CSVファイルをExcelファイルに変換するプログラムです。
CSVファイル内の数字は全て文字列として扱います。

本プログラムでは外部との通信を行わないため、安全に変換処理が行えます。

## 実行方法

本プログラムの実行方法は下記の通りです。

### 1. CSVファイルを用意する

CSVファイルを予め用意し、convert.py と同じフォルダーに保存します。

### 2. コマンドプロンプトを開く

convert.pyが保存されているフォルダーを開いた後、アドレスバーにcmdと入力し、コマンドプロンプト（黒画面）を開きます。

詳細は下記の画像・Webサイトを参考にしてください。

<img src="https://image.itmedia.co.jp/ait/articles/1806/14/wi-2000explorer01.png">

Webサイト: [コマンドプロンプトを素早く起動／実行する方法【Windows 11／10】｜＠IT](https://atmarkit.itmedia.co.jp/ait/articles/1806/14/news017.html#explorer)

``` bash
Microsoft Windows [Version 10.0.19045.3324]
(c) Microsoft Corporation. All rights reserved.

C:\Users\Owner\Downloads\csv2xlsx-main> _
```

最後の"_"が点滅しており、この部分に文字を入力できます。
この部分以外には入力できません。

### 3. コマンドを実行する

プログラムの実行にはコマンドプロンプトに下記のコマンドを入力し、実行します。

``` bash
python convert.py <CSVファイル名> <出力Excelファイル名 *任意>
      ^          ^            ^
# " ^ " の部分上部には半角スペースが必要です
```

Pythonファイル名を変更した際には変更後のファイル名を入力してください。
< や > の入力は不要です。


CSVファイル名には最後に".csv"の拡張子がついている必要があります。

出力ファイル名が指定されていない場合はCSVファイル名と同じファイル名が使用されます。

出力ファイル名には".xlsx"の拡張子を指定する必要がありません。（指定した場合も実行可能です。）

実行後はCSVファイルと同じフォルダーにExcelファイルが保存されます。

#### 実行例（1）

CSVファイル名: data.csv, 出力ファイル名: output.xlsx の場合

``` bash
python convert.py data.csv output
```

``` bash
# 出力

data.csv を読み込みました
output.xlsx を保存しました
```

この表示で処理が完了し、プログラムが終了します。ExcelファイルがCSVファイルと同じフォルダーに保存されています。

#### 実行例（2）

CSVファイル名: data.csv, 出力ファイル名: output.xlsx だが、既にoutput.xlsxがある場合

``` bash
python convert.py data.csv output
```

``` bash
# 出力

data.csv を読み込みました
output.xlsx が既に存在します。
上書きしてよろしいですか？(y/N) ... 
```

この場合はファイルを上書きするため、この画面で一度停止し、ユーザーに確認が取られます。

そのまま上書きする場合は "y" または "yes" を入力し上書きを許可します。
これにより最終的に以下の出力となります。

``` bash
# 出力

data.csv を読み込みました
output.xlsx が既に存在します。
上書きしてよろしいですか？(y/N) ... y
output.xlsx を保存しました
```

この表示で処理が完了し、プログラムが終了します。ExcelファイルがCSVファイルと同じフォルダーに保存されています。

#### 実行例（3）

CSVファイル名: data.csv, 出力ファイル名: output.xlsx だが、既にoutput.xlsxがあり、このExcelファイルを開いている場合

``` bash
python convert.py data.csv output
```

``` bash
# 出力

data.csv を読み込みました
output.xlsx が既に存在します。
上書きしてよろしいですか？(y/N) ... y
output.xlsx が開かれているため上書きできません。
Excelファイルを閉じるか、出力するファイル名を変更してください。
```

実行例2と同様に、この画面で一度停止し、ユーザーに確認が取られます。
この例ではユーザーが上書きを許可したという状況です。

しかし、Excel起動時は上書き保存が許可されていないため、処理が行われずプログラムが終了します。

この表示で処理が失敗し、プログラムが終了します。Excelファイルは保存されません。

## 初期設定

下記の手順に沿って初期設定を行ってください。
必要環境は下記の通りです。

### 必要環境

- Pythonが実行できる環境
- 変換するCSVファイル
- Excel 2007以降

### 1. GitHubからファイルを全てダウンロードする

[ここ](https://github.com/somando/csv2xlsx/archive/refs/heads/main.zip)をクリックしてZIPファイルをダウンロードします。

### 2. ローカルPCでZIPファイルを展開する

ファイルは以下の4つがダウンロードされます。

- convert.py
- requirements.txt
- .gitignore
- README.md

実運用で使用するファイルはconvert.pyのみです。
初期設定時にはrequirements.txtも使用します。

初期設定後は他のファイルの削除が可能です。

convert.pyはファイル名の変更及び他のフォルダーへ移動が可能です。

### 3. コマンドプロンプトを開く

convert.pyが保存されているフォルダーを開いた後、アドレスバーにcmdと入力し、コマンドプロンプト（黒画面）を開きます。

詳細は下記の画像・Webサイトを参考にしてください。

<img src="https://image.itmedia.co.jp/ait/articles/1806/14/wi-2000explorer01.png">

Webサイト: [コマンドプロンプトを素早く起動／実行する方法【Windows 11／10】｜＠IT](https://atmarkit.itmedia.co.jp/ait/articles/1806/14/news017.html#explorer)

``` bash
Microsoft Windows [Version 10.0.19045.3324]
(c) Microsoft Corporation. All rights reserved.

C:\Users\Owner\Downloads\csv2xlsx-main> _
```

最後の"_"が点滅しており、この部分に文字を入力できます。
この部分以外には入力できません。

### 4. Pythonがインストールされているか確認する

コマンドプロンプトに以下を入力しEnterキーを押します。

``` bash
python
```

と実行し、出力を確認します。

#### Pythonが既にインストールされている場合

Pythonが既にダウンロードされている場合は以下のような出力が表示されます。

``` bash
Python 3.12.2 (tags/v3.12.2:6abddd9, Feb  6 2024, 21:26:36) [MSC v.1937 64 bit (AMD64)] on win32
Type "help", "copyright", "credits" or "license" for more information.
>>>
```

この場合はPythonのインストールは不要です。以下のコマンドを入力後、Enterを押しPythonを一度終了します。

#### Pythonがインストールされていない場合

Microsoft Storeが自動起動するはずです。
画面左側の青い入手またはインストールボタンを押してインストールを待ちます。
（Microsoft IDでのログインが必要な場合があります。）

Webサイト: [Python 3.12｜Microsoft Store](https://www.microsoft.com/store/productId/9NCVDN91XZQP?ocid=pdpshare)

Microsoft Storeが自動で起動しない場合は下記リンクよりダウンロードを確認できます。
（手順が多く複雑なため、Microsoft Storeが起動した際はMicrosoft Storeの利用をお勧めします）

Webサイト: [Pythonの開発環境を用意しよう！（Windows）｜Progate](https://prog-8.com/docs/python-env-win)

※ Microsoft Storeを通じないインストールの場合、再起動が必要な場合があります。

インストールが完了した場合は再度上記のコードを実行し、出力を確認します。上記のような出力があればインストール完了です。

### 5. 実行に必要なライブラリをインストールする

下記のコマンドをコマンドプロンプトに入力し、実行します。

``` bash
pip install -r requirements.txt
```

このファイルには変換に必要なPythonライブラリがインストールされます。

ダウンロード後は requirements.txt の削除が可能です。

__以上で初期設定が完了です。__
