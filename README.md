# Excel → PDF 変換

複数のExcelファイルを選び、**表示されている各シートを1シート＝1PDF**で保存するWindows向け社内ツールです。画面は「ファイルを選ぶ → 保存先を選ぶ → 変換する」の3ステップだけに絞っています。

## 主な仕様

- `.xlsx` / `.xls` / `.xlsm` / `.xlsb` に対応
- 複数のExcelファイルをまとめて選択
- 表示シートだけを1シートずつPDF化（非表示シートは対象外）
- Excelの印刷範囲・ページ設定を使ってPDF化
- PDF名は `Excelファイル名_シート名.pdf`
- 同名PDFがある場合は `_2`、`_3` を付け、既存ファイルを上書きしない
- 元のExcelは読み取り専用で開き、変更・保存しない
- 進捗、完了件数、失敗内容を日本語で表示
- 完了後に保存先を開ける
- 専用のデスクトップアイコン
- 変換時間を画面に表示し、詳細な処理時間を診断ログへ記録
- ネットワークプリンターへ接続せず、WindowsのローカルPDF機能でページ配置を計算

## 利用条件

- Windows 10 / 11
- デスクトップ版 Microsoft Excel
- Windows標準の「Microsoft Print to PDF」機能
- ソースから起動する場合のみ Python 3.11以降

配布用exeにはPythonが含まれるため、利用者のPCにPythonを入れる必要はありません。

## ソースから起動する

PowerShellでプロジェクトのフォルダを開き、次を実行します。

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python run_app.py
```

## 使い方

1. 「Excelファイルを選ぶ」を押し、変換するファイルを選びます。
2. 「保存先を選ぶ」を押し、PDFを入れるフォルダを選びます。
3. 「PDFに変換する」を押します。
4. 完了後、「保存先を開く」から結果を確認します。

変換中はExcelを操作しないでください。変換できない場合は、対象ファイルをExcelで閉じてから、もう一度お試しください。

### 「プリンターの接続を待っています」と表示される場合

バージョン1.2.0以降では、変換がすべて終わるまでローカルの「Microsoft Print to PDF」を利用し、ネットワークプリンターへの接続待ちを回避します。Excel側にもプリンター名とポートを明示的に指定します。実際の印刷は行わず、処理終了時にはWindowsの既定プリンターを元の設定へ戻します。

「Microsoft Print to PDF」が無効なPCでは、Windowsの「Windowsの機能の有効化または無効化」から有効にするか、社内IT担当者へご相談ください。

### 変換に時間がかかる場合

バージョン1.1.0以降では、PDF変換に不要なExcelの画面更新、自動計算、イベント、マクロ処理を停止して高速化しています。シートごとのPDF出力はExcel自身の機能を使用するため、シート数、数式、画像、印刷ページ数によって時間が変わります。

処理時間の詳細は、次のファイルに自動記録されます。個人情報やセルの内容は記録されません。

```text
%LOCALAPPDATA%\ExcelPDFConverter\conversion.log
```

調査が必要な場合は、このログファイルを担当者へ渡してください。

## Windows用exeを作る

WindowsのPowerShellで次を実行します。

```powershell
.\build.ps1
```

テスト後、`dist\ExcelPDFConverter_v1.2.0.exe` が作られます。PyInstaller設定で `console=False` にしているため、完成版exeでは黒いコンソール画面は表示されません。

MacからWindows用exeを直接作ることはできません。Windows PCまたはGitHub ActionsのWindows環境でビルドしてください。

### GitHub Actionsで作る場合

このリポジトリをGitHubへ登録すると、「Actions」画面の「Windows版exeを作成」から手動実行できます。処理完了後、画面下部の `ExcelPDFConverter-Windows` からexeを取得できます。`v1.0.0` のようなタグをpushした場合も自動でビルドされます。

## 構成

```text
run_app.py                         起動ファイル
excel_pdf_converter/app.py        GUI（画面・操作）
excel_pdf_converter/converter.py  Excel変換処理
excel_pdf_converter/naming.py     安全なPDF名の生成
ExcelPDFConverter.spec            exe化設定
build.ps1                         テスト・ビルド
tests/                            自動テスト
```

## 配布前の確認

個人情報を含まないテスト用Excelで、次を確認してください。

- 複数ファイルを選べる
- 表示シートだけがPDFになる
- 日本語のファイル名・シート名で保存できる
- 印刷範囲や改ページが意図どおり反映される
- 一部失敗しても、ほかのファイルの処理が続く
- 会社PCのセキュリティ設定でexeが許可されている

未署名の自作exeは、Windows SmartScreenや会社の管理ルールで警告・ブロックされる場合があります。セキュリティ設定を無理に変更せず、必要に応じて社内IT担当者へ確認してください。
