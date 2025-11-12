# PDFにアウトラインを追加する

CSV/Excelのアウトライン表からPDFにアウトライン（しおり）を追加するスクリプトの使い方

依存ライブラリのインストール

  pip install pandas pdfrw reportlab openpyxl
  
  pip install pdfminer.six
  
  pip install pdfrw
  
  pip install reportlab


実行例
事前にPDFに対してExcelかcsvファイルで、class,title,page　列を作成し、その下にclassなら章(class =1に対応）/節(class =2に対応）/column(class =3に対応）で、titleは目次の名前、pageにはページ番号を作成しておく必要があります。（ChatGPTやGeminiで作成すること推奨）

（Excelに「class / title / page」列があり、PDFの論理ページ1が物理21ページ目なら --page_offset 20）

python add_outline_from_excel.py \
  --input_pdf "input.pdf" \
  --outline_file "outline.xlsx" \
  --output_pdf "output.pdf" \
  --page_offset 20


CSVで渡す場合
python add_outline_from_excel.py \
  --input_pdf "in.pdf" \
  --outline_file "outline.csv" \
  --output_pdf "out.pdf"
