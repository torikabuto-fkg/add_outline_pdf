# PDFにアウトラインを追加する

セットアップ:
pip install pdfminer.six
pip install pdfrw
pip install reportlab
pip install pandas pdfrw reportlab openpyxl



CSV/Excelのアウトライン表からPDFにアウトライン（しおり）を追加するスクリプトの使い方

依存ライブラリのインストール
pip install pandas pdfrw reportlab openpyxl


実行例
（Excelに「class / title / page」列があり、PDFの論理ページ1が物理21ページ目なら --page_offset 20）

python add_outline_from_table.py \
  --input_pdf "病気が見える_小児科_OCR.pdf" \
  --outline_file "病気が見える_小児科_アウトライン一覧.xlsx" \
  --output_pdf "病気が見える_小児科_アウトライン付き.pdf" \
  --page_offset 20


CSVで渡す場合
python add_outline_from_table.py \
  --input_pdf "in.pdf" \
  --outline_file "outline.csv" \
  --output_pdf "out.pdf"
