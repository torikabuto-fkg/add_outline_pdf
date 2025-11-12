import argparse
import pandas as pd
from pathlib import Path

from pdfrw import PdfReader
from pdfrw.buildxobj import pagexobj
from pdfrw.toreportlab import makerl
from reportlab.pdfgen.canvas import Canvas

# ===== 設定（必要なら編集） =========================================
# クラス名 → ブックマーク階層レベル
CLASS_TO_LEVEL = {
    "章": 1,
    "節": 2,
    "column": 3,
}

# 安全側の正規化（未知クラスは下位レベルに寄せる）
def normalize_level(raw_level: int) -> int:
    if raw_level < 1:
        return 1
    if raw_level > 3:
        return 3
    return raw_level
# ===================================================================


def load_outline_table(path: Path, sheet_name: str = "", page_col="page", class_col="class", title_col="title"):
    ext = path.suffix.lower()
    if ext in [".xlsx", ".xls"]:
        df = pd.read_excel(path, sheet_name=sheet_name if sheet_name else 0)
    elif ext in [".csv"]:
        df = pd.read_csv(path)
    else:
        raise ValueError(f"Unsupported outline file type: {ext}")

    # 必須列の存在確認
    for col in (page_col, class_col, title_col):
        if col not in df.columns:
            raise ValueError(f"'{col}' 列が見つかりませんでした。列名を指定するか、ファイルを確認してください。")

    # 欠損を除去しシンプル化
    df = df[[class_col, title_col, page_col]].dropna()
    # 型をそろえる
    df[page_col] = df[page_col].astype(int)
    df[class_col] = df[class_col].astype(str).str.strip()
    df[title_col] = df[title_col].astype(str).str.strip()

    return df.rename(columns={page_col: "page", class_col: "klass", title_col: "title"})


def to_outlines(df: pd.DataFrame, page_offset: int = 0):
    """
    DataFrame → [(title, page_index(0始まり), level), ...]
    """
    outlines = []
    for _, row in df.iterrows():
        klass = row["klass"]
        title = row["title"]
        logical_page = int(row["page"])

        # クラス → レベル
        level = CLASS_TO_LEVEL.get(klass, 3)
        level = normalize_level(level)

        # PDFの0始まりページindexへ換算
        # 例: 論理ページ=1, offset=+20 → 実物index = 1-1+20 = 20
        page_index = logical_page - 1 + int(page_offset)

        if page_index < 0:
            # 念のための防御
            continue

        outlines.append((title, page_index, level))
    return outlines


def add_outline_with_reportlab_pdfrw(input_pdf: Path, output_pdf: Path, outlines):
    """
    pdfrw + reportlab を使って各ページにブックマークを追加
    """
    pages = PdfReader(str(input_pdf), decompress=False).pages
    c = Canvas(str(output_pdf))

    # 先頭に目次ノード（任意）
    c.bookmarkPage("toc")
    c.addOutlineEntry("目次", "toc", 0)

    added = 0
    errors = 0

    # ページごとに流し込み
    for idx, page in enumerate(pages):
        xobj = pagexobj(page)
        c.setPageSize(tuple(xobj.BBox[2:]))
        c.doForm(makerl(c, xobj))

        # このページに属するアウトラインを抽出
        targets = [o for o in outlines if o[1] == idx]  # (title, page_index, level)
        for i, (title, _, level) in enumerate(targets):
            anchor = f"p{idx}-{i}"
            c.bookmarkPage(anchor)
            try:
                # reportlabのレベル指定は 0,1,2,...（ここでは 1→0, 2→1, 3→2 に写像）
                c.addOutlineEntry(title, anchor, max(0, level - 1))
                added += 1
            except Exception:
                # レベル周りの不整合は最下位レベルに落として再試行
                try:
                    c.addOutlineEntry(title, anchor, 2)
                    added += 1
                except Exception:
                    errors += 1

        c.showPage()

    c.showOutline()
    c.save()
    print(f"✅ アウトライン追加 完了: 成功 {added}, エラー {errors} → 出力: {output_pdf}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input_pdf", required=True, help="入力PDFパス")
    ap.add_argument("--outline_file", required=True, help="CSV/Excelのアウトライン表")
    ap.add_argument("--output_pdf", required=True, help="出力PDFパス")
    ap.add_argument("--page_offset", type=int, default=0, help="論理ページ→物理ページの補正（例: +20）")
    ap.add_argument("--sheet_name", default="", help="Excelの場合のシート名（空なら先頭）")
    ap.add_argument("--page_col", default="page", help="ページ番号列名")
    ap.add_argument("--class_col", default="class", help="クラス列名（章/節/supplement/column）")
    ap.add_argument("--title_col", default="title", help="タイトル列名")
    args = ap.parse_args()

    input_pdf = Path(args.input_pdf)
    outline_file = Path(args.outline_file)
    output_pdf = Path(args.output_pdf)

    df = load_outline_table(
        outline_file,
        sheet_name=args.sheet_name,
        page_col=args.page_col,
        class_col=args.class_col,
        title_col=args.title_col,
    )
    outlines = to_outlines(df, page_offset=args.page_offset)

    if len(outlines) == 0:
        raise RuntimeError("アウトラインが1件もありません。CSV/Excelの内容を確認してください。")

    add_outline_with_reportlab_pdfrw(input_pdf, output_pdf, outlines)


if __name__ == "__main__":
    main()
