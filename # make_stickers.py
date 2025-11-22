# make_stickers.py
# Create 50mm x 25mm sticker PDF from your Excel (robust header matching).

import os
import re
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# ---------- PATHS ----------
EXCEL_PATH = r"D:\Vaishnav\Upcoming Export-190 Pcs.xlsx"
OUTPUT_PDF = r"D:\Vaishnav\stickers.pdf"

# ---------- SIZE ----------
DPI = 300
MM_PER_INCH = 25.4
LABEL_W_MM = 50.0   # 5 cm
LABEL_H_MM = 25.0   # 2.5 cm

def mm_to_px(mm: float) -> int:
    return int(round(mm * DPI / MM_PER_INCH))

LABEL_W = mm_to_px(LABEL_W_MM)
LABEL_H = mm_to_px(LABEL_H_MM)

# ---------- LAYOUT ----------
PAD = mm_to_px(2.4)
LINE_GAP = mm_to_px(1.0)
MADE_IN_UAE_LEFT_OFFSET = mm_to_px(5.0)

# ---------- FONTS ----------
def load_font(size):
    for path in [
        r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\calibri.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/Library/Fonts/Arial.ttf",
    ]:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size=size)
            except:
                pass
    return ImageFont.load_default()

FONT_STYLE = load_font(34)
FONT_TYPE  = load_font(32)
FONT_TEXT  = load_font(30)
FONT_LAST  = load_font(30)

# ---------- HEADER NORMALIZATION ----------
def norm(s: str) -> str:
    s = str(s).replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s.strip()).lower()
    return s

NORM_MAP = {
    "style": "style no",
    "type_num": "type",
    "dia_pcs": "dia pcs",
    "dia_wt": "dia wt",
    "gem_pcs": "gem stone pcs",
    "gem_wt": "gem stone wt",
    "gross_wt": "gross wt",
    "net_wt": "net wt",
    "stock_code": "stock code",
}

# ---------- DRAW HELPERS ----------
def measure(draw, text, font):
    x0, y0, x1, y1 = draw.textbbox((0, 0), text, font=font)
    return (x1 - x0, y1 - y0)

def draw_fit(draw, x, y, text, font, max_w):
    t = "" if text is None else str(text)
    while True:
        w, h = measure(draw, t, font)
        if w <= max_w or len(t) <= 1:
            break
        t = (t[:-2] + "…") if len(t) > 2 else t[:-1]
    draw.text((x, y), t, fill=(0, 0, 0), font=font)
    return y + h

def fmt_int(x):
    try:
        if pd.isna(x):
            return "0"
        return str(int(round(float(x))))
    except:
        return str(x)

def fmt_wt(x):
    try:
        if pd.isna(x):
            return "0.00"
        return f"{float(x):.2f}"
    except:
        s = str(x).strip()
        return s if s else "0.00"

def build_column_index(df):
    return {norm(col): col for col in df.columns}

def get_val(row, col_index, want_norm_name):
    original = col_index.get(want_norm_name)
    if original is None:
        return None
    return row.get(original, None)

def render_label(row_dict, col_index):
    img = Image.new("RGB", (LABEL_W, LABEL_H), "white")
    d = ImageDraw.Draw(img)

    style = get_val(row_dict, col_index, NORM_MAP["style"]) or ""
    type_raw = get_val(row_dict, col_index, NORM_MAP["type_num"])
    dia_pcs = fmt_int(get_val(row_dict, col_index, NORM_MAP["dia_pcs"]))
    dia_wt = fmt_wt(get_val(row_dict, col_index, NORM_MAP["dia_wt"]))
    gem_pcs = fmt_int(get_val(row_dict, col_index, NORM_MAP["gem_pcs"]))
    gem_wt = fmt_wt(get_val(row_dict, col_index, NORM_MAP["gem_wt"]))
    gross_wt = fmt_wt(get_val(row_dict, col_index, NORM_MAP["gross_wt"]))
    net_wt = fmt_wt(get_val(row_dict, col_index, NORM_MAP["net_wt"]))
    stock = get_val(row_dict, col_index, NORM_MAP["stock_code"]) or ""

    type_txt = ""
    if type_raw is not None and str(type_raw).strip() != "":
        try:
            t = int(round(float(type_raw)))
            type_txt = f"{t}K"
        except:
            t = str(type_raw).strip()
            type_txt = t if t.upper().endswith("K") else t + "K"

    x = PAD
    y = PAD
    usable_w = LABEL_W - PAD * 2

    y = draw_fit(d, x, y, str(style), FONT_STYLE, usable_w)
    y += LINE_GAP
    y = draw_fit(d, x, y, type_txt, FONT_TYPE, usable_w)
    y += LINE_GAP
    y = draw_fit(d, x, y, f"Dia: {dia_pcs} / {dia_wt}", FONT_TEXT, usable_w)
    y += LINE_GAP
    y = draw_fit(d, x, y, f"Gem: {gem_pcs} / {gem_wt}", FONT_TEXT, usable_w)
    y += LINE_GAP
    y = draw_fit(d, x, y, f"Gross: {gross_wt} / Net: {net_wt}", FONT_TEXT, usable_w)

    right_text = "MADE IN UAE"
    left_text = str(stock)
    left_w, left_h = measure(d, left_text, FONT_LAST)
    y_bottom = y + LINE_GAP
    max_y_bottom = LABEL_H - PAD - left_h
    if y_bottom > max_y_bottom:
        y_bottom = max_y_bottom
    made_in_uae_x = x + left_w + MADE_IN_UAE_LEFT_OFFSET

    d.text((x, y_bottom), left_text, fill=(0, 0, 0), font=FONT_LAST)
    d.text((made_in_uae_x, y_bottom), right_text, fill=(0, 0, 0), font=FONT_LAST)

    return img

def save_pdf(images, out_path):
    if not images:
        raise ValueError("No images to save.")
    first, rest = images[0], images[1:]
    first.save(
        out_path,
        "PDF",
        resolution=DPI,
        save_all=bool(rest),
        append_images=rest,
    )

def main():
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"Excel not found: {EXCEL_PATH}")

    df = pd.read_excel(EXCEL_PATH)

    # Remove fully blank rows
    df = df.dropna(how='all')

    # Strip all string columns
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].astype(str).str.strip().replace({"": None, "nan": None})

    # Build column index
    col_index = build_column_index(df)
    style_col = col_index.get(NORM_MAP["style"])
    stock_col = col_index.get(NORM_MAP["stock_code"])

    # Keep only rows where both Style and Stock Code are present
    df = df.dropna(subset=[style_col, stock_col])

    # Remove duplicates
    df = df.drop_duplicates(subset=[style_col, stock_col])

    print(f"Rows to process: {len(df)}")  # Should match visible rows

    # Render stickers
    images = [render_label(row.to_dict(), col_index) for _, row in df.iterrows()]
    if not images:
        raise ValueError("No rows to create stickers.")
    os.makedirs(os.path.dirname(OUTPUT_PDF), exist_ok=True)
    save_pdf(images, OUTPUT_PDF)
    print(f"Saved: {OUTPUT_PDF}")

if __name__ == "__main__":
    main()
