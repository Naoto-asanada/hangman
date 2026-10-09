"""pptx を DADS ルールで機械チェックする。

- 文字サイズ 14px 未満（エラー）／14px（注記以外なら要確認）
- 文字と背景のコントラスト 4.5:1 未満（エラー）
- 位置・サイズが 8px グリッドに乗っていない（警告）
- 影の使用（警告）

使い方: python3 check_deck.py deck.pptx
"""
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

EMU_PER_PX = 9525


def lum(hexstr):
    c = [int(hexstr[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    c = [x / 12.92 if x <= 0.03928 else ((x + 0.055) / 1.055) ** 2.4 for x in c]
    return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]


def contrast(a, b):
    la, lb = lum(a), lum(b)
    return (max(la, lb) + 0.05) / (min(la, lb) + 0.05)


def fill_of(shape_or_cell):
    try:
        f = shape_or_cell.fill
        if f.type == 1:  # solid
            return str(f.fore_color.rgb)
    except Exception:
        pass
    return None


def bg_at(slide, shape):
    """shape の背後にある塗りつぶし図形の色（なければ白）。"""
    cx = shape.left + shape.width // 2
    cy = shape.top + shape.height // 2
    color = "FFFFFF"
    for s in slide.shapes:
        if s.shape_id == shape.shape_id:
            break
        if s.left is None or s.has_text_frame and s.text_frame.text.strip():
            continue
        if s.left <= cx <= s.left + s.width and s.top <= cy <= s.top + s.height:
            c = fill_of(s)
            if c:
                color = c
    return color


def check_runs(where, text_frame, bg, issues):
    for p in text_frame.paragraphs:
        for r in p.runs:
            if not r.text.strip():
                continue
            size = r.font.size
            if size is not None:
                px_ = round(size.pt / 0.75, 1)
                if px_ < 14:
                    issues.append(("ERROR", where, f"{px_}px の文字: 「{r.text[:20]}」"))
                elif px_ < 16:
                    issues.append(("CHECK", where, f"14px は出典・注記のみ: 「{r.text[:20]}」"))
            try:
                fg = str(r.font.color.rgb)
            except Exception:
                fg = "000000"
            cr = contrast(fg, bg)
            if cr < 4.5:
                issues.append(("ERROR", where, f"コントラスト {cr:.2f}:1 (#{fg} on #{bg}): 「{r.text[:20]}」"))
        if p.line_spacing is not None and isinstance(p.line_spacing, float) and p.line_spacing < 1.5:
            issues.append(("ERROR", where, f"行間 {p.line_spacing} 倍 < 1.5"))


def main(path):
    prs = Presentation(path)
    issues = []
    for n, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            where = f"slide {n} / {shape.name}"
            for v, name in ((shape.left, "x"), (shape.top, "y"), (shape.width, "w"), (shape.height, "h")):
                if v is not None and round(v / EMU_PER_PX) % 8 not in (0,):
                    issues.append(("WARN", where, f"{name}={round(v / EMU_PER_PX)}px が 8px グリッド外"))
                    break
            if shape._element.find(".//" + qn("a:outerShdw")) is not None:
                issues.append(("WARN", where, "影を使用"))
            if shape.has_text_frame:
                bg = fill_of(shape) or bg_at(slide, shape)
                check_runs(where, shape.text_frame, bg, issues)
            if shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                for row in shape.table.rows:
                    for cell in row.cells:
                        check_runs(where + " (表)", cell.text_frame, fill_of(cell) or "FFFFFF", issues)
    errors = [i for i in issues if i[0] == "ERROR"]
    for level, where, msg in issues:
        print(f"[{level}] {where}: {msg}")
    print(f"\n{len(prs.slides)} slides / ERROR {len(errors)} / "
          f"WARN {sum(i[0] == 'WARN' for i in issues)} / CHECK {sum(i[0] == 'CHECK' for i in issues)}")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
