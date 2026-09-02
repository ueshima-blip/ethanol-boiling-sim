"""進度予定表システム_鳴海.xlsm から記入済みデータを消し、配布できる状態に戻す。

VBA (xl/vbaProject.bin) とシート構成には一切触らない。シートを消すと VBA の
ドキュメントモジュール (Sheet4〜Sheet19) と対応が取れなくなるため、週シートは
残したまま中身だけ「週案_テンプレート」と同じ状態に戻す。
"""
import re
import shutil
import sys
import zipfile
from pathlib import Path

SRC = Path(sys.argv[1])
DST = Path(sys.argv[2])

TEMPLATE = "xl/worksheets/sheet3.xml"                  # 週案_テンプレート
WEEK_SHEETS = [f"xl/worksheets/sheet{i}.xml" for i in range(4, 20)]  # 第01〜16週
CONFIG = "xl/worksheets/sheet1.xml"                    # 基本設定
TIMETABLE = "xl/worksheets/sheet2.xml"                 # 時間割マスター

# 全データ消去（共通版 VBA）と同じ範囲
CONFIG_RANGES = ["B2:B4", "B6:B14", "B16:B20", "C16:C17", "B23:B35", "B38:B43"]
TIMETABLE_RANGES = ["B3:F9"]

# 自己終了 <c .../> と <c ...>...</c> を取り違えないよう、空要素を先に試す
CELL_RE = re.compile(r'<c\b[^>]*?/>|<c\b[^>]*?>.*?</c>', re.S)


def cell_ref(tag):
    m = re.match(r'<c\b[^>]*?\sr="([A-Z]+\d+)"', tag)
    return m.group(1) if m else None


def expand(ranges):
    out = set()
    for rng in ranges:
        (c1, r1), (c2, r2) = (re.match(r"([A-Z]+)(\d+)", p).groups() for p in rng.split(":"))
        for col in range(ord(c1), ord(c2) + 1):
            for row in range(int(r1), int(r2) + 1):
                out.add(f"{chr(col)}{row}")
    return out


def block(xml, tag):
    """<tag>...</tag>（空要素 <tag/> も）をそのまま切り出す。"""
    m = re.search(rf"<{tag}(?:\s[^>]*)?/>", xml)
    if m:
        return m.group(0)
    m = re.search(rf"<{tag}(?:\s[^>]*)?>.*?</{tag}>", xml, re.S)
    return m.group(0) if m else None


def clear_cells(xml, refs):
    """指定セルを ClearContents 相当（書式だけ残す空セル）にする。"""
    def sub(m):
        tag = m.group(0)
        ref = cell_ref(tag)
        if ref not in refs:
            return tag
        style = re.search(r'\ss="\d+"', tag)
        return f'<c r="{ref}"{style.group(0) if style else ""}/>'
    return CELL_RE.sub(sub, xml)


def strip_formula_cache(xml):
    """数式セルのキャッシュ値 <v> を捨てる。

    キャッシュには入力を消す前の計算結果（クラス名など）が残るため、
    数式だけ残して値を落とし、Excel に開いたとき計算し直させる。
    """
    def sub(m):
        tag = m.group(0)
        if '<f' not in tag:
            return tag
        tag = re.sub(r'<v>.*?</v>', '', tag, flags=re.S)
        tag = re.sub(r'\st="(?:str|s|e|b)"', '', tag, count=1)
        return tag
    return CELL_RE.sub(sub, xml)


def main():
    zin = zipfile.ZipFile(SRC)
    parts = {n: zin.read(n) for n in zin.namelist()}
    order = list(zin.namelist())
    zin.close()

    tpl = parts[TEMPLATE].decode("utf-8")
    tpl_data = block(tpl, "sheetData")
    tpl_dim = block(tpl, "dimension")

    # 1) 週シートの中身をテンプレートと同じ状態に戻す
    for name in WEEK_SHEETS:
        xml = parts[name].decode("utf-8")
        cur_data = block(xml, "sheetData")
        xml = xml.replace(cur_data, tpl_data, 1)
        cur_dim = block(xml, "dimension")
        if cur_dim and tpl_dim:
            xml = xml.replace(cur_dim, tpl_dim, 1)
        parts[name] = xml.encode("utf-8")

    # 2) 基本設定・時間割マスターの入力欄を消す
    parts[CONFIG] = clear_cells(parts[CONFIG].decode("utf-8"), expand(CONFIG_RANGES)).encode("utf-8")
    parts[TIMETABLE] = clear_cells(parts[TIMETABLE].decode("utf-8"), expand(TIMETABLE_RANGES)).encode("utf-8")

    # 3) 数式のキャッシュ値を全ワークシートから捨てる
    sheet_parts = [n for n in parts if re.match(r"^xl/worksheets/sheet\d+\.xml$", n)]
    for name in sheet_parts:
        parts[name] = strip_formula_cache(parts[name].decode("utf-8")).encode("utf-8")
    print(f"  数式キャッシュを削除: {len(sheet_parts)} シート")

    # Excel が開いたとき必ず全再計算するようにする
    wb = parts["xl/workbook.xml"].decode("utf-8")
    if "<calcPr" in wb:
        wb = re.sub(r"<calcPr([^>]*?)/>", lambda m: f'<calcPr{m.group(1)} fullCalcOnLoad="1"/>', wb, count=1)
    else:
        wb = wb.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    parts["xl/workbook.xml"] = wb.encode("utf-8")
    print("  workbook.xml に fullCalcOnLoad を設定")

    # 4) 参照されなくなった共有文字列を捨て、インデックスを振り直す
    ss = parts["xl/sharedStrings.xml"].decode("utf-8")
    sis = re.findall(r"<si>.*?</si>", ss, re.S)
    used = set()
    sheet_names = [n for n in parts if re.match(r"^xl/worksheets/sheet\d+\.xml$", n)]
    for name in sheet_names:
        for m in CELL_RE.finditer(parts[name].decode("utf-8")):
            if 't="s"' in m.group(0):
                v = re.search(r"<v>(\d+)</v>", m.group(0))
                if v:
                    used.add(int(v.group(1)))
    keep = sorted(used)
    remap = {old: new for new, old in enumerate(keep)}
    for name in sheet_names:
        xml = parts[name].decode("utf-8")

        def sub(m, _remap=remap):
            tag = m.group(0)
            if 't="s"' not in tag:
                return tag
            v = re.search(r"<v>(\d+)</v>", tag)
            if not v:
                return tag
            return tag.replace(f"<v>{v.group(1)}</v>", f"<v>{_remap[int(v.group(1))]}</v>", 1)

        parts[name] = CELL_RE.sub(sub, xml).encode("utf-8")

    total = sum(
        1
        for name in sheet_names
        for m in CELL_RE.finditer(parts[name].decode("utf-8"))
        if 't="s"' in m.group(0)
    )
    sst_open = re.search(r"<sst[^>]*>", ss).group(0)
    sst_open = re.sub(r'count="\d+"', f'count="{total}"', sst_open)
    sst_open = re.sub(r'uniqueCount="\d+"', f'uniqueCount="{len(keep)}"', sst_open)
    decl = ss[: ss.index("<sst")]
    parts["xl/sharedStrings.xml"] = (
        decl + sst_open + "".join(sis[i] for i in keep) + "</sst>"
    ).encode("utf-8")
    print(f"  共有文字列: {len(sis)} → {len(keep)} 件（{len(sis) - len(keep)} 件を削除）")

    # 5) 陳腐化した calcChain を捨てる（Excel が開いたとき作り直す）
    if "xl/calcChain.xml" in parts:
        del parts["xl/calcChain.xml"]
        order.remove("xl/calcChain.xml")
        ct = parts["[Content_Types].xml"].decode("utf-8")
        ct = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", ct)
        parts["[Content_Types].xml"] = ct.encode("utf-8")
        rels = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
        rels = re.sub(r"<Relationship[^>]*calcChain\.xml[^>]*/>", "", rels)
        parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
        print("  calcChain.xml を削除")

    # 6) 作成者名などのドキュメントプロパティを消す
    core = parts["docProps/core.xml"].decode("utf-8")
    core = re.sub(r"<cp:lastModifiedBy>.*?</cp:lastModifiedBy>", "<cp:lastModifiedBy></cp:lastModifiedBy>", core, flags=re.S)
    core = re.sub(r"<dc:creator>.*?</dc:creator>", "<dc:creator></dc:creator>", core, flags=re.S)
    core = re.sub(r"<cp:lastPrinted>.*?</cp:lastPrinted>", "", core, flags=re.S)
    parts["docProps/core.xml"] = core.encode("utf-8")
    print("  docProps/core.xml: creator / lastModifiedBy / lastPrinted を消去")

    tmp = DST.with_suffix(".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as zout:
        for name in order:
            zout.writestr(name, parts[name])
    shutil.move(tmp, DST)
    print(f"  → {DST}")


main()
