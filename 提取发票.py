#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
发票登记表自动录入脚本
========================
用途：读取「发票待整理」文件夹里的电子发票 PDF（带文字层），自动抽取 8 个字段，
     写入发票登记表模板，生成 发票登记表-YYYYMMDD.xlsx（当天日期）。

用法：
    python3 提取发票.py                      # 正常录入：读 发票待整理
    python3 提取发票.py --input-dir <目录>    # 指定要处理的文件夹（测试用）
    python3 提取发票.py --dry-run            # 只抽取并打印，不写 Excel（测试用）

字段口径（与模板一致，8 列）：
    发票类型 / 发票号码 / 开票日期 / 开票单位 / 进项税率 / 进项税额 /
    货物或应税劳务、服务名称 / 金额
    · 金额 = 价税合计（含税总额），取自发票上的中文大写金额
    · 进项税额 = 金额 × 税率 ÷ (1+税率)，免税票为 0
    · 进项税率 = 小数（0.13 / 0.06 / 0.03 / 0.01 …），免税票写「免税」

目录结构约定：脚本放在「发票登记表-模版.xlsx」同级目录，其下有「发票待整理」子目录。

注意：本脚本产出是初稿，录入后请人工核对金额、税额、税率后再对外使用。
"""

import os
import re
import sys
import datetime
from collections import Counter

import openpyxl
import pymupdf
from openpyxl.utils import get_column_letter

# ---- 路径：以脚本所在目录为基准 ----
BASE = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE, "发票待整理")
TEMPLATE = os.path.join(BASE, "发票登记表-模版.xlsx")

HEADERS = ["发票类型", "发票号码", "开票日期", "开票单位",
           "进项税率", "进项税额", "货物或应税劳务、服务名称", "金额"]

# ---- 中文大写金额解析 ----
CN_DIGITS = {"零": 0, "〇": 0, "一": 1, "壹": 1, "二": 2, "贰": 2, "两": 2,
             "三": 3, "叁": 3, "四": 4, "肆": 4, "五": 5, "伍": 5,
             "六": 6, "陆": 6, "七": 7, "柒": 7, "八": 8, "捌": 8,
             "九": 9, "玖": 9}
CN_UNITS = {"拾": 10, "十": 10, "佰": 100, "百": 100, "仟": 1000, "千": 1000,
            "万": 10000, "亿": 100000000}
CN_MONEY_RE = r"[零〇壹贰叁肆伍陆柒捌玖拾佰仟万亿两]+[元圆][零〇壹贰叁肆伍陆柒捌玖拾佰仟万亿角分整正]*"


def cn_int(s):
    """中文数字转整数，如 壹佰捌拾玖 -> 189。"""
    if not s:
        return 0
    total = 0
    section = 0
    num = 0
    for ch in s:
        if ch in CN_DIGITS:
            num = CN_DIGITS[ch]
        elif ch in CN_UNITS:
            unit = CN_UNITS[ch]
            if unit >= 10000:
                section = (section + num) * unit
                total += section
                section = 0
            else:
                section += (num if num else 1) * unit
            num = 0
    return total + section + num


def cn_money(s):
    """中文大写金额转浮点，如 柒拾圆玖角玖分 -> 70.99。"""
    jiao = fen = 0
    m = re.search(r"([零〇壹贰叁肆伍陆柒捌玖]+)角", s)
    if m:
        jiao = cn_int(m.group(1))
    m = re.search(r"([零〇壹贰叁肆伍陆柒捌玖]+)分", s)
    if m:
        fen = cn_int(m.group(1))
    m = re.search(r"(.*?)[元圆]", s)
    yuan = cn_int(m.group(1)) if m else 0
    return round(yuan + jiao * 0.1 + fen * 0.01, 2)


# ---- PDF 文本 ----
def pdf_text(path):
    doc = pymupdf.open(path)
    try:
        return "\n".join(page.get_text() for page in doc)
    finally:
        doc.close()


# ---- 字段抽取 ----
def extract_no(t):
    """发票号码：优先「发票号码：」后 18~20 位；否则全文唯一 20 位数字。"""
    m = re.search(r"发票号码\s*[:：]?\s*(\d{18,20})", t)
    if m:
        return m.group(1)
    m = re.search(r"(?<!\d)(\d{20})(?!\d)", t)
    if m:
        return m.group(1)
    return None


def extract_date(t):
    """开票日期：优先「开票日期」后；否则全文第一个 年月日。"""
    m = re.search(r"开票日期\s*[:：]?\s*(\d{4})年(\d{1,2})月(\d{1,2})日", t)
    if m:
        return "%s年%02d月%02d日" % (m.group(1), int(m.group(2)), int(m.group(3)))
    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", t)
    if m:
        return "%s年%02d月%02d日" % (m.group(1), int(m.group(2)), int(m.group(3)))
    return None


def extract_total(t):
    """金额 = 价税合计，取自中文大写金额（多个取最大）。"""
    matches = list(re.finditer(CN_MONEY_RE, t))
    if matches:
        vals = [cn_money(m.group(0)) for m in matches]
        vals = [v for v in vals if v is not None and v > 0]
        if vals:
            return round(max(vals), 2)
    # 备选：价税合计 之后第一个 ¥/￥ 金额
    m = re.search(r"价税合计[^¥￥]{0,120}?[¥￥]\s*([0-9][0-9,]*(?:\.[0-9]{1,2})?)", t, re.S)
    if m:
        v = m.group(1).replace(",", "")
        try:
            return round(float(v), 2)
        except ValueError:
            return None
    return None


def extract_rate(t):
    """进项税率：免税/不征税 -> '免税'；否则取出现最多的百分比转小数。"""
    mian = bool(re.search(r"免税|不征税", t))
    rates = re.findall(r"(\d{1,2}(?:\.\d+)?)\s*%", t)
    if mian and not rates:
        return "免税"
    if rates:
        c = Counter(rates)
        rate = c.most_common(1)[0][0]
        return round(float(rate) / 100.0, 4)
    if mian:
        return "免税"
    return None


def extract_tax(total, rate):
    """进项税额：免税为 0；否则 金额×税率/(1+税率)。"""
    if rate == "免税":
        return 0.0
    if isinstance(rate, (int, float)) and total is not None:
        return round(total * rate / (1 + rate), 2)
    return None


def extract_seller(t):
    """开票单位（销售方）：按行找「公司」结尾的名字。

    购买方通常是律师事务所（不含「公司」二字），因此只匹配「公司」即自然排除买方。
    若你的购买方也是公司，请自行扩展此处的过滤规则。
    """
    candidates = []
    for line in t.splitlines():
        line = line.strip()
        if not line:
            continue
        if any(w in line for w in ("银行", "地址", "电话", "账号", "开户", "订单号", "备注")):
            continue
        if "公司" in line:
            m = re.search(r"[一-龥A-Za-z0-9（）()·]{2,40}公司", line)
            if m:
                name = m.group(0)
                if name not in candidates:
                    candidates.append(name)
    return candidates[0] if candidates else None


UNIT_WORDS = {"个", "箱", "册", "次", "盒", "台", "套", "张", "桶", "袋", "支",
              "件", "本", "双", "瓶", "份", "条", "只", "卷", "包", "升", "米",
              "组", "把", "盒装", "袋装", "桶装"}
STOP_WORDS = {"免税", "不征税", "无", "默认", "详见", "不含税"}


def extract_goods(t):
    """货物/服务名称：*大类*明细，明细跨行拼接；铁路客票固定为「铁路旅客运输」。"""
    m = re.search(r"\*([^*\n]+?)\*", t)
    if not m:
        if re.search(r"铁路", t) and re.search(r"电子客票", t):
            return "铁路旅客运输"
        return None
    cat = m.group(1).strip()
    rest = t[m.end():]
    parts = []
    for line in rest.split("\n"):
        line = line.strip()
        if not line:
            continue
        if line.startswith("*"):
            break  # 遇到重复出现的货物条目（如京东票表内+合计区各一次），停止
        if any(w in line for w in ("¥", "￥", "%", "订单号", "备注", "下载次数",
                                    "销售方", "地址", "电话", "开户")):
            break
        if re.fullmatch(r"\d+(?:\.\d+)?", line):
            break
        if line in UNIT_WORDS or line in STOP_WORDS:
            break
        parts.append(line)
        if sum(len(p) for p in parts) >= 40:
            break
    detail = "".join(parts).strip()
    if not detail:
        return None
    return "*%s*%s" % (cat, detail[:40])


def extract_type(t):
    """发票类型：铁路电子客票单独一类；其余按「电子发票（XX）」规范化。"""
    if re.search(r"铁路电子客票|铁路（电子客票）|铁路电子发票", t):
        return "铁路电子客票"
    m = re.search(r"电子发票\s*[（(]\s*([^（）()\n]+?)\s*[）)]", t)
    if m:
        return "电子发票（%s）" % m.group(1).strip()
    if re.search(r"电子发票", t):
        return "电子发票"
    return None


def extract_all(path):
    """抽取一张发票的全部字段，返回 (字段dict, 缺失字段列表)。"""
    t = pdf_text(path)
    if not t.strip():
        return None, ["PDF 无文字层（疑似扫描件/图片）"]

    d = {
        "发票类型": extract_type(t),
        "发票号码": extract_no(t),
        "开票日期": extract_date(t),
        "开票单位": extract_seller(t),
        "进项税率": extract_rate(t),
        "进项税额": None,
        "货物或应税劳务、服务名称": extract_goods(t),
        "金额": extract_total(t),
    }
    d["进项税额"] = extract_tax(d["金额"], d["进项税率"])

    missing = [k for k in HEADERS if d.get(k) in (None, "")]
    return d, missing


# ---- Excel 写入 ----
def output_path():
    date = datetime.date.today().strftime("%Y%m%d")
    base = os.path.join(BASE, "发票登记表-%s.xlsx" % date)
    if not os.path.exists(base):
        return base
    i = 2
    while True:
        p = os.path.join(BASE, "发票登记表-%s-%d.xlsx" % (date, i))
        if not os.path.exists(p):
            return p
        i += 1


def write_xlsx(rows):
    from copy import copy
    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb["Sheet1"]

    start = 2
    style_src_row = 2 if ws.max_row >= 2 else None

    for r, row in enumerate(rows):
        excel_row = start + r
        if style_src_row is not None and excel_row > ws.max_row:
            for c in range(1, len(HEADERS) + 1):
                sc = ws.cell(row=style_src_row, column=c)
                dc = ws.cell(row=excel_row, column=c)
                if sc.has_style:
                    dc._style = copy(sc._style)
        for c, key in enumerate(HEADERS, start=1):
            cell = ws.cell(row=excel_row, column=c, value=row.get(key))
            # 进项税率：百分比（0.13→13%）；进项税额、金额：两位小数
            if c == 5:
                cell.number_format = "0%"
            elif c in (6, 8):
                cell.number_format = "0.00"

    _autofit(ws)
    out = output_path()
    wb.save(out)
    return out


def _disp_width(s):
    """显示宽度：中文/全角按 2 个字符宽，其余按 1 个。"""
    w = 0
    for ch in str(s):
        w += 2 if ord(ch) > 0x2E80 else 1
    return w


def _autofit(ws, min_width=8, padding=2):
    """按每列内容（含表头）自动设置列宽，保证不折行。"""
    for col in range(1, len(HEADERS) + 1):
        letter = get_column_letter(col)
        mw = 0
        for row in range(1, ws.max_row + 1):
            cell = ws.cell(row=row, column=col)
            v = cell.value
            if v is None:
                continue
            # 金额/税额列按「两位小数」的显示值计宽（如 189.00 而非 189.0）
            if cell.number_format == "0.00" and isinstance(v, (int, float)):
                s = "%.2f" % v
            else:
                s = str(v)
            mw = max(mw, _disp_width(s))
        ws.column_dimensions[letter].width = max(mw + padding, min_width)


def main():
    args = sys.argv[1:]
    dry_run = "--dry-run" in args
    input_dir = INPUT_DIR
    if "--input-dir" in args:
        i = args.index("--input-dir")
        if i + 1 < len(args):
            input_dir = os.path.abspath(os.path.expanduser(args[i + 1]))

    if not os.path.isdir(input_dir):
        print("❌ 找不到文件夹：%s" % input_dir)
        sys.exit(1)

    pdfs = sorted(
        f for f in os.listdir(input_dir)
        if f.lower().endswith(".pdf") and not f.startswith(".")
    )
    others = [f for f in os.listdir(input_dir)
              if not f.startswith(".") and not f.lower().endswith(".pdf")]

    print("=" * 70)
    print("待处理 PDF：%d 个　（目录：%s）" % (len(pdfs), input_dir))
    if others:
        print("跳过非 PDF 文件 %d 个：%s" % (len(others), "、".join(others)))
    print("=" * 70)

    rows = []
    seen_no = set()
    problems = []
    for f in pdfs:
        path = os.path.join(input_dir, f)
        d, missing = extract_all(path)
        if d is None:
            problems.append((f, missing))
            print("⚠️  跳过  %s —— %s" % (f, "；".join(missing)))
            continue

        no = d.get("发票号码")
        if no and no in seen_no:
            problems.append((f, ["发票号码重复，已跳过"]))
            print("⚠️  跳过  %s —— 发票号码 %s 已存在" % (f, no))
            continue
        if no:
            seen_no.add(no)

        rows.append(d)
        if missing:
            problems.append((f, missing))

        print("\n📄 %s" % f)
        for k in HEADERS:
            v = d.get(k)
            flag = "　" if k not in missing else " ◀缺失"
            print("   %-16s %s%s" % (k, "" if v is None else v, flag))

    print("\n" + "=" * 70)
    print("抽取完成：有效 %d 行，跳过/异常 %d 个" % (len(rows), len(problems)))
    if problems:
        print("需人工核对：")
        for f, ms in problems:
            print("  · %s —— %s" % (f, "；".join(ms)))
    print("提示：进项税额为按 金额×税率 计算；货物名称明细跨行时可能不完整，请一并核对。")
    print("=" * 70)

    if dry_run:
        print("（--dry-run：未写入 Excel）")
        return

    if not rows:
        print("没有可写入的数据，未生成文件。")
        return

    out = write_xlsx(rows)
    print("\n✅ 已生成：%s" % out)
    print("请人工核对金额、税额、税率后再对外使用。")


if __name__ == "__main__":
    main()
