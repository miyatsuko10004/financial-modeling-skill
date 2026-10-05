#!/usr/bin/env python3
"""
new_model.py: 財務モデリングExcel自動構築ジェネレータ

規約（modeling-rules.md）に100%準拠した財務モデルExcelを、
任意の前提・実績データ（または標準プロ向けデフォルト）からステップバイステップで構築する。

機能:
1. Type-01: 単一シート完結型 3表連動モデル（Single-Sheet 3-Statement Model）
2. Type-02: 多シート本格オペレーティングモデル（Multi-Sheet 3-Statement Model）
3. Type-03: DCF バリュエーションモデル（DCF Valuation Model）

使い方:
  python3 scripts/new_model.py --type single --title "グローバル製造株式会社 財務モデル" -o output.xlsx
  python3 scripts/new_model.py --type dcf --title "B2B SaaS 事業計画モデル" -o dcf_model.xlsx
  python3 scripts/new_model.py --config config.json -o custom_model.xlsx
"""

import sys
import os
import argparse
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- カラーパレット定義（規約§1準拠） ---
COLOR_BLUE_TEXT = "0000FF"      # ハードコード・前提数値（青字）
COLOR_BLACK_TEXT = "000000"     # 計算式・合計（黒字）
COLOR_GREEN_TEXT = "008000"     # 別シート参照（緑字）
COLOR_WHITE_TEXT = "FFFFFF"     # 白抜き見出し文字

COLOR_NAVY_HEADER = "1B365D"    # セクション見出し背景（濃紺）
COLOR_LIGHT_ORANGE = "FFF2CC"   # 入力可能セル背景（薄オレンジ）
COLOR_LIGHT_GRAY = "F2F2F2"     # 集計行・ヘッダー背景（薄グレー）
COLOR_OK_GREEN = "E2EFDA"       # チェックOK背景（薄緑）

FONT_FAMILY_JP = "Meiryo UI"

# 境界線スタイル
THIN_SIDE = Side(border_style="thin", color="D9D9D9")
DOUBLE_BOTTOM_SIDE = Side(border_style="double", color="000000")
TOP_THIN_SIDE = Side(border_style="thin", color="000000")

BORDER_SUBTOTAL = Border(top=TOP_THIN_SIDE, bottom=TOP_THIN_SIDE)
BORDER_TOTAL = Border(top=TOP_THIN_SIDE, bottom=DOUBLE_BOTTOM_SIDE)
BORDER_GRID = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

def apply_header_style(cell, text):
    cell.value = text
    cell.font = Font(name=FONT_FAMILY_JP, size=10, bold=True, color=COLOR_WHITE_TEXT)
    cell.fill = PatternFill(fill_type="solid", start_color=COLOR_NAVY_HEADER, end_color=COLOR_NAVY_HEADER)
    cell.alignment = Alignment(horizontal="center", vertical="center")

def apply_year_header(cell, text):
    cell.value = text
    cell.font = Font(name=FONT_FAMILY_JP, size=10, bold=True, color=COLOR_BLACK_TEXT)
    cell.fill = PatternFill(fill_type="solid", start_color=COLOR_LIGHT_GRAY, end_color=COLOR_LIGHT_GRAY)
    cell.alignment = Alignment(horizontal="right", vertical="center")

def apply_input_cell(cell, value, is_percent=False):
    cell.value = value
    cell.font = Font(name=FONT_FAMILY_JP, size=10, color=COLOR_BLUE_TEXT) # 青字
    cell.fill = PatternFill(fill_type="solid", start_color=COLOR_LIGHT_ORANGE, end_color=COLOR_LIGHT_ORANGE) # オレンジ背景
    cell.alignment = Alignment(horizontal="right", vertical="center")
    if is_percent:
        cell.number_format = "0.0%"
    else:
        cell.number_format = "#,##0"

def apply_formula_cell(cell, formula, is_percent=False, is_subtotal=False, is_total=False):
    cell.value = formula
    cell.font = Font(name=FONT_FAMILY_JP, size=10, bold=(is_subtotal or is_total), color=COLOR_BLACK_TEXT) # 黒字
    cell.alignment = Alignment(horizontal="right", vertical="center")
    if is_percent:
        cell.number_format = "0.0%"
    else:
        cell.number_format = "#,##0"
        
    if is_total:
        cell.border = BORDER_TOTAL
    elif is_subtotal:
        cell.border = BORDER_SUBTOTAL

def build_single_sheet_model(wb, title="財務業績予測・3表連動モデル", params=None):
    ws = wb.active
    ws.title = "業績予想"
    ws.views.sheetView[0].showGridLines = True

    p = params or {}
    title = p.get("title", title)
    unit = p.get("unit", "（単位：百万円）")
    years = p.get("years", ["2022A", "2023E", "2024E", "2025E", "2026E"])
    cols = ["C", "D", "E", "F", "G"]

    # カラム幅の設定
    ws.column_dimensions['A'].width = 18 # エレベーターコラム
    ws.column_dimensions['B'].width = 30 # 項目名
    ws.column_dimensions['C'].width = 14 # 実績
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 14
    ws.column_dimensions['F'].width = 14
    ws.column_dimensions['G'].width = 14

    # タイトル行
    ws['B1'].value = title
    ws['B1'].font = Font(name=FONT_FAMILY_JP, size=14, bold=True, color="1B365D")
    ws['G1'].value = unit
    ws['G1'].font = Font(name=FONT_FAMILY_JP, size=9, color="595959")
    ws['G1'].alignment = Alignment(horizontal="right")

    # ==========================================
    # （１）前提条件
    # ==========================================
    apply_header_style(ws['A2'], "（１）前提")
    apply_header_style(ws['B2'], "前提条件パラメータ")
    for col, yr in zip(cols, years):
        apply_year_header(ws[f'{col}2'], yr)

    assump_config = p.get("assumptions", {})
    rev_g = assump_config.get("revenue_growth", [0.05, 0.05, 0.04, 0.03])
    cogs_r = assump_config.get("cogs_ratio", [0.65, 0.64, 0.63, 0.62, 0.62])
    sga_r = assump_config.get("sga_ratio", [0.20, 0.20, 0.19, 0.19, 0.19])
    tax_r = assump_config.get("tax_rate", [0.30, 0.30, 0.30, 0.30, 0.30])
    capex_vals = assump_config.get("capex", [-600, -600, -500, -500])
    depr_r = assump_config.get("depr_rate", [0.10, 0.10, 0.10, 0.10])
    ar_d = assump_config.get("ar_days", [60.0, 60.0, 60.0, 60.0])
    inv_d = assump_config.get("inv_days", [45.0, 45.0, 45.0, 45.0])
    ap_d = assump_config.get("ap_days", [40.0, 40.0, 40.0, 40.0])
    div_p = assump_config.get("div_payout", [0.30, 0.30, 0.30, 0.30])
    debt_rep = assump_config.get("debt_repay", [-300, -300, -300, -300])
    int_r = assump_config.get("interest_rate", [0.02, 0.02, 0.02, 0.02])

    assumptions = [
        ("売上成長率", None, rev_g[0], rev_g[1], rev_g[2], rev_g[3], True),       # Row 3
        ("売上原価率", cogs_r[0], cogs_r[1], cogs_r[2], cogs_r[3], cogs_r[4], True), # Row 4
        ("販管費率（対売上）", sga_r[0], sga_r[1], sga_r[2], sga_r[3], sga_r[4], True), # Row 5
        ("実効税率", tax_r[0], tax_r[1], tax_r[2], tax_r[3], tax_r[4], True),   # Row 6
        ("設備投資額（Capex）", None, capex_vals[0], capex_vals[1], capex_vals[2], capex_vals[3], False), # Row 7
        ("減価償却率（期首PP&E比）", None, depr_r[0], depr_r[1], depr_r[2], depr_r[3], True), # Row 8
        ("売掛金回転日数（日）", None, ar_d[0], ar_d[1], ar_d[2], ar_d[3], False), # Row 9
        ("棚卸資産回転日数（日）", None, inv_d[0], inv_d[1], inv_d[2], inv_d[3], False), # Row 10
        ("買掛金回転日数（日）", None, ap_d[0], ap_d[1], ap_d[2], ap_d[3], False), # Row 11
        ("配当性向", None, div_p[0], div_p[1], div_p[2], div_p[3], True),        # Row 12
        ("借入金約定返済額", None, debt_rep[0], debt_rep[1], debt_rep[2], debt_rep[3], False), # Row 13
        ("借入金利", None, int_r[0], int_r[1], int_r[2], int_r[3], True)         # Row 14
    ]

    for idx, (label, c_val, d_val, e_val, f_val, g_val, is_pct) in enumerate(assumptions, start=3):
        ws[f'B{idx}'].value = label
        ws[f'B{idx}'].font = Font(name=FONT_FAMILY_JP, size=9)
        vals = [c_val, d_val, e_val, f_val, g_val]
        for col, v in zip(cols, vals):
            if v is not None:
                apply_input_cell(ws[f'{col}{idx}'], v, is_percent=is_pct)

    # ==========================================
    # （２）損益計算書（PL）
    # ==========================================
    apply_header_style(ws['A16'], "（２）損益計算書")
    apply_header_style(ws['B16'], "損益計算書（PL）")
    for col, yr in zip(cols, years):
        apply_year_header(ws[f'{col}16'], yr)

    base_cfg = p.get("base_year", {})
    b_rev = base_cfg.get("revenue", 10000)
    b_ppe = base_cfg.get("ppe", 4000)
    b_re = base_cfg.get("retained_earnings", 2500)
    b_ar = base_cfg.get("ar", 1650)
    b_inv = base_cfg.get("inventory", 800)
    b_ap = base_cfg.get("ap", 700)
    b_debt = base_cfg.get("debt", 3000)
    b_int = base_cfg.get("interest", 60)
    b_cash = base_cfg.get("cash", 1500)
    b_cap = base_cfg.get("capital", (b_cash + b_ar + b_inv + b_ppe) - (b_ap + b_debt + b_re))

    # 売上高 (Row 17)
    ws['B17'].value = "売上高"
    ws['B17'].font = Font(name=FONT_FAMILY_JP, size=10, bold=True)
    apply_input_cell(ws['C17'], b_rev) # 基準年実績
    for i, col in enumerate(["D", "E", "F", "G"]):
        prev_col = cols[i]
        apply_formula_cell(ws[f'{col}17'], f"={prev_col}17*(1+{col}3)")

    # 売上原価 (Row 18) - 規約: 費用はマイナス
    ws['B18'].value = "売上原価"
    ws['B18'].font = Font(name=FONT_FAMILY_JP, size=9)
    for col in cols:
        apply_formula_cell(ws[f'{col}18'], f"=-{col}17*{col}4")

    # 売上総利益 (Row 19) - 自然加算
    ws['B19'].value = "売上総利益"
    ws['B19'].font = Font(name=FONT_FAMILY_JP, size=10, bold=True)
    for col in cols:
        apply_formula_cell(ws[f'{col}19'], f"={col}17+{col}18", is_subtotal=True)

    # 販売費及び一般管理費 (Row 20)
    ws['B20'].value = "販売費及び一般管理費"
    ws['B20'].font = Font(name=FONT_FAMILY_JP, size=9)
    for col in cols:
        apply_formula_cell(ws[f'{col}20'], f"=-{col}17*{col}5")

    # 営業利益 (Row 21)
    ws['B21'].value = "営業利益（EBIT）"
    ws['B21'].font = Font(name=FONT_FAMILY_JP, size=10, bold=True)
    for col in cols:
        apply_formula_cell(ws[f'{col}21'], f"={col}19+{col}20", is_subtotal=True)

    # 支払利息 (Row 22) - デットスケジュールと連動
    ws['B22'].value = "支払利息"
    ws['B22'].font = Font(name=FONT_FAMILY_JP, size=9)
    for col in cols:
        apply_formula_cell(ws[f'{col}22'], f"=-{col}50") # 後述の借入利息を参照

    # 経常利益 / 税引前当期純利益 (Row 23)
    ws['B23'].value = "税引前当期純利益"
    ws['B23'].font = Font(name=FONT_FAMILY_JP, size=10, bold=True)
    for col in cols:
        apply_formula_cell(ws[f'{col}23'], f"={col}21+{col}22", is_subtotal=True)

    # 法人税等 (Row 24)
    ws['B24'].value = "法人税等"
    ws['B24'].font = Font(name=FONT_FAMILY_JP, size=9)
    for col in cols:
        apply_formula_cell(ws[f'{col}24'], f"=-MAX(0,{col}23*{col}6)")

    # 税引後当期純利益 (Row 25)
    ws['B25'].value = "税引後当期純利益"
    ws['B25'].font = Font(name=FONT_FAMILY_JP, size=10, bold=True)
    for col in cols:
        apply_formula_cell(ws[f'{col}25'], f"={col}23+{col}24", is_total=True)

    # ==========================================
    # （３）有形固定資産（PP&E）スケジュール
    # ==========================================
    apply_header_style(ws['A27'], "（３）有形固定資産")
    apply_header_style(ws['B27'], "有形固定資産（PP&E）スケジュール")
    for col, yr in zip(cols, years):
        apply_year_header(ws[f'{col}27'], yr)

    ws['B28'].value = "期首有形固定資産"
    ws['B29'].value = "設備投資（Capex・マイナス格納→期末残高は加算）"
    ws['B30'].value = "減価償却費"
    ws['B31'].value = "期末有形固定資産"
    ws['B31'].font = Font(name=FONT_FAMILY_JP, size=9, bold=True)

    apply_input_cell(ws['C31'], b_ppe) # 基準年実績（期末残高）

    for i, col in enumerate(["D", "E", "F", "G"]):
        prev_col = cols[i]
        apply_formula_cell(ws[f'{col}28'], f"={prev_col}31") # 前期末
        apply_formula_cell(ws[f'{col}29'], f"={col}7")       # Capex
        apply_formula_cell(ws[f'{col}30'], f"=-{col}28*{col}8") # 減価償却
        apply_formula_cell(ws[f'{col}31'], f"={col}28-{col}29+{col}30", is_subtotal=True)

    # ==========================================
    # （４）利益剰余金スケジュール
    # ==========================================
    apply_header_style(ws['A33'], "（４）利益剰余金")
    apply_header_style(ws['B33'], "利益剰余金スケジュール")
    for col, yr in zip(cols, years):
        apply_year_header(ws[f'{col}33'], yr)

    ws['B34'].value = "期首利益剰余金"
    ws['B35'].value = "当期純利益"
    ws['B36'].value = "配当金支払"
    ws['B37'].value = "期末利益剰余金"
    ws['B37'].font = Font(name=FONT_FAMILY_JP, size=9, bold=True)

    apply_input_cell(ws['C37'], b_re) # 基準年実績（期末残高）

    for i, col in enumerate(["D", "E", "F", "G"]):
        prev_col = cols[i]
        apply_formula_cell(ws[f'{col}34'], f"={prev_col}37")
        apply_formula_cell(ws[f'{col}35'], f"={col}25")
        apply_formula_cell(ws[f'{col}36'], f"=-MAX(0,{col}35*{col}12)")
        apply_formula_cell(ws[f'{col}37'], f"={col}34+{col}35+{col}36", is_subtotal=True)

    # ==========================================
    # （５）運転資本（Working Capital）スケジュール
    # ==========================================
    apply_header_style(ws['A39'], "（５）運転資本")
    apply_header_style(ws['B39'], "運転資本（Working Capital）スケジュール")
    for col, yr in zip(cols, years):
        apply_year_header(ws[f'{col}39'], yr)

    ws['B40'].value = "売掛金（売上/365*日数）"
    ws['B41'].value = "棚卸資産（原価/365*日数）"
    ws['B42'].value = "買掛金（原価/365*日数）"
    ws['B43'].value = "正味運転資本（OWC）"
    ws['B43'].font = Font(name=FONT_FAMILY_JP, size=9, bold=True)

    apply_input_cell(ws['C40'], b_ar) # 基準年実績
    apply_input_cell(ws['C41'], b_inv)
    apply_input_cell(ws['C42'], b_ap)
    apply_formula_cell(ws['C43'], "=C40+C41-C42", is_subtotal=True)
    for col in ["D", "E", "F", "G"]:
        apply_formula_cell(ws[f'{col}40'], f"=({col}17/365)*{col}9")
        apply_formula_cell(ws[f'{col}41'], f"=(-{col}18/365)*{col}10")
        apply_formula_cell(ws[f'{col}42'], f"=(-{col}18/365)*{col}11")
        apply_formula_cell(ws[f'{col}43'], f"={col}40+{col}41-{col}42", is_subtotal=True)

    # ==========================================
    # （６）借入金（Debt）スケジュール
    # ==========================================
    apply_header_style(ws['A45'], "（６）借入金")
    apply_header_style(ws['B45'], "借入金（Debt）スケジュール")
    for col, yr in zip(cols, years):
        apply_year_header(ws[f'{col}45'], yr)

    ws['B46'].value = "期首借入金残高"
    ws['B47'].value = "約定返済額"
    ws['B48'].value = "新規借入額"
    ws['B49'].value = "期末借入金残高"
    ws['B49'].font = Font(name=FONT_FAMILY_JP, size=9, bold=True)
    ws['B50'].value = "支払利息（期首残高*金利: 循環回避）"

    apply_input_cell(ws['C49'], b_debt) # 基準年実績（期末残高）
    apply_input_cell(ws['C50'], b_int)   # 基準年実績（支払利息）

    for i, col in enumerate(["D", "E", "F", "G"]):
        prev_col = cols[i]
        apply_formula_cell(ws[f'{col}46'], f"={prev_col}49")
        apply_formula_cell(ws[f'{col}47'], f"={col}13")
        apply_input_cell(ws[f'{col}48'], 0)
        apply_formula_cell(ws[f'{col}49'], f"={col}46+{col}47+{col}48", is_subtotal=True)
        apply_formula_cell(ws[f'{col}50'], f"={col}46*{col}14")

    # ==========================================
    # （７）貸借対照表（BS）
    # ==========================================
    apply_header_style(ws['A52'], "（７）貸借対照表")
    apply_header_style(ws['B52'], "貸借対照表（BS）")
    for col, yr in zip(cols, years):
        apply_year_header(ws[f'{col}52'], yr)

    ws['B53'].value = "【資産の部】"
    ws['B53'].font = Font(name=FONT_FAMILY_JP, size=9, bold=True)
    ws['B54'].value = "現預金"
    ws['B55'].value = "売掛金"
    ws['B56'].value = "棚卸資産"
    ws['B57'].value = "有形固定資産（PP&E）"
    ws['B58'].value = "資産合計"
    ws['B58'].font = Font(name=FONT_FAMILY_JP, size=10, bold=True)

    ws['B60'].value = "【負債の部】"
    ws['B60'].font = Font(name=FONT_FAMILY_JP, size=9, bold=True)
    ws['B61'].value = "買掛金"
    ws['B62'].value = "借入金"
    ws['B63'].value = "【純資産の部】"
    ws['B63'].font = Font(name=FONT_FAMILY_JP, size=9, bold=True)
    ws['B64'].value = "資本金"
    ws['B65'].value = "利益剰余金"
    ws['B66'].value = "負債・純資産合計"
    ws['B66'].font = Font(name=FONT_FAMILY_JP, size=10, bold=True)

    # バランスチェック行（常設）
    ws['B67'].value = "バランスチェック（差額＝0検証）"
    ws['B67'].font = Font(name=FONT_FAMILY_JP, size=9, bold=True, color="C00000")

    # 実績 C列
    apply_formula_cell(ws['C54'], "=C81") # CF期末残高
    apply_formula_cell(ws['C55'], "=C40")
    apply_formula_cell(ws['C56'], "=C41")
    apply_formula_cell(ws['C57'], "=C31")
    apply_formula_cell(ws['C58'], "=SUM(C54:C57)", is_total=True)

    apply_formula_cell(ws['C61'], "=C42")
    apply_formula_cell(ws['C62'], "=C49")
    apply_input_cell(ws['C64'], b_cap) # 資本金（基準年で貸借一致）
    apply_formula_cell(ws['C65'], "=C37")
    apply_formula_cell(ws['C66'], "=C61+C62+C64+C65", is_total=True)
    apply_formula_cell(ws['C67'], "=C58-C66")

    # 予測 D〜G列
    for col in ["D", "E", "F", "G"]:
        apply_formula_cell(ws[f'{col}54'], f"={col}81") # CF期末残高
        apply_formula_cell(ws[f'{col}55'], f"={col}40")
        apply_formula_cell(ws[f'{col}56'], f"={col}41")
        apply_formula_cell(ws[f'{col}57'], f"={col}31")
        apply_formula_cell(ws[f'{col}58'], f"=SUM({col}54:{col}57)", is_total=True)

        apply_formula_cell(ws[f'{col}61'], f"={col}42")
        apply_formula_cell(ws[f'{col}62'], f"={col}49")
        apply_formula_cell(ws[f'{col}64'], "=C64") # 資本金据置
        apply_formula_cell(ws[f'{col}65'], f"={col}37")
        apply_formula_cell(ws[f'{col}66'], f"={col}61+{col}62+{col}64+{col}65", is_total=True)
        apply_formula_cell(ws[f'{col}67'], f"={col}58-{col}66") # 差額0

    # ==========================================
    # （８）キャッシュフロー計算書（CF）
    # ==========================================
    apply_header_style(ws['A69'], "（８）キャッシュフロー計算書")
    apply_header_style(ws['B69'], "キャッシュフロー計算書（間接法CF）")
    for col, yr in zip(cols, years):
        apply_year_header(ws[f'{col}69'], yr)

    ws['B70'].value = "税引後当期純利益"
    ws['B71'].value = "減価償却費（足し戻し）"
    ws['B72'].value = "Δ運転資本（前期−当期）"
    ws['B73'].value = "営業活動によるキャッシュフロー"
    ws['B73'].font = Font(name=FONT_FAMILY_JP, size=10, bold=True)

    ws['B74'].value = "設備投資額（Capex）"
    ws['B75'].value = "投資活動によるキャッシュフロー"
    ws['B75'].font = Font(name=FONT_FAMILY_JP, size=10, bold=True)

    ws['B76'].value = "借入返済額"
    ws['B77'].value = "新規借入額"
    ws['B78'].value = "配当金支払額"
    ws['B79'].value = "財務活動によるキャッシュフロー"
    ws['B79'].font = Font(name=FONT_FAMILY_JP, size=10, bold=True)

    ws['B80'].value = "当期純キャッシュ増減（Net Cash Flow）"
    ws['B80'].font = Font(name=FONT_FAMILY_JP, size=10, bold=True)
    ws['B81'].value = "期末現預金残高"
    ws['B81'].font = Font(name=FONT_FAMILY_JP, size=10, bold=True)

    # 実績 C列（ベース期）
    apply_input_cell(ws['C81'], b_cash) # 実績手元現金

    # 予測 D〜G列
    for i, col in enumerate(["D", "E", "F", "G"]):
        prev_col = cols[i]
        apply_formula_cell(ws[f'{col}70'], f"={col}25")
        apply_formula_cell(ws[f'{col}71'], f"=-{col}30") # プラスで足し戻し
        apply_formula_cell(ws[f'{col}72'], f"={prev_col}43-{col}43") # 前期OWC - 当期OWC
        apply_formula_cell(ws[f'{col}73'], f"=SUM({col}70:{col}72)", is_subtotal=True)

        apply_formula_cell(ws[f'{col}74'], f"={col}29") # Capex(マイナス)
        apply_formula_cell(ws[f'{col}75'], f"={col}74", is_subtotal=True)

        apply_formula_cell(ws[f'{col}76'], f"={col}47") # 借入返済(マイナス)
        apply_formula_cell(ws[f'{col}77'], f"={col}48") # 新規借入(プラス)
        apply_formula_cell(ws[f'{col}78'], f"={col}36") # 配当(マイナス)
        apply_formula_cell(ws[f'{col}79'], f"=SUM({col}76:{col}78)", is_subtotal=True)

        apply_formula_cell(ws[f'{col}80'], f"={col}73+{col}75+{col}79", is_total=True)
        apply_formula_cell(ws[f'{col}81'], f"={prev_col}81+{col}80", is_total=True)

    # 基準年 C列の現金連動
    apply_formula_cell(ws['C54'], "=C81")

    # ==========================================
    # （９）DCF バリュエーション（WACC・PGRは入力セル化、感応度マトリクス付き）
    # ==========================================
    apply_header_style(ws['A83'], "（９）DCF 企業価値算定")
    apply_header_style(ws['B83'], "DCF 企業価値算定（バリュエーション）")
    for col, yr in zip(cols, years):
        apply_year_header(ws[f'{col}83'], yr)

    labels = {
        84: "割引率（WACC）",
        85: "永久成長率（PGR）",
        86: "割引期間（年央調整: t-0.5）",
        87: "営業利益（EBIT）",
        88: "みなし税引後営業利益（NOPAT）",
        89: "減価償却費（足し戻し）",
        90: "設備投資（Capex）",
        91: "Δ運転資本（流出はマイナス）",
        92: "フリーキャッシュフロー（FCFF）",
        93: "ディスカウントファクター",
        94: "現在価値（PV of FCF）",
    }
    for r, t in labels.items():
        ws[f'B{r}'].value = t
        ws[f'B{r}'].font = Font(name=FONT_FAMILY_JP, size=10 if r in (92, 94) else 9, bold=(r in (92, 94)))

    dcf_cfg = p.get("dcf", {})
    wacc_val = dcf_cfg.get("wacc", 0.07)
    pgr_val = dcf_cfg.get("pgr", 0.005)
    wacc_rng = dcf_cfg.get("wacc_range", [0.06, 0.065, 0.07, 0.075, 0.08])
    pgr_rng = dcf_cfg.get("pgr_range", [0.0, 0.005, 0.01, 0.015])

    apply_input_cell(ws['D84'], wacc_val, is_percent=True)
    apply_input_cell(ws['D85'], pgr_val, is_percent=True)
    apply_input_cell(ws['D86'], 0.5)
    ws['D86'].number_format = "0.0"
    for i, col in enumerate(["D", "E", "F", "G"]):
        if i > 0:
            apply_formula_cell(ws[f'{col}86'], f"={cols[i]}86+1")
            ws[f'{col}86'].number_format = "0.0"
        apply_formula_cell(ws[f'{col}87'], f"={col}21")
        apply_formula_cell(ws[f'{col}88'], f"={col}87*(1-{col}6)")
        apply_formula_cell(ws[f'{col}89'], f"=-{col}30")
        apply_formula_cell(ws[f'{col}90'], f"={col}29")
        apply_formula_cell(ws[f'{col}91'], f"={col}72")
        apply_formula_cell(ws[f'{col}92'], f"={col}88+{col}89+{col}90+{col}91", is_subtotal=True)
        apply_formula_cell(ws[f'{col}93'], f"=1/(1+$D$84)^{col}86")
        ws[f'{col}93'].number_format = "0.0000"
        apply_formula_cell(ws[f'{col}94'], f"={col}92*{col}93", is_total=True)

    ws['B96'].value = "予測期間FCF現在価値合計"
    apply_formula_cell(ws['D96'], "=SUM(D94:G94)", is_total=True)
    ws['B97'].value = "ターミナルバリュー現在価値（永久成長率法）"
    apply_formula_cell(ws['D97'], "=G92*(1+D85)/(D84-D85)*G93", is_total=True)
    ws['B98'].value = "事業価値（EV: Enterprise Value）"
    ws['B98'].font = Font(name=FONT_FAMILY_JP, size=10, bold=True)
    apply_formula_cell(ws['D98'], "=D96+D97", is_total=True)
    ws['B99'].value = "現預金（非事業用資産）"
    apply_formula_cell(ws['D99'], "=C54")
    ws['B100'].value = "有利子負債（有利子負債控除）"
    apply_formula_cell(ws['D100'], "=-C62")
    ws['B101'].value = "株式価値（Equity Value）"
    ws['B101'].font = Font(name=FONT_FAMILY_JP, size=11, bold=True, color="1B365D")
    apply_formula_cell(ws['D101'], "=D98+D99+D100", is_total=True)

    # 感応度マトリクス（WACC × PGR → 株式価値）: 数式直接計算（TABLE機能非依存）
    apply_header_style(ws['A103'], "（10）感応度分析")
    apply_header_style(ws['B103'], "感応度マトリクス（株式価値: WACC × PGR）")
    for col in cols:
        apply_year_header(ws[f'{col}103'], "")
    ws['C104'].value = "WACC＼PGR"
    ws['C104'].font = Font(name=FONT_FAMILY_JP, size=9, bold=True)
    ws['C104'].alignment = Alignment(horizontal="right")
    for col, g in zip(["D", "E", "F", "G"], pgr_rng):
        apply_input_cell(ws[f'{col}104'], g, is_percent=True)
    for r, w in zip(range(105, 110), wacc_rng):
        apply_input_cell(ws[f'C{r}'], w, is_percent=True)
        for col in ["D", "E", "F", "G"]:
            f = (f"=SUMPRODUCT($D$92:$G$92,1/(1+$C{r})^$D$86:$G$86)"
                 f"+$G$92*(1+{col}$104)/($C{r}-{col}$104)/(1+$C{r})^$G$86+$D$99+$D$100")
            apply_formula_cell(ws[f'{col}{r}'], f)

def main():
    parser = argparse.ArgumentParser(description="財務モデリングExcel自動構築ジェネレータ")
    parser.add_argument("--type", choices=["single", "multi", "dcf"], default="single", help="モデル型 (single, multi, dcf)")
    parser.add_argument("--title", default="財務業績予測・3表連動モデル", help="モデルのタイトル")
    parser.add_argument("-o", "--output", default="financial_model.xlsx", help="出力Excelファイル名")
    parser.add_argument("--config", help="カスタムパラメータJSONファイル")
    args = parser.parse_args()

    wb = openpyxl.Workbook()

    params = None
    if args.config:
        if os.path.exists(args.config):
            with open(args.config, "r", encoding="utf-8") as f:
                params = json.load(f)
            print(f"カスタムパラメータ読込: {args.config}")
        else:
            print(f"\033[31m[ERROR]\033[0m 設定ファイルが見つかりません: {args.config}")
            sys.exit(1)

    if args.type == "single" or args.type == "dcf":
        print(f"財務モデル構築中 [Type-01: 単一シート3表連動 + DCF] -> {args.output}")
        build_single_sheet_model(wb, title=args.title, params=params)
    else:
        print("\033[33m[NOTE]\033[0m Type-02（マルチシート）は未実装のため Type-01 で生成します")
        build_single_sheet_model(wb, title=args.title, params=params)

    wb.save(args.output)
    print(f"\033[32m[SUCCESS]\033[0m Excelモデルを生成しました: {args.output}")

    # 自動バリデーション連携
    script_dir = os.path.dirname(os.path.abspath(__file__))
    check_script = os.path.join(script_dir, "check_model.py")
    if os.path.exists(check_script):
        print("\n--- 自動バリデーション実行 (check_model.py) ---")
        os.system(f"python3 {check_script} {args.output}")

if __name__ == "__main__":
    main()
