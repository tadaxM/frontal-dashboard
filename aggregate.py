#!/usr/bin/env python3
"""フロンタル予算実績ダッシュボード — データ集計スクリプト"""

import openpyxl
import csv
import json
import math
from datetime import datetime
from pathlib import Path

DATA_DIR = Path(__file__).parent / "data"

# --- 日報Excel読み取り ---
def read_nippo(filepath, office_type):
    """
    日報Excelを読み取り、月別・区分別に売上・原価を集計する。
    office_type: 'honsha', 'kyoto', 'fjs'

    本社・京都: 配車先区分(Col H)が「自社」→ 一般(ippan)、「傭車」→ 利用(riyo)
    FJS: すべて利用(riyo)に分類
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    # {(category, month): total}  category = 'ippan' or 'riyo'
    result = {'ippan_sales': {}, 'ippan_cost': {}, 'riyo_sales': {}, 'riyo_cost': {}}

    for row in ws.iter_rows(min_row=2, values_only=True):
        date_val = row[0]  # Col A: 運行日
        if date_val is None:
            continue

        # 日付パース
        if isinstance(date_val, datetime):
            month = date_val.month
        elif isinstance(date_val, str):
            try:
                dt = datetime.strptime(date_val.strip(), "%Y/%m/%d")
                month = dt.month
            except ValueError:
                continue
        else:
            continue

        # 区分判定
        haisha = str(row[7]).strip() if len(row) > 7 and row[7] else ''
        if office_type == 'fjs':
            # FJS: 自社 → 一般(京都に加算)
            # 傭車で傭車先=シクロシュプリーム → 利用(京都に加算)
            # その他傭車 → 除外(フロンタルに関係ない)
            yosha_name = str(row[40]).strip() if len(row) > 40 and row[40] else ''
            if haisha == '自社':
                category = 'ippan'
            elif haisha == '傭車' and 'シクロ' in yosha_name:
                category = 'riyo'
            else:
                continue  # フロンタルに関係ない傭車は除外
        elif office_type == 'kyoto':
            # 京都: 自社 → 一般、傭車 → 保留(除外)
            if haisha == '自社':
                category = 'ippan'
            else:
                continue  # 京都の傭車は保留
        else:
            # 本社: 配車先区分(Col H, index 7)で判定
            if haisha == '傭車':
                category = 'riyo'
            else:
                category = 'ippan'  # 自社 or その他 → 一般

        # 請求金額 (index 89) / 支払金額 (index 100)
        # ※DriveDoorのExcel出力は列数が変動する場合があるため、
        #   ヘッダー行から動的に取得するのが望ましいが、現時点ではindex固定
        sales = row[89] if len(row) > 89 else None
        # 支払金額
        cost = row[100] if len(row) > 100 else None

        sales = float(sales) if sales is not None else 0
        cost = float(cost) if cost is not None else 0

        result[f'{category}_sales'][month] = result[f'{category}_sales'].get(month, 0) + sales
        result[f'{category}_cost'][month] = result[f'{category}_cost'].get(month, 0) + cost

    wb.close()
    return result


# --- 車両経費CSV読み取り ---
def read_sharyo_keihi(filepath):
    """車両経費CSVを読み取り、月別・費目別に集計する。"""
    # CSV「車両整備経費区分」→ ダッシュボードカテゴリ のマッピング
    # 本社固定費, 京都固定費, リース料, 倉庫賃料, 看板, 税金・保険・印紙代等 は
    # 固定費（CF_FIXED）に含まれるため除外
    category_map = {
        '燃料': 'nenryo',
        '通行料': 'kotsu',
        '固定車両費': 'sharyo',
        'タイヤ': 'sharyo',
        'オイル': 'sharyo',
        '修繕費': 'sharyo',
        '保険料': 'hoken',
        '法定福利費': 'jinken',
        '諸経費': 'jinken',
    }

    # {category: {month: total}}
    monthly_costs = {cat: {} for cat in ['nenryo', 'kotsu', 'sharyo', 'jinken', 'hoken']}

    with open(filepath, encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            # 発生年月日から月を取得
            date_str = row.get('発生年月日', '').strip()
            if not date_str:
                continue
            try:
                dt = datetime.strptime(date_str, "%Y/%m/%d")
                month = dt.month
            except ValueError:
                continue

            # 車両整備経費区分
            himoku = row.get('車両整備経費区分', '').strip()
            cat_key = category_map.get(himoku)
            if cat_key is None:
                continue

            # 経費金額
            amount_str = row.get('経費金額', '0').strip().replace(',', '')
            try:
                amount = float(amount_str)
            except ValueError:
                amount = 0

            monthly_costs[cat_key][month] = monthly_costs[cat_key].get(month, 0) + amount

    return monthly_costs


def round_sen(value):
    """円を千円に変換（四捨五入）"""
    if value is None:
        return None
    return round(value / 1000)


def main():
    # --- 日報読み取り ---
    honsha = read_nippo(DATA_DIR / "nippo_honsha.xlsx", "honsha")
    kyoto = read_nippo(DATA_DIR / "nippo_kyoto.xlsx", "kyoto")
    fjs = read_nippo(DATA_DIR / "nippo_fjs.xlsx", "fjs")

    # --- 集約: 全事業所のippan/riyoを合算 ---
    # 一般(ippan) = 本社自社 + 京都自社 + FJS自社
    # 利用(riyo) = 本社傭車 + 京都傭車 + FJS傭車(得意先フロンタルのみ)
    all_months = set()
    for data in [honsha, kyoto, fjs]:
        for key in data:
            all_months.update(data[key].keys())
    max_month = max(all_months) if all_months else 3

    # 月別集計（1〜12月）
    ippan_sales_act = []
    riyo_sales_act = []
    riyo_cost_act = []

    for m in range(1, 13):
        if m <= max_month:
            is_val = (honsha['ippan_sales'].get(m, 0) + kyoto['ippan_sales'].get(m, 0)
                      + fjs['ippan_sales'].get(m, 0))
            ippan_sales_act.append(round_sen(is_val))

            rs_val = (honsha['riyo_sales'].get(m, 0) + kyoto['riyo_sales'].get(m, 0)
                      + fjs['riyo_sales'].get(m, 0))
            rc_val = (honsha['riyo_cost'].get(m, 0) + kyoto['riyo_cost'].get(m, 0)
                      + fjs['riyo_cost'].get(m, 0))
            riyo_sales_act.append(round_sen(rs_val))
            riyo_cost_act.append(round_sen(rc_val))
        else:
            ippan_sales_act.append(None)
            riyo_sales_act.append(None)
            riyo_cost_act.append(None)

    # 粗利計算 (一般原価は後で車両経費CSVから算出するため、仮にNone)
    riyo_gross_act = []
    for m in range(12):
        if riyo_sales_act[m] is not None and riyo_cost_act[m] is not None:
            riyo_gross_act.append(riyo_sales_act[m] - riyo_cost_act[m])
        else:
            riyo_gross_act.append(None)

    # --- 車両経費CSV読み取り ---
    sharyo_costs = read_sharyo_keihi(DATA_DIR / "sharyokeihi.csv")

    cost_nenryo_act = []
    cost_kotsu_act = []
    cost_sharyo_act = []
    cost_jinken_act = []
    cost_hoken_act = []

    for m in range(1, 13):
        if m <= max_month:
            cost_nenryo_act.append(round_sen(sharyo_costs['nenryo'].get(m, 0)))
            cost_kotsu_act.append(round_sen(sharyo_costs['kotsu'].get(m, 0)))
            cost_sharyo_act.append(round_sen(sharyo_costs['sharyo'].get(m, 0)))
            cost_jinken_act.append(round_sen(sharyo_costs['jinken'].get(m, 0)))
            cost_hoken_act.append(round_sen(sharyo_costs['hoken'].get(m, 0)))
        else:
            cost_nenryo_act.append(None)
            cost_kotsu_act.append(None)
            cost_sharyo_act.append(None)
            cost_jinken_act.append(None)
            cost_hoken_act.append(None)

    # 費目別原価がすべて0の月はnullにする（未入力扱い）
    for m in range(12):
        month_num = m + 1
        if month_num <= max_month:
            total_cost_items = sum(x or 0 for x in [
                cost_nenryo_act[m], cost_kotsu_act[m], cost_sharyo_act[m],
                cost_jinken_act[m], cost_hoken_act[m]
            ])
            if total_cost_items == 0:
                cost_nenryo_act[m] = None
                cost_kotsu_act[m] = None
                cost_sharyo_act[m] = None
                cost_jinken_act[m] = None
                cost_hoken_act[m] = None

    # 一般原価 = 車両経費CSV費目合計 (燃料+交通+車両+人件+保険)
    ippan_cost_act = []
    ippan_gross_act = []
    for m in range(12):
        cost_items = [cost_nenryo_act[m], cost_kotsu_act[m], cost_sharyo_act[m],
                      cost_jinken_act[m], cost_hoken_act[m]]
        if any(v is not None for v in cost_items):
            total_cost = sum(v or 0 for v in cost_items)
            ippan_cost_act.append(total_cost)
            if ippan_sales_act[m] is not None:
                ippan_gross_act.append(ippan_sales_act[m] - total_cost)
            else:
                ippan_gross_act.append(None)
        else:
            ippan_cost_act.append(None)
            ippan_gross_act.append(None)

    # --- 出力 ---
    result = {
        'ippan_sales_act': ippan_sales_act,
        'ippan_cost_act': ippan_cost_act,
        'ippan_gross_act': ippan_gross_act,
        'riyo_sales_act': riyo_sales_act,
        'riyo_cost_act': riyo_cost_act,
        'riyo_gross_act': riyo_gross_act,
        'cost_nenryo_act': cost_nenryo_act,
        'cost_kotsu_act': cost_kotsu_act,
        'cost_sharyo_act': cost_sharyo_act,
        'cost_jinken_act': cost_jinken_act,
        'cost_hoken_act': cost_hoken_act,
    }

    print(json.dumps(result, ensure_ascii=False, indent=2))

    # サマリー表示
    print("\n--- サマリー ---")
    for m in range(max_month):
        mn = m + 1
        ig = ippan_gross_act[m]
        rg = riyo_gross_act[m]
        ig_str = f"{ig:,}" if ig is not None else "算出不可"
        rg_str = f"{rg:,}" if rg is not None else "算出不可"
        print(f"  {mn}月: 一般売上={ippan_sales_act[m]:,} 一般粗利={ig_str} 利用粗利={rg_str}")


if __name__ == "__main__":
    main()
