"""フロンタル 月次データ集計 (2026-06-03 run)"""
import pandas as pd
import json
from pathlib import Path

RAW_DIR = Path("/sessions/busy-dreamy-knuth/mnt/Downloads/Frontal/01_DriveDoor_Raw")
DATE_TAG = "20260603"

def load_nippo(path):
    return pd.read_excel(path, header=0)

honsha = load_nippo(RAW_DIR / "日報一覧" / f"日報一覧_本社_{DATE_TAG}.xlsx")
kyoto = load_nippo(RAW_DIR / "日報一覧" / f"日報一覧_京都_{DATE_TAG}.xlsx")
fjs = load_nippo(RAW_DIR / "日報一覧" / f"日報一覧_FJS_{DATE_TAG}.xlsx")

COL_DATE="運行日"; COL_KUBUN="配車先区分"; COL_YOSHA="傭車先名称"; COL_BILL="請求金額"

def aggregate(df, dept):
    out = {"ippan":[0]*12, "riyo":[0]*12}
    if len(df)==0: return out
    df=df.copy()
    df[COL_DATE]=pd.to_datetime(df[COL_DATE],errors='coerce')
    df[COL_BILL]=pd.to_numeric(df[COL_BILL],errors='coerce').fillna(0)
    df["_m"]=df[COL_DATE].dt.month
    df["_y"]=df[COL_YOSHA].fillna("").astype(str)
    df["_k"]=df[COL_KUBUN].fillna("").astype(str)
    for _,r in df.iterrows():
        m=r["_m"]
        if pd.isna(m) or m<1 or m>12: continue
        m=int(m)-1; bill=r[COL_BILL]; k=r["_k"]; y=r["_y"]
        if dept=="本社":
            if k=="自社": out["ippan"][m]+=bill
            else: out["riyo"][m]+=bill
        elif dept=="京都":
            if k=="自社": out["ippan"][m]+=bill
        elif dept=="FJS":
            if k=="自社": out["ippan"][m]+=bill
            elif "シクロ" in y: out["riyo"][m]+=bill
    return out

ah=aggregate(honsha,"本社"); ak=aggregate(kyoto,"京都"); af=aggregate(fjs,"FJS")
ippan_sales=[round((ah["ippan"][i]+ak["ippan"][i]+af["ippan"][i])/1000) for i in range(12)]
riyo_sales =[round((ah["riyo"][i]+ak["riyo"][i]+af["riyo"][i])/1000) for i in range(12)]
print("一般売上(千円):",ippan_sales)
print("利用売上(千円):",riyo_sales)

csv_path=RAW_DIR/"車両経費"/f"車両経費一覧_{DATE_TAG}.csv"
shrk=pd.read_csv(csv_path,encoding="utf-8-sig")
shrk["発生年月日"]=pd.to_datetime(shrk["発生年月日"],errors='coerce')
shrk["_m"]=shrk["発生年月日"].dt.month
shrk["経費金額"]=pd.to_numeric(shrk["経費金額"],errors='coerce').fillna(0)
EXCLUDE={"本社固定費","京都固定費","リース料","倉庫賃料","看板","税金・保険・印紙代等"}
CMAP={"燃料":"nenryo","通行料":"kotsu","固定車両費":"sharyo","タイヤ":"sharyo","オイル":"sharyo","修繕費":"sharyo","法定福利費":"jinken","諸経費":"jinken","保険料":"hoken"}
cc={k:[0]*12 for k in ["nenryo","kotsu","sharyo","jinken","hoken"]}
seen={}
for _,r in shrk.iterrows():
    cat=str(r["車両整備経費区分"]); seen[cat]=seen.get(cat,0)+1
    if cat in EXCLUDE: continue
    t=CMAP.get(cat)
    if not t: continue
    m=r["_m"]
    if pd.isna(m) or m<1 or m>12: continue
    cc[t][int(m)-1]+=r["経費金額"]
print("車両経費 区分一覧:",seen)
cost_nenryo=[round(v/1000) for v in cc["nenryo"]]
cost_kotsu =[round(v/1000) for v in cc["kotsu"]]
cost_sharyo=[round(v/1000) for v in cc["sharyo"]]
cost_jinken=[round(v/1000) for v in cc["jinken"]]
cost_hoken =[round(v/1000) for v in cc["hoken"]]
ippan_cost=[cost_nenryo[i]+cost_kotsu[i]+cost_sharyo[i]+cost_jinken[i]+cost_hoken[i] for i in range(12)]
print("一般原価合計(千円):",ippan_cost)
print("燃料費:",cost_nenryo); print("交通費:",cost_kotsu); print("車両費:",cost_sharyo)
print("人件費:",cost_jinken); print("保険料:",cost_hoken)

def agg_riyo_cost(df,dept):
    out=[0]*12
    if len(df)==0: return out
    df=df.copy()
    df["運行日"]=pd.to_datetime(df["運行日"],errors='coerce')
    df["_m"]=df["運行日"].dt.month
    df["支払金額"]=pd.to_numeric(df.get("支払金額",0),errors='coerce').fillna(0)
    df["_y"]=df.get("傭車先名称",pd.Series([""]*len(df))).fillna("").astype(str)
    df["_k"]=df.get("配車先区分",pd.Series([""]*len(df))).fillna("").astype(str)
    for _,r in df.iterrows():
        m=r["_m"]
        if pd.isna(m) or m<1 or m>12: continue
        m=int(m)-1; pay=r["支払金額"]; k=r["_k"]; y=r["_y"]
        if dept=="本社":
            if k!="自社": out[m]+=pay
        elif dept=="FJS":
            if k!="自社" and "シクロ" in y: out[m]+=pay
    return out
rch=agg_riyo_cost(honsha,"本社"); rcf=agg_riyo_cost(fjs,"FJS")
riyo_cost=[round((rch[i]+rcf[i])/1000) for i in range(12)]
print("利用原価(千円):",riyo_cost)

TODAY_MONTH=6
def to_act(a,mm=TODAY_MONTH): return [a[i] if i<mm else None for i in range(12)]
result={
 "ippan_sales_act":to_act(ippan_sales),
 "riyo_sales_act":to_act(riyo_sales),
 "ippan_cost_act":to_act(ippan_cost),
 "riyo_cost_act":to_act(riyo_cost),
 "cost_nenryo_act":to_act(cost_nenryo),
 "cost_kotsu_act":to_act(cost_kotsu),
 "cost_sharyo_act":to_act(cost_sharyo),
 "cost_jinken_act":to_act(cost_jinken),
 "cost_hoken_act":to_act(cost_hoken),
}
with open("/sessions/busy-dreamy-knuth/mnt/outputs/aggregated.json",'w',encoding='utf-8') as f:
    json.dump(result,f,ensure_ascii=False,indent=2)
print("\n=== JSON ===")
print(json.dumps(result,ensure_ascii=False,indent=2))
