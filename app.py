"""
GST Audit Reconciliation Tool
Created by: Purva Doshi | linkedin.com/in/purvadoshi26
Tabs: ITC Reco (invoice-wise) | Sales Reco | 3B Check (3B vs GSTR-1 vs Books vs 2B)
"""
import streamlit as st
import io, re, datetime, hashlib
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict

st.set_page_config(page_title="GST Audit Reco Tool", page_icon="📊",
                   layout="wide", initial_sidebar_state="expanded")
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@300;400;500;600&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif}
.hdr{background:linear-gradient(135deg,#0f2447 0%,#1a3a6e 60%,#0d3055 100%);
     padding:1.2rem 1.6rem 1rem;border-radius:10px;margin-bottom:1.2rem;
     box-shadow:0 4px 20px rgba(0,0,0,.25)}
.hdr h1{color:#fff;margin:0;font-family:'DM Serif Display',serif;font-size:1.45rem}
.hdr .sub{color:#90b8e0;font-size:.82rem;margin-top:3px}
.hdr .cred{margin-top:6px;font-size:.76rem;color:#5fa0d0}
.hdr .cred a{color:#7ec8f0;text-decoration:none;border-bottom:1px solid #7ec8f030}
.fmtbox{background:#EEF4FB;border-left:4px solid #1a3a6e;padding:.6rem .9rem;
        border-radius:0 8px 8px 0;font-size:.8rem;margin:.3rem 0 .7rem;line-height:1.5}
div[data-testid="stDownloadButton"] button{
    background:linear-gradient(135deg,#0f2447,#1a3a6e);color:white;border:none;
    padding:.6rem 1.2rem;border-radius:8px;font-size:.9rem;font-weight:600;width:100%}
</style>""", unsafe_allow_html=True)

# ── Excel styles ──────────────────────────────────────────────────────────────
NAVY="1F3864";BLUE="2E75B6";TEAL="1B6B7B";GREEN="1E5631";PURPLE="4B0082";TITLE="0F2447"
MATCH="D6E4BC";MIS="FFD7D7";ONLY2="FFF2CC";ONLY1="D6E4F7";RCM_C="E8F5E9";CUST_C="FEF3E2";EXEMPT="F5F5F5"
INR="#,##0.00";NUM="#,##0"
def F(h): return PatternFill("solid",start_color=h,end_color=h)
def HF(sz=9,c="FFFFFF"): return Font(name="Calibri",bold=True,size=sz,color=c)
def DF(sz=9,c="222222",bold=False): return Font(name="Calibri",size=sz,color=c,bold=bold)
def BD():
    s=Side(style="thin",color="DDDDDD"); return Border(left=s,right=s,top=s,bottom=s)
def CTR(): return Alignment(horizontal="center",vertical="center",wrap_text=True)
def LFT(): return Alignment(horizontal="left",vertical="center")
def RGT(): return Alignment(horizontal="right",vertical="center")
def safe(v):
    try: return float(str(v or "0").replace(",","").strip())
    except: return 0.0
def cgstin(s): return re.sub(r"[^A-Z0-9]","",str(s or "").upper().strip())
def cname(s):  return re.sub(r"[^a-z0-9 ]","",str(s or "").lower().strip())
def norm_inv(s): return re.sub(r"[-\s]","",str(s or "").strip().upper())

# ── File loader ───────────────────────────────────────────────────────────────
def load_wb(fb):
    try: return openpyxl.load_workbook(io.BytesIO(fb),data_only=True)
    except:
        try: return openpyxl.load_workbook(io.BytesIO(fb),data_only=True,keep_vba=True)
        except Exception as e: raise ValueError(f"Cannot open file: {e}\nSave as .xlsx in Excel.")

# ─────────────────────────────────────────────────────────────────────────────
#  TAB 1 — ITC RECO (GSTR-2B vs Purchase)
# ─────────────────────────────────────────────────────────────────────────────
def words(s):
    STOP={"pvt","ltd","private","limited","and","co","the","ms","mrs","india","p","l","m/s"}
    return set(re.findall(r"[a-z0-9]+",str(s).lower()))-STOP

def detect_pr_format(ws):
    import datetime as _dt
    rows=list(ws.iter_rows(min_row=1,max_row=min(25,ws.max_row),values_only=True))
    # FORMAT A: Date+Particulars header confirmed by datetime in data
    for i,row in enumerate(rows):
        rl=[str(v or "").strip().lower() for v in row]
        if "date" in rl and any(v in rl for v in ["particulars","party","vendor","supplier"]) and sum(1 for v in rl if v)>=4:
            for j in range(i+1,min(i+5,len(rows))):
                if rows[j][0] and isinstance(rows[j][0],(_dt.datetime,_dt.date)): return "A",i
    # FORMAT C: Pivot — strict "sum of input" check
    for i,row in enumerate(rows):
        vl=[str(v or "").lower() for v in row]
        if any("row labels" in v for v in vl) or sum(1 for v in vl if "sum of input" in v)>=2:
            if i+1<len(rows) and rows[i+1][0] and not str(rows[i+1][0]).strip()[0].isdigit():
                return "C",i
    # FORMAT B: sub-row fallback
    for i,row in enumerate(rows):
        if row[0] and str(row[0]).strip().lower()=="date": return "B",i
    raise ValueError("Cannot detect purchase register format.\n\nNeeds: Date | Particulars | GSTIN | INPUT IGST/CGST/SGST columns")

def parse_purchase(fb):
    wb=load_wb(fb)
    ws=None
    for sn in wb.sheetnames:
        if "purchase register" in sn.lower(): ws=wb[sn]; break
    if not ws: ws=wb[wb.sheetnames[0]]
    rows=list(ws.iter_rows(min_row=1,max_row=ws.max_row,values_only=True))
    fmt,hr=detect_pr_format(ws)
    hdrs=[str(v or "").strip() for v in rows[hr]]; hl=[h.lower() for h in hdrs]
    out=defaultdict(lambda:{"party":"","gstin":"","igst":0.,"cgst":0.,"sgst":0.,"total":0.,"category":"Regular B2B","inv_count":0})
    details=[]
    def fc(keys):
        for k in keys:
            for i,h in enumerate(hl):
                if k==h or (len(k)>4 and k in h): return i
        return None

    if fmt=="C":
        cgst_c=[i for i,h in enumerate(hl) if "cgst" in h]
        sgst_c=[i for i,h in enumerate(hl) if "sgst" in h]
        igst_c=[i for i,h in enumerate(hl) if "igst" in h]
        for row in rows[hr+1:]:
            if not row[0]: continue
            p=str(row[0]).strip()
            if p.lower() in ("grand total","total",""): continue
            cgst=sum(safe(row[c]) for c in cgst_c if c<len(row))
            sgst=sum(safe(row[c]) for c in sgst_c if c<len(row))
            igst=sum(safe(row[c]) for c in igst_c if c<len(row))
            if igst==cgst==sgst==0: continue
            k=cname(p)
            out[k]["party"]=p; out[k]["igst"]+=igst; out[k]["cgst"]+=cgst
            out[k]["sgst"]+=sgst; out[k]["total"]+=igst+cgst+sgst; out[k]["inv_count"]+=1
            details.append({"Party":p,"IGST":igst,"CGST":cgst,"SGST":sgst,"Total Tax":igst+cgst+sgst})
        if out: return dict(out),details

    if fmt=="A":
        ci={"date":fc(["date"]),"party":fc(["particulars","party","vendor","supplier"]),
            "vch":fc(["voucher no.","voucher no","invoice no"]),"sinv":fc(["supplier invoice no.","supplier invoice no"]),
            "gstin":fc(["gstin/uin","gstin"])}
        igst_cols=[i for i,h in enumerate(hl) if "igst" in h and "input" in h]
        cgst_cols=[i for i,h in enumerate(hl) if "cgst" in h and "input" in h]
        sgst_cols=[i for i,h in enumerate(hl) if "sgst" in h and "input" in h]
        # Confirm format A by checking for numeric tax values in first 150 rows
        has_v=any(rows[hr+j][c] for j in range(min(150,len(rows)-hr)) for c in igst_cols+cgst_cols if c<len(rows[hr+j]) and isinstance(rows[hr+j][c],(int,float)) and rows[hr+j][c]>0)
        if has_v:
            for row in rows[hr+1:]:
                if not any(row): continue
                dt=row[ci["date"]] if ci["date"] is not None and ci["date"]<len(row) else None
                if not isinstance(dt,(datetime.datetime,datetime.date)): continue
                p=str(row[ci["party"]] if ci["party"] is not None and ci["party"]<len(row) else "").strip()
                if not p or p.lower() in ("grand total","total",""): continue
                g=cgstin(row[ci["gstin"]]) if ci["gstin"] is not None and ci["gstin"]<len(row) else ""
                vch=str(row[ci["vch"]] if ci["vch"] is not None and ci["vch"]<len(row) else "").strip()
                sinv=str(row[ci["sinv"]] if ci["sinv"] is not None and ci["sinv"]<len(row) else "").strip()
                igst=sum(safe(row[c]) for c in igst_cols if c<len(row))
                cgst=sum(safe(row[c]) for c in cgst_cols if c<len(row))
                sgst=sum(safe(row[c]) for c in sgst_cols if c<len(row))
                cust_amt=sum(safe(row[c]) for c in [i for i,h in enumerate(hl) if "igst" in h and "custom" in h] if c<len(row))
                cat="Custom Duty / Import" if cust_amt>0 else ("No ITC / RCM" if igst==cgst==sgst==0 else "Regular B2B")
                k=g if g else cname(p)
                out[k]["party"]=p; out[k]["gstin"]=g; out[k]["igst"]+=igst; out[k]["cgst"]+=cgst
                out[k]["sgst"]+=sgst; out[k]["total"]+=igst+cgst+sgst; out[k]["category"]=cat; out[k]["inv_count"]+=1
                details.append({"Date":dt,"Party":p,"GSTIN":g,"VchNo":vch,"SupplierInv":sinv,
                                 "IGST":igst,"CGST":cgst,"SGST":sgst,"Total Tax":igst+cgst+sgst,"Category":cat})
            if out: return dict(out),details

    # FORMAT B: sub-row
    cur={}; ig=cg=sg=0.
    def save_b():
        if cur.get("p"):
            cust=cur.get("cust",0)
            cat="Custom Duty / Import" if cust>0 else ("No ITC / RCM" if ig==cg==sg==0 else "Regular B2B")
            k=cur.get("g","") or cname(cur["p"])
            out[k]["party"]=cur["p"]; out[k]["gstin"]=cur.get("g","")
            out[k]["igst"]+=ig; out[k]["cgst"]+=cg; out[k]["sgst"]+=sg
            out[k]["total"]+=ig+cg+sg; out[k]["category"]=cat; out[k]["inv_count"]+=1
            details.append({"Date":cur.get("d"),"Party":cur["p"],"GSTIN":cur.get("g",""),
                             "VchNo":cur.get("v",""),"SupplierInv":cur.get("si",""),
                             "IGST":ig,"CGST":cg,"SGST":sg,"Total Tax":ig+cg+sg,"Category":cat})
    for row in rows[hr+1:]:
        if not any(row): continue
        r=[row[i] if i<len(row) else None for i in range(max(8,len(row)))]
        dt,part,vn,si,_,_,db,_=r[:8]
        if isinstance(dt,(datetime.datetime,datetime.date)):
            save_b(); p=str(part or "").strip()
            if not p or p.lower() in ("total","grand total"): cur={}; ig=cg=sg=0.; continue
            cur={"d":dt,"p":p,"v":str(vn or ""),"si":str(si or ""),"g":""}; ig=cg=sg=0.
        elif part and cur:
            lbl=str(part).upper(); amt=safe(db)
            if "IGST" in lbl and "CUSTOM" in lbl: cur["cust"]=cur.get("cust",0)+amt; ig+=amt
            elif "IGST" in lbl: ig+=amt
            elif "CGST" in lbl: cg+=amt
            elif "SGST" in lbl or "UTGST" in lbl: sg+=amt
            elif "GSTIN" in lbl and db: cur["g"]=cgstin(db)
    save_b()
    if not out: raise ValueError("No ITC data found in purchase file.")
    return dict(out),details

def parse_2b(fb):
    wb=load_wb(fb)
    if "B2B" not in wb.sheetnames:
        raise ValueError("2B file must have 'B2B' sheet.\nDownload Excel from: GST Portal → Returns → GSTR-2B → Download Excel")
    ws=wb["B2B"]
    data_start=7
    for i,row in enumerate(ws.iter_rows(min_row=4,max_row=10,values_only=True),4):
        if row[0] and re.match(r'^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][0-9][A-Z][0-9A-Z]$',str(row[0]).strip()):
            data_start=i; break
    out=defaultdict(lambda:{"party":"","gstin":"","igst":0.,"cgst":0.,"sgst":0.,"total":0.,"rcm":False,"inv_count":0})
    for row in ws.iter_rows(min_row=data_start,max_row=ws.max_row,values_only=True):
        if not row[0] or not str(row[0]).strip(): continue
        g=cgstin(row[0]); nm=str(row[1] or "").strip()
        igst=safe(row[9]); cgst=safe(row[10]); sgst=safe(row[11])
        rcm=str(row[7] or "").upper()=="YES"
        k=g if g else cname(nm)
        out[k]["party"]=nm; out[k]["gstin"]=g; out[k]["igst"]+=igst; out[k]["cgst"]+=cgst
        out[k]["sgst"]+=sgst; out[k]["total"]+=igst+cgst+sgst; out[k]["rcm"]|=rcm; out[k]["inv_count"]+=1
    impg_total=0.; impg_rows=[]
    for sn in ["IMPG","IMPGSEZ"]:
        if sn in wb.sheetnames:
            for row in wb[sn].iter_rows(min_row=7,max_row=wb[sn].max_row,values_only=True):
                if not row[0]: continue
                v=safe(row[5]); impg_total+=v
                impg_rows.append({"Ref Date":str(row[0]),"Port":str(row[1] or ""),"BE No":str(row[2] or ""),
                                   "BE Date":str(row[3] or ""),"Taxable":safe(row[4]),"IGST":v})
    return dict(out),impg_total,impg_rows

def match_parties(books,b2b):
    GENERIC={"transport","carrier","carriers","roadlines","logistics","trading","enterprises","services","industries","pvt","ltd","private","limited"}
    gstin_idx={d["gstin"]:k for k,d in b2b.items() if d.get("gstin")}
    b2b_w={gk:words(gd.get("party","") or "") for gk,gd in b2b.items()}
    mapping={}
    for bk,bd in books.items():
        g=bd.get("gstin","")
        if g and g in gstin_idx: mapping[bk]=gstin_idx[g]; continue
        bw=words(bd.get("party","") or bk); best_sc,best_k=0.,None
        for gk,gd in b2b.items():
            gw=b2b_w[gk]
            if not bw or not gw: continue
            ol=bw&gw
            if not ol: continue
            if len(ol-GENERIC)==0 and len(ol)<=1: continue
            sc=len(ol)/len(bw|gw)
            if sc>best_sc: best_sc,best_k=sc,gk
        if best_k and best_sc>=0.25: mapping[bk]=best_k
    return mapping

def run_itc_reco(pr_bytes,b2b_bytes):
    b2b_dict,impg_total,impg_rows=parse_2b(b2b_bytes)
    books_dict,details=parse_purchase(pr_bytes)
    mapping=match_parties(books_dict,b2b_dict)
    used_b2b=set(mapping.values())
    # Consolidate books entries mapping to same 2B party
    cons=defaultdict(lambda:{"party":"","gstin":"","igst":0.,"cgst":0.,"sgst":0.,"inv_count":0,"category":"Regular B2B","names":[]})
    unmatched=[]
    for bk,bd in books_dict.items():
        gk=mapping.get(bk)
        if not gk: unmatched.append((bk,bd)); continue
        cons[gk]["igst"]+=bd["igst"]; cons[gk]["cgst"]+=bd["cgst"]; cons[gk]["sgst"]+=bd["sgst"]
        cons[gk]["inv_count"]+=bd["inv_count"]; cons[gk]["names"].append(bd["party"])
        if not cons[gk]["party"]: cons[gk]["party"]=bd["party"]; cons[gk]["gstin"]=bd.get("gstin",""); cons[gk]["category"]=bd.get("category","Regular B2B")
    rows=[]
    def st_rm(t_i,t_c,t_s,g_i,g_c,g_s,cat,rcm,has_match):
        if "Custom" in cat: return "Custom Duty / Import","Custom duty/Import — verify IMPG in 2B; ITC via ICEGATE Sec 4(A)(1)"
        if rcm: return "RCM — Verify Books","RCM — raise self-invoice, pay tax in 3B Sec 3.1(d), claim Sec 4(A)(3)"
        if not has_match: 
            if "Import" in cat: return "Only in Books","Import — may reflect in 2B next month (ICEGATE delay)"
            return "Only in Books","Booked in Tally but supplier not filed GSTR-1 — ITC blocked Sec 16(2)(c)"
        di,dc,ds=t_i-g_i,t_c-g_c,t_s-g_s
        if abs(di)<1 and abs(dc)<1 and abs(ds)<1: return "Matched","✓ ITC claimable — books match 2B"
        parts=[]
        if abs(di)>1: parts.append(f"IGST diff ₹{abs(di):,.0f} ({'books>2B' if di>0 else '2B>books'})")
        if abs(dc)>1: parts.append(f"CGST diff ₹{abs(dc):,.0f}")
        if abs(ds)>1: parts.append(f"SGST diff ₹{abs(ds):,.0f}")
        hint="Prior-month invoices booked this period" if (t_i+t_c+t_s)>(g_i+g_c+g_s) else "Invoices pending booking"
        return "Mismatch",("; ".join(parts)+f" — {hint}") if parts else hint
    for gk,cd in cons.items():
        gd=b2b_dict.get(gk,{})
        tp=cd["party"]; gp=gd.get("party","")
        if len(cd["names"])>1: tp=cd["names"][0]+f" (+{len(cd['names'])-1} more)"
        st,rm=st_rm(cd["igst"],cd["cgst"],cd["sgst"],gd.get("igst",0),gd.get("cgst",0),gd.get("sgst",0),cd["category"],gd.get("rcm",False),True)
        rows.append({"T_Party":tp,"T_GSTIN":cd.get("gstin",""),"G_Party":gp,"G_GSTIN":gd.get("gstin",""),
                     "T_Inv":cd["inv_count"],"T_IGST":cd["igst"],"T_CGST":cd["cgst"],"T_SGST":cd["sgst"],"T_Tax":cd["igst"]+cd["cgst"]+cd["sgst"],
                     "G_Inv":gd.get("inv_count",0),"G_IGST":gd.get("igst",0),"G_CGST":gd.get("cgst",0),"G_SGST":gd.get("sgst",0),"G_Tax":gd.get("total",0),
                     "D_IGST":cd["igst"]-gd.get("igst",0),"D_CGST":cd["cgst"]-gd.get("cgst",0),"D_Tax":cd["igst"]+cd["cgst"]+cd["sgst"]-gd.get("total",0),
                     "Status":st,"Remark":rm,"Category":cd["category"]})
    for bk,bd in unmatched:
        cat=bd.get("category","Regular B2B")
        st,rm=st_rm(bd["igst"],bd["cgst"],bd["sgst"],0,0,0,cat,False,False)
        rows.append({"T_Party":bd["party"],"T_GSTIN":bd.get("gstin",""),"G_Party":"","G_GSTIN":"",
                     "T_Inv":bd["inv_count"],"T_IGST":bd["igst"],"T_CGST":bd["cgst"],"T_SGST":bd["sgst"],"T_Tax":bd["igst"]+bd["cgst"]+bd["sgst"],
                     "G_Inv":0,"G_IGST":0,"G_CGST":0,"G_SGST":0,"G_Tax":0,
                     "D_IGST":bd["igst"],"D_CGST":bd["cgst"],"D_Tax":bd["igst"]+bd["cgst"]+bd["sgst"],
                     "Status":st,"Remark":rm,"Category":cat})
    for gk,gd in b2b_dict.items():
        if gk in used_b2b: continue
        st="RCM — Verify Books" if gd.get("rcm",False) else "Only in 2B"
        rm="RCM — raise self-invoice, pay in 3B" if gd.get("rcm",False) else "Supplier filed but not in Tally — check if purchase to be booked"
        rows.append({"T_Party":"","T_GSTIN":"","G_Party":gd["party"],"G_GSTIN":gd.get("gstin",""),
                     "T_Inv":0,"T_IGST":0,"T_CGST":0,"T_SGST":0,"T_Tax":0,
                     "G_Inv":gd["inv_count"],"G_IGST":gd["igst"],"G_CGST":gd["cgst"],"G_SGST":gd["sgst"],"G_Tax":gd["total"],
                     "D_IGST":-gd["igst"],"D_CGST":-gd["cgst"],"D_Tax":-gd["total"],
                     "Status":st,"Remark":rm,"Category":"Regular B2B"})
    df=pd.DataFrame(rows)
    ORDER={"Matched":0,"Mismatch":1,"RCM — Verify Books":2,"Custom Duty / Import":3,"Only in 2B":4,"Only in Books":5}
    df["_s"]=df["Status"].map(ORDER).fillna(9)
    return df.sort_values(["_s","T_Party"]).drop(columns=["_s"]).reset_index(drop=True),details,impg_total,impg_rows

def build_itc_excel(df,details,impg_rows,impg_total,client,period):
    wb=openpyxl.Workbook(); cnt=df["Status"].value_counts().to_dict()
    ITC_BG={"Matched":MATCH,"Mismatch":MIS,"RCM — Verify Books":RCM_C,"Custom Duty / Import":CUST_C,"Only in 2B":ONLY2,"Only in Books":ONLY1}
    ITC_FC={"Matched":"1E5631","Mismatch":"C00000","RCM — Verify Books":"2E7D32","Custom Duty / Import":"7F4F00","Only in 2B":"7A5C00","Only in Books":"0D3055"}

    def mhdr(ws,txt,nc,row=1):
        ws.merge_cells(f"A{row}:{get_column_letter(nc)}{row}")
        c=ws.cell(row,1); c.value=txt; c.font=Font(name="Calibri",bold=True,size=12,color="FFFFFF")
        c.fill=F(TITLE); c.alignment=CTR(); ws.row_dimensions[row].height=28
    def sub(ws,txt,nc,row=2):
        ws.merge_cells(f"A{row}:{get_column_letter(nc)}{row}")
        c=ws.cell(row,1); c.value=f"Created by: Purva Doshi  |  linkedin.com/in/purvadoshi26  |  {txt}"
        c.font=Font(name="Calibri",size=8,italic=True,color="7F9FC0"); c.alignment=CTR(); ws.row_dimensions[row].height=14

    ws1=wb.active; ws1.title="Summary"; ws1.sheet_view.showGridLines=False
    mhdr(ws1,f"{client}  —  GSTR-2B vs Tally ITC Reco  |  {period}",7)
    sub(ws1,f"ITC Reco | {period}",7)
    for ci,v in enumerate(["Status","Tally (Books)","GSTR-2B","Diff IGST","Diff CGST+SGST","Action","Ref"],1):
        c=ws1.cell(4,ci); c.value=v; c.font=Font(name="Calibri",bold=True,size=9,color=NAVY)
        c.fill=F("E8EFF8"); c.alignment=CTR(); c.border=BD()
    sr=[("✅ Matched",cnt.get("Matched",0),cnt.get("Matched",0),0,0,"ITC claimable — verify Sec 16","Sec 16"),
        ("⚠️ Mismatch",cnt.get("Mismatch",0),cnt.get("Mismatch",0),"—","—","Investigate before claiming","Sec 16(2)"),
        ("🔵 RCM",cnt.get("RCM — Verify Books",0),0,0,0,"Self-invoice + pay in 3B","Sec 9(3)"),
        ("🟠 Custom/Import",cnt.get("Custom Duty / Import",0),cnt.get("Custom Duty / Import",0),"—","—","Verify IMPG in 2B","Sec 4(A)(1)"),
        ("🟡 Only in 2B",0,cnt.get("Only in 2B",0),"—","—","Book purchase in Tally","Sec 16(2)(b)"),
        ("⚪ Only in Books",cnt.get("Only in Books",0),0,"—","—","Supplier GSTR-1 pending","Sec 16(2)(c)")]
    for ri,rd in enumerate(sr,5):
        for ci,val in enumerate(rd,1):
            c=ws1.cell(ri,ci); c.value=val; c.border=BD(); c.font=DF(9,bold=(ci==1))
            if ci in(2,3,4,5) and isinstance(val,(int,float)): c.number_format="#,##0"; c.alignment=RGT()
            else: c.alignment=LFT()
        ws1.row_dimensions[ri].height=17
    # ITC amounts
    ws1.merge_cells("A12:G12"); c=ws1.cell(12,1); c.value="ITC AMOUNTS (₹)"
    c.font=HF(10,NAVY); c.fill=F("E8EFF8"); c.alignment=CTR(); c.border=BD()
    for ci,v in enumerate(["","2B IGST","2B CGST","2B SGST","Books IGST","Books CGST","Books SGST"],1):
        c=ws1.cell(13,ci); c.value=v; c.font=HF(8); c.fill=F(NAVY); c.alignment=CTR(); c.border=BD()
    for ri,(lbl,b2i,b2c,b2s,bki,bkc,bks) in enumerate([
        ("Total",df["G_IGST"].sum(),df["G_CGST"].sum(),df["G_SGST"].sum(),df["T_IGST"].sum(),df["T_CGST"].sum(),df["T_SGST"].sum()),
        ("✅ Claimable",df[df.Status=="Matched"]["G_IGST"].sum(),df[df.Status=="Matched"]["G_CGST"].sum(),df[df.Status=="Matched"]["G_SGST"].sum(),
         df[df.Status=="Matched"]["T_IGST"].sum(),df[df.Status=="Matched"]["T_CGST"].sum(),df[df.Status=="Matched"]["T_SGST"].sum()),
        ("IMPG (Imports)",impg_total,0,0,"—","—","—")
    ],14):
        for ci,val in enumerate([lbl,b2i,b2c,b2s,bki,bkc,bks],1):
            c=ws1.cell(ri,ci); c.value=val; c.border=BD(); c.font=DF(9,bold=(ci==1))
            if ci in(2,3,4,5,6,7) and isinstance(val,(int,float)): c.number_format=INR; c.alignment=RGT()
            else: c.alignment=LFT()
    for col,w in zip("ABCDEFG",[30,14,14,14,14,14,14]): ws1.column_dimensions[col].width=w

    # Auditor Actions sheet
    ws_act=wb.create_sheet("Auditor Actions"); ws_act.sheet_view.showGridLines=False
    mhdr(ws_act,f"AUDITOR ACTION ITEMS — {client} | {period}",5)
    ws_act.merge_cells("A2:E2"); c=ws_act.cell(2,1)
    c.value="Focus on material items first. Minor differences (<₹1,000) are generally not significant. RCM & blocked ITC impact cash flow directly."
    c.font=Font(name="Calibri",size=8,italic=True,color="555555"); c.alignment=LFT(); ws_act.row_dimensions[2].height=14
    for ci,v in enumerate(["Priority","Action Required","Detail","Amount (₹)","Legal Ref"],1):
        c=ws_act.cell(3,ci); c.value=v; c.font=HF(9); c.fill=F(NAVY); c.alignment=CTR(); c.border=BD()
    actions=[]
    mat_mis=df[(df.Status=="Mismatch") & (df["D_Tax"].abs()>10000)]
    for _,r in mat_mis.iterrows():
        actions.append(("🔴 HIGH",f"Mismatch — {(r['T_Party'] or r['G_Party'])[:40]}",r["Remark"][:80],f"₹{abs(r['D_Tax']):,.0f}","Sec 16(2)"))
    ob=df[df.Status=="Only in 2B"]
    if len(ob)>0: actions.append(("🔴 HIGH",f"Book {len(ob)} pending supplier invoices in Tally","ITC available in 2B but not booked — book before 3B filing",f"₹{ob['G_Tax'].sum():,.0f}","Sec 16(2)(b)"))
    nbk=df[(df.Status=="Only in Books")&(df["Category"]!="Custom Duty / Import")]
    if len(nbk)>0: actions.append(("🟠 MED",f"Follow up with {len(nbk)} suppliers for GSTR-1 filing","ITC blocked — reverse if supplier doesn't file within time limit",f"₹{nbk['T_Tax'].sum():,.0f}","Sec 16(2)(c)"))
    rcm=df[df.Status=="RCM — Verify Books"]
    if len(rcm)>0: actions.append(("🟠 MED",f"Verify RCM — {len(rcm)} parties","Self-invoice raised? Tax paid in 3B Sec 3.1(d)? Claimed in Sec 4(A)(3)?",f"₹{rcm['G_Tax'].sum():,.0f}","Sec 9(3)"))
    if impg_total>0: actions.append(("🟡 LOW","Reconcile IMPG with ICEGATE",f"IGST on imports. Verify Bill of Entry in 2B IMPG sheet.",f"₹{impg_total:,.0f}","Sec 4(A)(1)"))
    mi=df[df.Status=="Mismatch"]; minor_d=mi[mi["D_Tax"].abs()<=10000]
    if len(minor_d)>0: actions.append(("🟡 LOW",f"{len(minor_d)} minor mismatches (≤₹10,000)","Likely rounding/partial invoice. Verify if material in aggregate.",f"₹{minor_d['D_Tax'].abs().sum():,.0f}","Materiality"))
    actions.append(("✅ OK",f"Matched ITC — ready to claim (verify Sec 16 conditions)",f"{cnt.get('Matched',0)} parties, IGST+CGST+SGST claimable after Sec 16 check",f"₹{df[df.Status=='Matched']['G_Tax'].sum():,.0f}","Sec 16"))
    pri_bg={"🔴 HIGH":MIS,"🟠 MED":CUST_C,"🟡 LOW":ONLY2,"✅ OK":MATCH}
    for ri,item in enumerate(actions,4):
        bg=pri_bg.get(item[0][:6] if len(item[0])>5 else item[0],"FFFFFF")
        for ci,(val,w) in enumerate(zip(item,[14,38,60,16,12]),1):
            c=ws_act.cell(ri,ci); c.value=val; c.border=BD()
            c.fill=F(bg if ci<=2 else "FFFFFF"); c.font=DF(9,bold=(ci<=2))
            c.alignment=CTR() if ci in(1,4,5) else LFT()
        ws_act.row_dimensions[ri].height=26
    nr=len(actions)+5; ws_act.merge_cells(f"A{nr}:E{nr}"); c=ws_act.cell(nr,1)
    c.value="NOTE: Verify Sec 16(4) time limit compliance. Sec 17(5) blocked credits (vehicles, food, club) must not be claimed. Rule 42 reversal for exempt/personal use."
    c.font=Font(name="Calibri",size=8,italic=True,color="666666"); c.alignment=LFT()
    for col,w in zip("ABCDE",[14,38,60,16,12]): ws_act.column_dimensions[col].width=w

    # ITC Reco detail
    ws2=wb.create_sheet("ITC Reco"); ws2.sheet_view.showGridLines=False
    mhdr(ws2,f"{client}  —  GSTR-2B vs Tally  |  {period}",16); sub(ws2,f"ITC Reco | {period}",16)
    rc=[("#",4),("Party (Books)",34),("Party (2B)",32),("GSTIN",20),
        ("Tally #",8),("Tally IGST",15),("Tally CGST",15),("Tally SGST",15),("Tally Total",15),
        ("2B #",8),("2B IGST",15),("2B CGST",15),("2B SGST",15),("2B Total",15),("Status",22),("Remarks",55)]
    for ci,(hdr,w) in enumerate(rc,1):
        c=ws2.cell(4,ci); c.value=hdr; c.font=HF(8); c.fill=F(NAVY); c.alignment=CTR(); c.border=BD()
        ws2.column_dimensions[get_column_letter(ci)].width=w
    ws2.row_dimensions[4].height=24
    for ri,row in df.iterrows():
        r=ri+5; st=row["Status"]; bg=ITC_BG.get(st,"FFFFFF")
        gstin=str(row.get("T_GSTIN") or row.get("G_GSTIN") or "")
        vals=[ri+1,row["T_Party"],row["G_Party"],gstin,int(row["T_Inv"]),
              row["T_IGST"],row["T_CGST"],row["T_SGST"],row["T_Tax"],
              int(row["G_Inv"]),row["G_IGST"],row["G_CGST"],row["G_SGST"],row["G_Tax"],st,row["Remark"]]
        for ci,val in enumerate(vals,1):
            c=ws2.cell(r,ci); c.value=val; c.fill=F(bg); c.border=BD()
            if ci in(6,7,8,9,11,12,13,14): c.font=DF(9); c.alignment=RGT(); c.number_format=INR
            elif ci in(5,10): c.font=DF(9); c.alignment=CTR(); c.number_format=NUM
            elif ci==15: c.font=Font(name="Calibri",bold=True,size=9,color=ITC_FC.get(st,"222")); c.alignment=CTR()
            else: c.font=DF(9); c.alignment=LFT()
        ws2.row_dimensions[r].height=15
    tr=len(df)+5; ws2.merge_cells(f"A{tr}:D{tr}"); tc=ws2.cell(tr,1); tc.value="TOTAL"
    tc.font=HF(10); tc.fill=F(NAVY); tc.alignment=CTR(); tc.border=BD()
    for ci in range(5,15):
        c=ws2.cell(tr,ci); c.font=HF(9); c.fill=F(NAVY); c.border=BD()
        if ci in(6,7,8,9,11,12,13,14): c.value=f"=SUM({get_column_letter(ci)}5:{get_column_letter(ci)}{tr-1})"; c.number_format=INR; c.alignment=RGT()
        elif ci in(5,10): c.value=f"=SUM({get_column_letter(ci)}5:{get_column_letter(ci)}{tr-1})"; c.number_format=NUM; c.alignment=CTR()
    ws2.freeze_panes="E5"

    # Tally detail
    def det_tab(wb,df_in,cols,title,hc):
        ws=wb.create_sheet(title[:31]); ws.sheet_view.showGridLines=False
        ws.merge_cells(f"A1:{get_column_letter(len(cols))}1"); h=ws.cell(1,1); h.value=title
        h.font=Font(name="Calibri",bold=True,size=11,color="FFFFFF"); h.fill=F(hc); h.alignment=CTR()
        for ci,(hdr,w,fld) in enumerate(cols,1):
            c=ws.cell(2,ci); c.value=hdr; c.font=HF(9); c.fill=F(hc); c.alignment=CTR(); c.border=BD()
            ws.column_dimensions[get_column_letter(ci)].width=w
        for ri,row in df_in.iterrows():
            r=ri+3
            for ci,(_,_,fld) in enumerate(cols,1):
                c=ws.cell(r,ci); c.value=row.get(fld,""); c.border=BD(); c.font=DF(9)
                if fld in("Date","InvDate"): c.number_format="DD-MMM-YY"; c.alignment=CTR()
                elif fld in("IGST","CGST","SGST","Total Tax","Taxable","igst","taxable","IMPG IGST","G_Tax"): c.number_format=INR; c.alignment=RGT()
                else: c.alignment=LFT()
            ws.row_dimensions[r].height=14
        ws.freeze_panes="A3"

    if details:
        df_det=pd.DataFrame(details)
        for col in ["Date","Party","GSTIN","VchNo","SupplierInv","IGST","CGST","SGST","Total Tax","Category"]:
            if col not in df_det.columns: df_det[col]=""
        det_tab(wb,df_det,[("Date",12,"Date"),("Party",36,"Party"),("GSTIN",20,"GSTIN"),
                           ("Voucher No",20,"VchNo"),("Supplier Inv",20,"SupplierInv"),
                           ("IGST",15,"IGST"),("CGST",15,"CGST"),("SGST",15,"SGST"),("Total Tax",15,"Total Tax"),("Category",22,"Category")],
                f"Tally Detail | {period}",BLUE)
    if impg_rows:
        df_impg=pd.DataFrame(impg_rows)
        det_tab(wb,df_impg,[("Ref Date",14,"Ref Date"),("Port",10,"Port"),("BE No",18,"BE No"),
                            ("BE Date",12,"BE Date"),("Taxable",16,"Taxable"),("IGST",15,"IGST")],
                "IMPG Imports",TEAL)

    # Legend
    wl=wb.create_sheet("Legend"); wl.sheet_view.showGridLines=False
    wl.merge_cells("A1:D1"); h=wl.cell(1,1); h.value="Color Legend & Sec 16 Audit Checklist"
    h.font=Font(name="Calibri",bold=True,size=11,color="FFFFFF"); h.fill=F(TITLE); h.alignment=CTR()
    wl.merge_cells("A2:D2"); c=wl.cell(2,1); c.value="Created by: Purva Doshi  |  linkedin.com/in/purvadoshi26"
    c.font=Font(name="Calibri",size=9,italic=True,color="5585B5"); c.alignment=CTR()
    for i,(bg,st,desc) in enumerate([(MATCH,"Matched","ITC claimable — verify Sec 16 conditions"),(MIS,"Mismatch","Tax diff — investigate"),(RCM_C,"RCM","Self-invoice + pay in GSTR-3B"),(CUST_C,"Custom/Import","IMPG — verify via ICEGATE"),(ONLY2,"Only in 2B","Book purchase — ITC available"),(ONLY1,"Only in Books","Supplier not filed — ITC blocked")],4):
        wl.cell(i,1).value=st; wl.cell(i,1).fill=F(bg); wl.cell(i,1).font=DF(9,bold=True); wl.cell(i,1).border=BD()
        wl.merge_cells(f"B{i}:D{i}"); wl.cell(i,2).value=desc; wl.cell(i,2).fill=F(bg); wl.cell(i,2).font=DF(9); wl.cell(i,2).border=BD(); wl.cell(i,2).alignment=LFT()
    wl.merge_cells("A11:D11"); c=wl.cell(11,1); c.value="SEC 16 AUDIT CHECKLIST"
    c.font=HF(10,NAVY); c.fill=F("E8EFF8"); c.alignment=CTR(); c.border=BD()
    for i,(ico,sec,txt) in enumerate([("✅","Sec 16(2)(a)","Tax invoice received"),("✅","Sec 16(2)(b)","Goods/services received"),("✅","Sec 16(2)(c)","Tax paid by supplier — confirmed via 2B"),("✅","Sec 16(2)(d)","GSTR-3B filed"),("⚠️","Sec 16(4)","ITC within time limit (Sep return / annual return)"),("⚠️","Rule 42","Reversal for exempt + personal use"),("⚠️","Sec 17(5)","Blocked credits not claimed (vehicles, food, club)")],12):
        wl.cell(i,1).value=f"{ico} {sec}"; wl.cell(i,1).font=DF(9,bold=True); wl.cell(i,1).border=BD()
        wl.merge_cells(f"B{i}:D{i}"); wl.cell(i,2).value=txt; wl.cell(i,2).font=DF(9); wl.cell(i,2).border=BD()
    for col,w in zip("ABCD",[20,20,20,28]): wl.column_dimensions[col].width=w
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

# ─────────────────────────────────────────────────────────────────────────────
#  TAB 2 — SALES RECO (E-Invoice / GSTR-1 vs Books)  — header-driven engine
#  Works for any company: columns are found by HEADER NAME, never by position.
#  Ledger columns in the books register are auto-classified (and user-overridable).
# ─────────────────────────────────────────────────────────────────────────────
import numpy as np

def _hn(x):
    return re.sub(r"\s+"," ",str(x if x is not None else "").replace("\n"," ").replace("\r"," ").replace("\t"," ")).strip().lower()

def num(v):
    if v is None: return 0.0
    if isinstance(v,bool): return 0.0
    if isinstance(v,(int,float)): return 0.0 if v!=v else float(v)
    t=str(v).strip().replace(",","").replace("₹","").replace("Rs.","").strip()
    if t in ("","-","--","nan","None"): return 0.0
    neg=t.startswith("(") and t.endswith(")")
    t=t.strip("()").strip()
    try: x=float(t)
    except: return 0.0
    return -x if neg else x

def is_num(v):
    if isinstance(v,bool) or v is None: return False
    if isinstance(v,(int,float)): return v==v
    t=str(v).strip().replace(",","").strip("()")
    if not t: return False
    try: float(t); return True
    except: return False

def to_date(v):
    if v is None or v=="": return None
    if isinstance(v,datetime.datetime): return v.date()
    if isinstance(v,datetime.date): return v
    if isinstance(v,(int,float)) and not isinstance(v,bool) and 30000<v<70000:
        return (datetime.datetime(1899,12,30)+datetime.timedelta(days=float(v))).date()
    t=str(v).strip()
    for fmt in ("%d-%b-%Y","%d-%b-%y","%d/%m/%Y","%d-%m-%Y","%d.%m.%Y","%Y-%m-%d","%d/%m/%y","%d-%m-%y","%d %b %Y","%Y-%m-%d %H:%M:%S","%d-%b-%Y %H:%M:%S"):
        try: return datetime.datetime.strptime(t,fmt).date()
        except: pass
    try:
        d=pd.to_datetime(t,dayfirst=True,errors="coerce")
        return None if pd.isna(d) else d.date()
    except: return None

def inorm(s): return re.sub(r"[^A-Z0-9]","",str(s or "").upper())

# ── header dictionaries (regex on lower-cased header) ────────────────────────
BK_COLS={
 "date":  [r"^date$",r"^(invoice|bill|voucher|doc(ument)?)\.? ?date$"],
 "party": [r"^particulars$",r"^party( name)?$",r"^customer( name)?$",r"^buyer( name)?$",r"^name of (the )?(party|customer)$",r"^ledger( name)?$"],
 "vtype": [r"^voucher type$",r"^vch\.? type$",r"^doc(ument)? type$",r"^type$"],
 "inv":   [r"^voucher (no|number)\.?$",r"^invoice (no|number)\.?$",r"^inv\.? ?(no|number)\.?$",r"^bill (no|number)\.?$",r"^doc(ument)? (no|number)\.?$",r"^vch\.? ?no\.?$",r"^(voucher|invoice) ?(no|number)"],
 "gstin": [r"^gstin(/uin)?$",r"^gst ?(no|number|in)\.?$",r"^(customer|party|buyer|recipient) gstin(/uin)?$",r"gstin"],
 "gross": [r"^gross total$",r"^invoice (value|amount|total)$",r"^total( amount| value)?$",r"^bill (value|amount)$",r"^gross (amount|value)$",r"^grand total$"],
 "narr":  [r"^narration$",r"^remarks?$"],
}
INFO_RX=[r"^date$",r"^particulars$",r"^buyer",r"^voucher",r"^vch",r"gstin",r"^gst ?no",r"^pan( no)?\.?$",r"^narration$",r"^order (no|number)",
 r"^terms of ",r"^other references?$",r"^place of (receipt|supply|delivery)",r"^vessel",r"^port (code|of)",r"^country",r"^shipping (no|date|bill)",
 r"^qty\b",r"^quantity",r"^rate( per.*)?$",r"^value$",r"^gross",r"^total$",r"^grand total$",r"^s\.? ?no\.?$",r"^sr\.? ?no\.?$",r"^sl\.? ?no\.?$",r"^sr\.?$",
 r"address",r"^state( name)?$",r"^hsn",r"^uqc$",r"^unit$",r"e-?way",r"\birn\b",r"^ack",r"^reference",r"^ref\.?\b",r"^invoice (no|number|date|value|amount)",
 r"^doc(ument)? (no|number|date|type)",r"^supply type",r"^reverse charge",r"^remarks?$",r"^cost cent",r"^currency",r"^exchange"]
def _is_info(h): return any(re.search(p,h) for p in INFO_RX)

LEDGER_CLASSES=["Taxable","IGST","CGST","SGST","Cess","Other","Ignore"]
_VAL_KW=r"sale|sales|charge|income|revenue|freight|handling|storage|service|export|exempt|detention|shortage|supply|turnover|job ?work|labou?r|rent|commission|packing|discount"
def guess_ledger_class(h):
    t=_hn(h)
    if re.search(r"round(ing)?[ -]?(off|up|down)?\b|\btcs\b|\btds\b",t) and not re.search(r"sale",t): return "Other"
    if re.search(r"\bcess\b",t): return "Cess"
    head=None
    if re.search(r"\bigst\b|integrated (gst|tax)",t): head="IGST"
    elif re.search(r"\bcgst\b|central (gst|tax)",t): head="CGST"
    elif re.search(r"\bsgst\b|\butgst\b|state (gst|tax)|\but tax\b|union territory",t): head="SGST"
    if head:
        strong=re.search(r"output|payab|payb|o/p",t) or re.fullmatch(r"(output |o/p )?(i|c|s|ut)gst( ?\(?\d+(\.\d+)?%?\)?)?",t)
        if strong: return head
        if re.search(_VAL_KW,t): return "Taxable"
        return head
    return "Taxable"

def guess_vtype(v):
    t=_hn(v)
    return "Credit/Debit Note" if re.search(r"credit note|cr\.? ?note|debit note|dr\.? ?note|sales? return|\bcn\b|\bdn\b",t) else "Invoice"

def _find_hdr(rows,spec,required,min_hits=2,scan=40):
    best=(None,{})
    for i,row in enumerate(rows[:scan]):
        hs=[_hn(v) for v in row]; cm={}
        for field,pats in spec.items():
            hit=None
            for pat in pats:
                for ci,h in enumerate(hs):
                    if h and ci not in cm.values() and re.search(pat,h): hit=ci; break
                if hit is not None: break
            if hit is not None: cm[field]=hit
        if all(r in cm for r in required) and len(cm)>=min_hits:
            if best[0] is None or len(cm)>len(best[1]): best=(i,cm)
    return best

# ── BOOKS: scan (layout + ledger guesses) ────────────────────────────────────
def scan_books(fb):
    wb=load_wb(fb); recs=[]; layout=[]; warns=[]
    for ws in wb.worksheets:
        rows=[list(r) for r in ws.iter_rows(values_only=True)]
        if not rows: continue
        hr,cm=_find_hdr(rows,BK_COLS,["inv"],2)
        if hr is None:
            warns.append(f"Sheet '{ws.title}' skipped: no header row containing a Voucher/Invoice No column."); continue
        hdr=rows[hr]; data=rows[hr+1:]; used=set(cm.values()); lcols=[]; seen={}
        def colvals(ci): return [r[ci] if ci<len(r) else None for r in data]
        for ci,h in enumerate(hdr):
            hs=_hn(h)
            if not hs or ci in used or _is_info(hs): continue
            if not any(is_num(v) and num(v)!=0 for v in colvals(ci)): continue
            name=str(h).strip()
            if name in seen: seen[name]+=1; name=f"{name} ({seen[name]})"
            else: seen[name]=1
            lcols.append((ci,name))
        if not any(guess_ledger_class(n)=="Taxable" for _,n in lcols):
            for ci,h in enumerate(hdr):
                hs=_hn(h)
                if ci in used or not re.fullmatch(r"(taxable )?value|taxable( amount)?|basic( value| amount)?|amount|assessable value|net amount",hs): continue
                if any(is_num(v) and num(v)!=0 for v in colvals(ci)):
                    lcols.append((ci,ws.title.strip() if hs in("value","amount") else str(h).strip())); break
        n0=len(recs)
        for r in data:
            g=lambda f: (r[cm[f]] if f in cm and cm[f]<len(r) else None)
            inv=str(g("inv") if g("inv") is not None else "").strip()
            if not inv: continue
            party=str(g("party") or "").strip()
            if party.lower() in ("total","grand total","sub total") or inv.lower() in ("total","grand total"): continue
            rec={"_sheet":ws.title,"_inv":inv,"_date":to_date(g("date")),"_party":party,"_gstin":cgstin(g("gstin")),
                 "_vtype":str(g("vtype") or "").strip(),"_gross":(abs(num(g("gross"))) if "gross" in cm else np.nan),"_narr":str(g("narr") or "")}
            for ci,name in lcols: rec[name]=num(r[ci]) if ci<len(r) else 0.0
            recs.append(rec)
        layout.append({"sheet":ws.title,"header_row":hr+1,"rows":len(recs)-n0,
                       "mapped":{k:str(hdr[v]).strip() for k,v in cm.items()},"ledgers":[n for _,n in lcols]})
    if not recs: raise ValueError("No voucher rows found in the books file.\nNeeded headers: Voucher No / Invoice No, Date, GSTIN, and value/tax ledger columns.")
    bw=pd.DataFrame(recs); led=[c for c in bw.columns if not c.startswith("_")]
    for c in led: bw[c]=pd.to_numeric(bw[c],errors="coerce").fillna(0.0)
    if not led: raise ValueError("No value/tax ledger columns detected in the books file (need numeric columns such as sales, IGST, CGST, SGST).")
    return {"bw":bw,"ledger_guess":{c:guess_ledger_class(c) for c in led},
            "vtype_guess":{v:guess_vtype(v) for v in bw["_vtype"].unique()},"layout":layout,"warns":warns}

STD_RATES=[0,0.1,0.25,1,1.5,3,5,6,7.5,12,14,18,28,40]
def build_books(bw,ledger_class,vtype_class):
    led=[c for c in bw.columns if not c.startswith("_")]; d=bw.copy()
    def ssum(cls):
        cols=[c for c in led if ledger_class.get(c,"Taxable")==cls]
        return d[cols].sum(axis=1) if cols else pd.Series(0.0,index=d.index)
    d["Taxable"]=ssum("Taxable"); d["IGST"]=ssum("IGST"); d["CGST"]=ssum("CGST"); d["SGST"]=ssum("SGST"); d["Cess"]=ssum("Cess"); d["OtherL"]=ssum("Other")
    d["vclass"]=d["_vtype"].map(lambda v: vtype_class.get(v,"Invoice"))
    d=d[d.vclass!="Exclude"].copy(); d["inv_norm"]=d["_inv"].map(inorm)
    tot=(d.Taxable.abs()+d.IGST.abs()+d.CGST.abs()+d.SGST.abs()+d.Cess.abs()+d.OtherL.abs()+d[led].abs().sum(axis=1))
    d["_cancel"]=[("cancel" in p.lower()) or (t==0 and (g=="" or "cancel" in n.lower())) for p,t,g,n in zip(d._party,tot,d._gstin,d._narr)]
    dups=d.groupby(["_sheet","inv_norm"]).size(); dups=dups[dups>1]
    g=d.groupby(["vclass","inv_norm"],sort=False).agg({"_inv":"first","_date":"first","_party":"first","_vtype":"first","_gross":"max",
        "Taxable":"sum","IGST":"sum","CGST":"sum","SGST":"sum","Cess":"sum","OtherL":"sum","_cancel":"all",
        "_gstin":lambda x: next((v for v in x if v),"")}).reset_index()
    g["gstin_n"]=g["_gstin"]
    g["calc_gross"]=g.Taxable+g.IGST+g.CGST+g.SGST+g.Cess+g.OtherL
    chk=g[g._gross.notna()&(g._gross>0)].copy(); chk["gdiff"]=chk._gross-chk.calc_gross.abs()
    gross_fail=chk[chk.gdiff.abs()>1]
    r=g[(g.Taxable>0)&((g.IGST+g.CGST+g.SGST)>0)].copy(); r["eff"]=(r.IGST+r.CGST+r.SGST)/r.Taxable*100
    def okrate(x): return any(abs(x-s)<=0.25 for s in STD_RATES)
    rate_fail=r[~r.eff.map(okrate)]
    z=g[(g.Taxable<=0)&((g.IGST+g.CGST+g.SGST)>0)&(~g._cancel)]
    mixed=g[(g.IGST>0)&((g.CGST>0)|(g.SGST>0))]
    return {"inv":g[g.vclass=="Invoice"].reset_index(drop=True),"notes":g[g.vclass=="Credit/Debit Note"].reset_index(drop=True),
            "gross_checked":len(chk),"gross_fail":gross_fail,"rate_fail":pd.concat([rate_fail,z]).drop_duplicates(subset=["vclass","inv_norm"]),
            "mixed":mixed,"dups":dups}

# ── PORTAL ───────────────────────────────────────────────────────────────────
PT_COLS={
 "gstin":[r"^gstin/uin of recipient$",r"^gstin of recipient$",r"^recipient gstin(/uin)?$",r"^gstin/uin of",r"^gstin(/uin)?$"],
 "name": [r"^receiver name$",r"^recipient name$",r"^name of (the )?(recipient|receiver)$",r"^trade/legal name$",r"^name$"],
 "inv":  [r"^invoice (number|no\.?)$",r"^note (number|no\.?)$",r"^document (number|no\.?)$"],
 "date": [r"^invoice date$",r"^note date$",r"^document date$"],
 "value":[r"^invoice value$",r"^note value$",r"^document value$"],
 "taxable":[r"^taxable value$",r"^total taxable value$",r"^taxable amount$"],
 "igst": [r"^integrated tax( amount)?$",r"^igst( amount)?$"],
 "cgst": [r"^central tax( amount)?$",r"^cgst( amount)?$"],
 "sgst": [r"^state/ut tax( amount)?$",r"^state tax$",r"^sgst( amount)?$",r"^state/ut tax"],
 "ntype":[r"^note type$",r"^document type$"],
 "status":[r"^e-?invoice status$"],
}
HSN_COLS={"hsn":[r"^hsn"],"taxable":PT_COLS["taxable"],"igst":[r"^integrated tax",r"^igst"],"cgst":[r"^central tax",r"^cgst"],"sgst":[r"^state/ut tax",r"^state tax",r"^sgst"]}

def parse_portal(fb):
    wb=load_wb(fb); inv=[]; notes=[]; hsn=None; layout=[]; warns=[]; has_tax=False
    for ws in wb.worksheets:
        rows=[list(r) for r in ws.iter_rows(values_only=True)]
        if not rows: continue
        hr,cm=_find_hdr(rows,PT_COLS,["inv","taxable"],3)
        if hr is None:
            hr2,cm2=_find_hdr(rows,HSN_COLS,["hsn","taxable"],3)
            if hr2 is not None:
                t=dict(Taxable=0.0,IGST=0.0,CGST=0.0,SGST=0.0)
                for r in rows[hr2+1:]:
                    if r[cm2["hsn"]] in (None,""): continue
                    t["Taxable"]+=num(r[cm2["taxable"]]); t["IGST"]+=num(r[cm2["igst"]]) if "igst" in cm2 else 0
                    t["CGST"]+=num(r[cm2["cgst"]]) if "cgst" in cm2 else 0; t["SGST"]+=num(r[cm2["sgst"]]) if "sgst" in cm2 else 0
                hsn=t; layout.append({"sheet":ws.title,"role":"HSN summary","header_row":hr2+1,"mapped":{k:str(rows[hr2][v]).strip() for k,v in cm2.items()},"rows":0})
            continue
        hdr=rows[hr]; is_note=_hn(hdr[cm["inv"]]).startswith("note")
        tax_here=all(k in cm for k in ("igst","cgst","sgst")); has_tax=has_tax or tax_here
        n0=len(inv)+len(notes)
        for r in rows[hr+1:]:
            g=lambda f: (r[cm[f]] if f in cm and cm[f]<len(r) else None)
            iv=str(g("inv") if g("inv") is not None else "").strip()
            if not iv: continue
            rec={"inv":iv,"date":to_date(g("date")),"name":str(g("name") or "").strip(),"gstin":cgstin(g("gstin")),
                 "Taxable":num(g("taxable")),"IGST":num(g("igst")),"CGST":num(g("cgst")),"SGST":num(g("sgst")),"value":num(g("value")),
                 "ntype":(str(g("ntype") or "C").strip().upper()[:1] or "C") if is_note else "",
                 "cancelled":"cancel" in str(g("status") or "").lower(),"sheet":ws.title}
            (notes if is_note else inv).append(rec)
        layout.append({"sheet":ws.title,"role":"Credit/Debit notes" if is_note else "Invoices","header_row":hr+1,
                       "mapped":{k:str(hdr[v]).strip() for k,v in cm.items()},"rows":len(inv)+len(notes)-n0})
    if not inv and not notes:
        raise ValueError("Portal file: no sheet found with 'Invoice number' + 'Taxable Value' headers.\nDownload E-Invoice / GSTR-1 Excel from GST Portal.")
    def agg(lst):
        if not lst: return pd.DataFrame(columns=["inv","date","name","gstin","Taxable","IGST","CGST","SGST","value","ntype","cancelled","sheet","inv_norm"])
        d=pd.DataFrame(lst); d["inv_norm"]=d.inv.map(inorm)
        a=d.groupby(["inv_norm","gstin"],sort=False).agg({"inv":"first","date":"first","name":"first","Taxable":"sum","IGST":"sum","CGST":"sum","SGST":"sum","value":"max","ntype":"first","cancelled":"all","sheet":"first"}).reset_index()
        return a
    di,dn=agg(inv),agg(notes)
    if not has_tax: warns.append("Portal file has no IGST/CGST/SGST amount columns - only taxable value is compared.")
    return {"inv":di,"notes":dn,"hsn":hsn,"layout":layout,"warns":warns,"has_tax":has_tax}

# ── MATCHING ─────────────────────────────────────────────────────────────────
STAT_ORDER=["Matched","Matched - Inv No differs","Mismatch","Only in Books","No GSTIN (B2C / unregistered)","Only in Portal",
            "Cancelled in books","Cancelled in books - LIVE on portal","LIVE in books - cancelled on portal"]
def _ind(x): return f"{abs(x):,.0f}"

def reco_docs(b,p,tol,has_tax,doc="Invoice"):
    b=b.copy().reset_index(drop=True); p=p.copy().reset_index(drop=True)
    if len(b):
        for c in ("Taxable","IGST","CGST","SGST"): b[c]=b[c].abs()
    for c in ("Taxable","IGST","CGST","SGST"): p[c]=p[c].abs()
    pdup=set(p.inv_norm[p.inv_norm.duplicated(keep=False)]) if len(p) else set()
    b["key"]=[(i+"|"+g) if i in pdup else i for i,g in zip(b.inv_norm,b.gstin_n)] if len(b) else []
    p["key"]=[(i+"|"+g) if i in pdup else i for i,g in zip(p.inv_norm,p.gstin)] if len(p) else []
    p["t2"]=False
    # tier-2: same GSTIN + taxable (+date) but different invoice number
    for use_date in (True,False):
        ub=b[(~b.key.isin(p.key))&(~b._cancel)&(b.Taxable>0)&(b.gstin_n!="")] if len(b) else b
        up=p[(~p.key.isin(b.key))&(p.gstin!="")] if len(p) else p
        bmap={}; pmap={}
        for i,r in ub.iterrows(): bmap.setdefault((r.gstin_n,round(r.Taxable),r._date if use_date else None),[]).append(i)
        for i,r in up.iterrows(): pmap.setdefault((r.gstin,round(r.Taxable),r.date if use_date else None),[]).append(i)
        for k,bl in bmap.items():
            pl=pmap.get(k,[])
            if len(bl)==1 and len(pl)==1:
                p.loc[pl[0],"key"]=b.loc[bl[0],"key"]; p.loc[pl[0],"t2"]=True
    bb=b.rename(columns={"_inv":"B_Inv","_date":"B_Date","_party":"B_Party","_vtype":"B_VType","_gstin":"B_GSTIN","Taxable":"B_Tax","IGST":"B_IGST","CGST":"B_CGST","SGST":"B_SGST","_cancel":"B_Cancel"})
    bb=bb[["key","B_Inv","B_Date","B_Party","B_VType","B_GSTIN","B_Tax","B_IGST","B_CGST","B_SGST","B_Cancel"]]
    pp=p.rename(columns={"inv":"P_Inv","date":"P_Date","name":"P_Party","gstin":"P_GSTIN","Taxable":"P_Tax","IGST":"P_IGST","CGST":"P_CGST","SGST":"P_SGST","cancelled":"P_Cancel","ntype":"P_NType"})
    pp=pp[["key","P_Inv","P_Date","P_Party","P_GSTIN","P_Tax","P_IGST","P_CGST","P_SGST","P_Cancel","P_NType","t2"]]
    m=bb.merge(pp,on="key",how="outer",indicator="mrg")
    for c in ("B_Tax","B_IGST","B_CGST","B_SGST","P_Tax","P_IGST","P_CGST","P_SGST"): m[c]=m[c].fillna(0.0)
    for c in ("B_Inv","P_Inv","B_Party","P_Party","B_VType","B_GSTIN","P_GSTIN","P_NType"): m[c]=m[c].fillna("")
    m["B_Cancel"]=m["B_Cancel"].fillna(False).astype(bool); m["P_Cancel"]=m["P_Cancel"].fillna(False).astype(bool); m["t2"]=m["t2"].fillna(False).astype(bool)
    sts=[];pri=[];rem=[]
    for r in m.itertuples():
        bg=r.B_IGST+r.B_CGST+r.B_SGST; pg=r.P_IGST+r.P_CGST+r.P_SGST
        if r.mrg=="left_only":
            if r.B_Cancel or (r.B_Tax==0 and bg==0):
                sts.append("Cancelled in books"); pri.append("Low"); rem.append("Cancelled / nil-value voucher in books; not on portal - OK")
            elif r.B_GSTIN=="":
                sts.append("No GSTIN (B2C / unregistered)"); pri.append("Low"); rem.append("No GSTIN in books - B2C/unregistered, not expected in B2B e-invoice file (check B2C reporting)")
            elif bg>0:
                sts.append("Only in Books"); pri.append("High"); rem.append(f"Taxable supply (GST Rs.{_ind(bg)}) in books but NOT on portal - e-invoice / GSTR-1 missing?")
            else:
                sts.append("Only in Books"); pri.append("Medium"); rem.append("Nil GST, not on portal - confirm exempt / non-GST / bond transfer / zero-rated and check where it is reported in GSTR-1")
        elif r.mrg=="right_only":
            if r.P_Cancel: sts.append("Cancelled in books"); pri.append("Low"); rem.append("Cancelled on portal; not in books - OK")
            else: sts.append("Only in Portal"); pri.append("High"); rem.append("On portal but not in books - check other registers / period cut-off / missed booking")
        else:
            books_dead=r.B_Cancel or (r.B_Tax==0 and bg==0)
            if books_dead and r.P_Cancel: sts.append("Cancelled in books"); pri.append("Low"); rem.append("Cancelled on both sides")
            elif books_dead and (r.P_Tax>0 or pg>0):
                sts.append("Cancelled in books - LIVE on portal"); pri.append("High")
                rem.append(f"ACTION: cancelled in books but live on portal - taxable Rs.{_ind(r.P_Tax)}, GST Rs.{_ind(pg)} auto-populated in GSTR-1. Check IRN cancellation / credit note before filing")
            elif r.P_Cancel and (r.B_Tax>0 or bg>0):
                sts.append("LIVE in books - cancelled on portal"); pri.append("High"); rem.append("Live in books but e-invoice cancelled on portal - check")
            else:
                iss=[]; big=False
                if abs(r.B_Tax-r.P_Tax)>tol: iss.append(f"Taxable differs by Rs.{_ind(r.B_Tax-r.P_Tax)}"); big=True
                if has_tax:
                    for h in ("IGST","CGST","SGST"):
                        dv=getattr(r,f"B_{h}")-getattr(r,f"P_{h}")
                        if abs(dv)>tol: iss.append(f"{h} differs by Rs.{_ind(dv)}"); big=True
                if r.B_GSTIN and r.P_GSTIN and r.B_GSTIN!=r.P_GSTIN: iss.append(f"GSTIN differs (Books {r.B_GSTIN} / Portal {r.P_GSTIN})")
                bd=r.B_Date if not pd.isna(r.B_Date) else None; pdt=r.P_Date if not pd.isna(r.P_Date) else None
                if bd and pdt and bd!=pdt: iss.append(f"Date differs (Books {bd:%d-%b-%y} / Portal {pdt:%d-%b-%y})")
                if iss: sts.append("Mismatch"); pri.append("High" if big else "Medium"); rem.append("; ".join(iss)+(f" | Inv No: Books {r.B_Inv} / Portal {r.P_Inv}" if r.t2 else ""))
                elif r.t2: sts.append("Matched - Inv No differs"); pri.append("Low"); rem.append(f"Matched on GSTIN + value; invoice no differs: Books '{r.B_Inv}' vs Portal '{r.P_Inv}'")
                else: sts.append("Matched"); pri.append(""); rem.append("")
    m["Status"]=sts; m["Priority"]=pri; m["Remark"]=rem
    m["Date"]=[bd if not pd.isna(bd) else pdt for bd,pdt in zip(m.B_Date,m.P_Date)]
    m["Party"]=[bp or pp_ for bp,pp_ in zip(m.B_Party,m.P_Party)]
    m["GSTIN"]=[bg or pg for bg,pg in zip(m.B_GSTIN,m.P_GSTIN)]
    order={s:i for i,s in enumerate(STAT_ORDER)}; m["_o"]=m.Status.map(order)
    m["_d"]=[d if d is not None and not pd.isna(d) else datetime.date(2100,1,1) for d in m.Date]
    m["_n"]=[bi or pi for bi,pi in zip(m.B_Inv,m.P_Inv)]
    return m.sort_values(["_o","_d","_n"]).drop(columns=["_o","_d","_n","mrg","key"]).reset_index(drop=True)

def _run_sales_reco_raw(portal_bytes,books_bytes,ledger_class=None,vtype_class=None,tol=1.0,scan=None):
    scan=scan or scan_books(books_bytes)
    lc=dict(scan["ledger_guess"]); lc.update(ledger_class or {})
    vc=dict(scan["vtype_guess"]); vc.update(vtype_class or {})
    bk=build_books(scan["bw"],lc,vc); pt=parse_portal(portal_bytes)
    books_tax=float(pd.concat([bk["inv"],bk["notes"]])[["IGST","CGST","SGST"]].abs().sum().sum())>0
    htx=pt["has_tax"] and books_tax
    reco=reco_docs(bk["inv"],pt["inv"],tol,htx,"Invoice")
    notes=reco_docs(bk["notes"],pt["notes"],tol,htx,"Note")
    # likely original invoice for portal-only notes (same GSTIN + same taxable in books invoices)
    lik=[]
    for r in notes.itertuples():
        txt=""
        if r.P_Inv and not r.B_Inv:
            c=bk["inv"][(bk["inv"].gstin_n==r.P_GSTIN)&((bk["inv"].Taxable-r.P_Tax).abs()<1)]
            txt=", ".join(c._inv.tolist()[:3])
        lik.append(txt)
    notes["Likely_Orig"]=lik
    return {"reco":reco,"notes":notes,"books":bk,"portal":pt,"ledger_class":lc,"vtype_class":vc,"scan":scan,"tol":tol,
            "warns":scan["warns"]+pt["warns"]+([] if (not pt["has_tax"] or books_tax) else ["Books file has no IGST/CGST/SGST ledger columns (or they are classed as Other/Ignore) - only TAXABLE value was compared. Add the tax ledger columns for a full tax check."])+([f"{len(bk['dups'])} duplicate voucher number(s) inside the same books sheet were merged - verify: "+", ".join(map(lambda x:str(x[1]),list(bk['dups'].index[:8])))] if len(bk['dups']) else [])}


# ── notes post-processing + Excel builder ────────────────────────────────────
def _fix_notes(res):
    n=res["notes"]
    if len(res["books"]["notes"])==0 and len(n):
        m=n.Status=="Only in Portal"
        n.loc[m,"Remark"]="Books file has no credit/debit-note vouchers - provide Credit Note register to reconcile"
        n.loc[m,"Priority"]="Medium"
    return res
def run_sales_reco(*a,**k): return _fix_notes(_run_sales_reco_raw(*a,**k))

def build_sales_excel(res,client,period):
    wb=openpyxl.Workbook(); reco=res["reco"]; notes=res["notes"]; bk=res["books"]; pt=res["portal"]
    BG={"Matched":MATCH,"Matched - Inv No differs":"E2EFDA","Mismatch":MIS,"Only in Books":ONLY1,"No GSTIN (B2C / unregistered)":EXEMPT,
        "Only in Portal":ONLY2,"Cancelled in books":EXEMPT,"Cancelled in books - LIVE on portal":MIS,"LIVE in books - cancelled on portal":MIS}
    def title(ws,txt,nc,r=1):
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=nc); c=ws.cell(r,1,txt)
        c.font=Font(name="Calibri",bold=True,size=12,color="FFFFFF"); c.fill=F(TITLE); c.alignment=CTR(); ws.row_dimensions[r].height=26
    def sub(ws,txt,nc,r=2):
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=nc); c=ws.cell(r,1,f"Created by: Purva Doshi  |  linkedin.com/in/purvadoshi26  |  {txt}")
        c.font=Font(name="Calibri",size=8,italic=True,color="7F9FC0"); c.alignment=CTR()
    def head(ws,r,hs,color=NAVY):
        for ci,h in enumerate(hs,1):
            c=ws.cell(r,ci,h); c.font=HF(9); c.fill=F(color); c.alignment=CTR(); c.border=BD()
        ws.row_dimensions[r].height=28
    def widths(ws,ws_):
        for i,w in enumerate(ws_,1): ws.column_dimensions[get_column_letter(i)].width=w
    def dval(x):
        return None if (x is None or (not isinstance(x,(datetime.date,)) and pd.isna(x))) else x

    # ---------- Invoice Reco ----------
    wi=wb.active; wi.title="Invoice Reco"; wi.sheet_view.showGridLines=False
    H=["#","Inv No (Books)","Inv No (Portal)","Date","Party","Voucher Type","GSTIN (Books)","GSTIN (Portal)","Books Taxable","Portal Taxable","Taxable Diff",
       "Books IGST","Portal IGST","Books CGST","Portal CGST","Books SGST","Portal SGST","GST Diff","Status","Priority","Remark"]
    title(wi,f"{client}  -  E-Invoice / GSTR-1 vs Books  |  Invoice Reco  |  {period}",len(H)); sub(wi,"Differences are live formulas (Books - Portal)",len(H)); head(wi,4,H)
    r0=5
    for i,r in reco.iterrows():
        x=r0+i
        vals=[i+1,r.B_Inv,r.P_Inv,dval(r.Date),r.Party,r.B_VType,r.B_GSTIN,r.P_GSTIN,r.B_Tax,r.P_Tax,f"=I{x}-J{x}",r.B_IGST,r.P_IGST,r.B_CGST,r.P_CGST,r.B_SGST,r.P_SGST,
              f"=(L{x}+N{x}+P{x})-(M{x}+O{x}+Q{x})",r.Status,r.Priority,r.Remark]
        for ci,v in enumerate(vals,1):
            c=wi.cell(x,ci,v); c.font=DF(9); c.border=BD(); c.fill=F(BG.get(r.Status,"FFFFFF"))
            if 9<=ci<=18: c.number_format=INR; c.alignment=RGT()
            elif ci==4: c.number_format="DD-MMM-YY"; c.alignment=CTR()
            elif ci in (1,19,20): c.alignment=CTR()
        wi.cell(x,19).font=DF(9,bold=True)
    last=r0+len(reco)-1; tr=last+1
    wi.cell(tr,5,"TOTAL").font=DF(9,bold=True)
    for ci in range(9,19):
        c=wi.cell(tr,ci,f"=SUM({get_column_letter(ci)}{r0}:{get_column_letter(ci)}{last})"); c.font=DF(9,bold=True); c.number_format=INR; c.fill=F("D9E1F2"); c.border=BD()
    widths(wi,[5,20,20,11,38,24,18,18,15,15,14,13,13,13,13,13,13,13,30,9,70]); wi.freeze_panes="C5"; wi.auto_filter.ref=f"A4:{get_column_letter(len(H))}{last}"
    RG=lambda col: f"'Invoice Reco'!${col}${r0}:${col}${last}"

    # ---------- Summary ----------
    sm=wb.create_sheet("Summary",0); sm.sheet_view.showGridLines=False
    title(sm,f"{client}  -  E-Invoice / GSTR-1 vs Books Sales Reco  |  {period}",7); sub(sm,f"Sales Reco | {period}",7)
    head(sm,4,["A. Invoices - status","Count","Books Taxable","Portal Taxable","Taxable Diff","GST Diff (Books-Portal)","Meaning"])
    mean={"Matched":"All fields agree","Matched - Inv No differs":"Same GSTIN + value; number format differs","Mismatch":"Same invoice, value/GSTIN/date differs",
          "Only in Books":"Not on portal - see Exceptions","No GSTIN (B2C / unregistered)":"B2C - not expected in B2B file","Only in Portal":"Not found in books",
          "Cancelled in books":"Dead voucher, nothing on portal","Cancelled in books - LIVE on portal":"ACTION - tax on portal without books","LIVE in books - cancelled on portal":"ACTION - check"}
    for i,s in enumerate(STAT_ORDER):
        x=5+i; sm.cell(x,1,s); sm.cell(x,2,f'=COUNTIF({RG("S")},A{x})'); sm.cell(x,3,f'=SUMIF({RG("S")},A{x},{RG("I")})'); sm.cell(x,4,f'=SUMIF({RG("S")},A{x},{RG("J")})')
        sm.cell(x,5,f"=C{x}-D{x}"); sm.cell(x,6,f'=SUMIF({RG("S")},A{x},{RG("R")})'); sm.cell(x,7,mean[s])
    xt=5+len(STAT_ORDER); sm.cell(xt,1,"TOTAL")
    for ci in range(2,7): sm.cell(xt,ci,f"=SUM({get_column_letter(ci)}5:{get_column_letter(ci)}{xt-1})")
    for x in range(5,xt+1):
        for ci in range(1,8):
            c=sm.cell(x,ci); c.border=BD(); c.font=DF(9,bold=(x==xt or ci==1))
            if 3<=ci<=6: c.number_format=INR; c.alignment=RGT()
            if ci==2: c.alignment=CTR(); c.number_format="#,##0"
            if x==xt: c.fill=F("D9E1F2")
            elif ci==1: c.fill=F(BG.get(sm.cell(x,1).value,"FFFFFF"))
    q=xt+2
    head(sm,q,["B. Tie-out","","Amount (Rs.)","","","",""])
    tie=[("Total taxable per Books",f"='Invoice Reco'!I{tr}"),("Total taxable per Portal",f"='Invoice Reco'!J{tr}"),("Difference (Books - Portal)",f"=C{q+1}-C{q+2}"),
         ("Explained by status rows above (total taxable diff)",f"=E{xt}"),("Unexplained",f"=C{q+3}-C{q+4}")]
    for i,(a,b) in enumerate(tie):
        sm.cell(q+1+i,1,a); sm.cell(q+1+i,3,b); sm.cell(q+1+i,3).number_format=INR
        for ci in (1,2,3): sm.cell(q+1+i,ci).border=BD(); sm.cell(q+1+i,ci).font=DF(9,bold=(i in (2,4)))
    q2=q+len(tie)+2
    head(sm,q2,["C. Credit / Debit notes","Count","Books Taxable","Portal Taxable","Taxable Diff","","Note"])
    nn=len(notes)
    # notes summary written after notes sheet is built (needs ranges) -> placeholders filled below
    # ---------- Notes sheet ----------
    wn=wb.create_sheet("Credit-Debit Notes"); wn.sheet_view.showGridLines=False
    NH=["#","Note No (Books)","Note No (Portal)","Date","Party","Type (C/D)","GSTIN","Books Taxable","Portal Taxable","Taxable Diff","Books GST","Portal GST","GST Diff","Status","Priority","Likely original invoice (books)","Remark"]
    title(wn,f"{client}  -  Credit / Debit Notes Reco  |  {period}",len(NH)); sub(wn,"Absolute values compared. 'Likely original' = books invoice with same GSTIN + same taxable value (indicative)",len(NH)); head(wn,4,NH)
    n0=5
    for i,r in notes.iterrows():
        x=n0+i
        vals=[i+1,r.B_Inv,r.P_Inv,dval(r.Date),r.Party,r.P_NType,r.GSTIN,r.B_Tax,r.P_Tax,f"=H{x}-I{x}",r.B_IGST+r.B_CGST+r.B_SGST,r.P_IGST+r.P_CGST+r.P_SGST,f"=K{x}-L{x}",r.Status,r.Priority,r.Likely_Orig,r.Remark]
        for ci,v in enumerate(vals,1):
            c=wn.cell(x,ci,v); c.font=DF(9); c.border=BD(); c.fill=F(BG.get(r.Status,"FFFFFF"))
            if 8<=ci<=13: c.number_format=INR; c.alignment=RGT()
            elif ci==4: c.number_format="DD-MMM-YY"; c.alignment=CTR()
            elif ci in (1,6,14,15): c.alignment=CTR()
    nlast=n0+nn-1
    if nn:
        wn.cell(nlast+1,5,"TOTAL").font=DF(9,bold=True)
        for ci in range(8,14):
            c=wn.cell(nlast+1,ci,f"=SUM({get_column_letter(ci)}{n0}:{get_column_letter(ci)}{nlast})"); c.font=DF(9,bold=True); c.number_format=INR; c.fill=F("D9E1F2"); c.border=BD()
    widths(wn,[5,20,20,11,38,9,18,15,15,14,13,13,13,22,9,34,70]); wn.freeze_panes="C5"
    NR=lambda col: f"'Credit-Debit Notes'!${col}${n0}:${col}${max(nlast,n0)}"
    nst=[s for s in STAT_ORDER if (notes.Status==s).any()] if nn else []
    for i,s in enumerate(nst):
        x=q2+1+i; sm.cell(x,1,s); sm.cell(x,2,f'=COUNTIF({NR("N")},A{x})'); sm.cell(x,3,f'=SUMIF({NR("N")},A{x},{NR("H")})'); sm.cell(x,4,f'=SUMIF({NR("N")},A{x},{NR("I")})'); sm.cell(x,5,f"=C{x}-D{x}")
        if s=="Only in Portal" and len(bk["notes"])==0: sm.cell(x,7,"Books file has no credit-note vouchers - need Credit Note register")
        for ci in range(1,8):
            c=sm.cell(x,ci); c.border=BD(); c.font=DF(9,bold=(ci==1))
            if 3<=ci<=5: c.number_format=INR; c.alignment=RGT()
            if ci==2: c.alignment=CTR()
    if not nst:
        sm.cell(q2+1,1,"No credit/debit notes found in portal or books file").font=DF(9,c="777777")
    q3=q2+max(len(nst),1)+2
    # observations
    sm.cell(q3,1,"D. Key observations").font=Font(name="Calibri",bold=True,size=11,color=NAVY)
    obs=[]; vc=reco.Status.value_counts().to_dict(); hi=reco[reco.Priority=="High"]
    obs.append(f"{vc.get('Matched',0)+vc.get('Matched - Inv No differs',0)} of {len(reco)} invoice rows reconcile fully (GSTIN, date, taxable value, GST).")
    for s,t in [("Cancelled in books - LIVE on portal","invoice(s) cancelled in books but LIVE on portal - GST auto-populated in GSTR-1 (see Exceptions)"),
                ("LIVE in books - cancelled on portal","invoice(s) live in books but cancelled on portal"),
                ("Mismatch","invoice(s) with value / GSTIN / date mismatch"),
                ("Only in Portal","invoice(s) on portal but not in books"),
                ("Only in Books","invoice(s) in books but not on portal (taxable ones are High priority; nil-GST ones need classification check)"),
                ("No GSTIN (B2C / unregistered)","B2C voucher(s) without GSTIN - not expected in B2B file")]:
        if vc.get(s,0): obs.append(f"{vc[s]} {t}.")
    ob=reco[(reco.Status=="Only in Books")]
    if len(ob): obs.append(f"Only-in-Books split: {int((ob.Priority=='High').sum())} taxable (High), {int((ob.Priority=='Medium').sum())} nil-GST (Rs.{ob[ob.Priority=='Medium'].B_Tax.sum():,.0f} taxable - confirm exempt / non-GST / bond transfer).")
    if nn and len(bk["notes"])==0: obs.append(f"{nn} credit/debit notes on portal but none in books file - reconcile with Credit Note register.")
    if pt["hsn"] is not None: obs.append("HSN summary tie-out is in 'Mapping & Checks'.")
    for w in res["warns"]: obs.append("WARNING: "+w)
    for i,t in enumerate(obs):
        x=q3+1+i; c=sm.cell(x,1,t); c.font=DF(9); c.alignment=Alignment(wrap_text=True,vertical="top"); sm.merge_cells(start_row=x,start_column=1,end_row=x,end_column=7); sm.row_dimensions[x].height=28
    widths(sm,[44,9,18,18,18,20,52])

    # ---------- Exceptions ----------
    we=wb.create_sheet("Exceptions"); we.sheet_view.showGridLines=False
    EH=["Priority","Doc","Status","Inv / Note No","Date","Party","GSTIN","Books Taxable","Portal Taxable","Books GST","Portal GST","Remark"]
    title(we,f"{client}  -  Exceptions needing action  |  {period}",len(EH)); sub(we,"All rows that are not clean matches / dead vouchers",len(EH)); head(we,4,EH)
    rows=[]
    for df_,doc in ((reco,"Invoice"),(notes,"Note")):
        for r in df_.itertuples():
            if r.Status in ("Matched","Matched - Inv No differs","Cancelled in books"): continue
            rows.append((r.Priority,doc,r.Status,r.B_Inv or r.P_Inv,dval(r.Date),r.Party,r.GSTIN,r.B_Tax,r.P_Tax,r.B_IGST+r.B_CGST+r.B_SGST,r.P_IGST+r.P_CGST+r.P_SGST,r.Remark))
    po={"High":0,"Medium":1,"Low":2,"":3}; rows.sort(key=lambda t:(po.get(t[0],3),t[1],t[2]))
    for i,rw in enumerate(rows):
        for ci,v in enumerate(rw,1):
            c=we.cell(5+i,ci,v); c.font=DF(9,bold=(ci==1),c=("C00000" if (ci==1 and rw[0]=="High") else "222222")); c.border=BD()
            if 8<=ci<=11: c.number_format=INR; c.alignment=RGT()
            if ci==5: c.number_format="DD-MMM-YY"; c.alignment=CTR()
    widths(we,[9,9,30,22,11,38,18,15,15,13,13,90]); we.freeze_panes="A5"; we.auto_filter.ref=f"A4:L{4+max(len(rows),1)}"

    # ---------- Mapping & Checks ----------
    wm=wb.create_sheet("Mapping & Checks"); wm.sheet_view.showGridLines=False
    title(wm,f"{client}  -  What the app detected + integrity checks  |  {period}",6); sub(wm,"Review this once per client - if a ledger is mis-classified, fix it in the app and re-run",6)
    r=4
    def block(t): 
        nonlocal r; c=wm.cell(r,1,t); c.font=Font(name="Calibri",bold=True,size=11,color=NAVY); r+=1
    def row_(vals,hdr=False,fmt=None):
        nonlocal r
        for ci,v in enumerate(vals,1):
            c=wm.cell(r,ci,v); c.border=BD()
            if hdr: c.font=HF(9); c.fill=F(NAVY); c.alignment=CTR()
            else:
                c.font=DF(9)
                if fmt and ci in fmt: c.number_format=fmt[ci]; c.alignment=RGT()
        r+=1
    block("1. Books file - columns found (by header name)")
    row_(["Sheet","Header row","Voucher rows","Field -> header used"],True)
    for L_ in res["scan"]["layout"]: row_([L_["sheet"],L_["header_row"],L_["rows"],"; ".join(f"{k}: {v}" for k,v in L_["mapped"].items())])
    r+=1; block("2. Ledger columns - class used in reco (Taxable / IGST / CGST / SGST / Cess / Other / Ignore)")
    row_(["Ledger (column header)","Class used","Auto-guess","Total in file (Rs.)"],True)
    bw=res["scan"]["bw"]
    for k,v in res["ledger_class"].items(): row_([k,v,res["scan"]["ledger_guess"].get(k,""),float(bw[k].sum()) if k in bw else 0.0],fmt={4:INR})
    r+=1; block("3. Voucher types")
    row_(["Voucher type","Treated as","Vouchers"],True)
    for k,v in res["vtype_class"].items(): row_([k or "(blank)",v,int((bw._vtype==k).sum())])
    r+=1; block("4. Portal file - sheets and columns found")
    row_(["Sheet","Role","Header row","Rows","Field -> header used"],True)
    for L_ in pt["layout"]: row_([L_["sheet"],L_["role"],L_["header_row"],L_["rows"],"; ".join(f"{k}: {v}" for k,v in L_["mapped"].items())])
    r+=1; block("5. Books internal checks")
    row_(["Check","Result","Detail"],True)
    row_(["Gross total = value + tax (per voucher)",f"{bk['gross_checked']-len(bk['gross_fail'])} of {bk['gross_checked']} pass",("Check ledger classes - failures listed below" if len(bk['gross_fail']) else "OK")])
    row_(["Effective GST rate is a standard slab",f"{len(bk['rate_fail'])} unusual","Unusual rate usually means a ledger is mis-classified (value vs tax)" if len(bk['rate_fail']) else "OK"])
    row_(["IGST and CGST/SGST on same voucher",f"{len(bk['mixed'])} found","OK" if not len(bk['mixed']) else "Check place of supply"])
    row_(["Duplicate voucher numbers in same sheet",f"{len(bk['dups'])} found","OK" if not len(bk['dups']) else "Merged - verify"])
    for lab,df_ in (("Gross mismatch",bk["gross_fail"]),("Unusual rate",bk["rate_fail"])):
        for _,q_ in df_.head(30).iterrows(): row_([lab,q_["_inv"],f"Taxable {q_.Taxable:,.0f} IGST {q_.IGST:,.0f} CGST {q_.CGST:,.0f} SGST {q_.SGST:,.0f}"])
    if pt["hsn"] is not None:
        r+=1; block("6. HSN summary tie-out (portal invoices - credit notes + debit notes vs HSN sheet)")
        row_(["","Taxable","IGST","CGST","SGST"],True)
        sg=lambda col,sign: float(notes[(notes.P_NType==sign)][col].sum())
        inv_t={k:float(pt["inv"][k].sum()) for k in ("Taxable","IGST","CGST","SGST")}
        cn={k:float(pt["notes"][pt["notes"].ntype!="D"][k].sum()) for k in ("Taxable","IGST","CGST","SGST")} if len(pt["notes"]) else {k:0.0 for k in inv_t}
        dn={k:float(pt["notes"][pt["notes"].ntype=="D"][k].sum()) for k in ("Taxable","IGST","CGST","SGST")} if len(pt["notes"]) else {k:0.0 for k in inv_t}
        a=r; row_(["Portal invoices"]+[inv_t[k] for k in ("Taxable","IGST","CGST","SGST")],fmt={2:INR,3:INR,4:INR,5:INR})
        row_(["Less: credit notes"]+[-cn[k] for k in ("Taxable","IGST","CGST","SGST")],fmt={2:INR,3:INR,4:INR,5:INR})
        row_(["Add: debit notes"]+[dn[k] for k in ("Taxable","IGST","CGST","SGST")],fmt={2:INR,3:INR,4:INR,5:INR})
        row_(["Expected HSN total"]+[f"=SUM({get_column_letter(c)}{a}:{get_column_letter(c)}{a+2})" for c in (2,3,4,5)],fmt={2:INR,3:INR,4:INR,5:INR})
        row_(["HSN sheet total"]+[pt["hsn"][k] for k in ("Taxable","IGST","CGST","SGST")],fmt={2:INR,3:INR,4:INR,5:INR})
        row_(["Difference (HSN - expected)"]+[f"={get_column_letter(c)}{a+4}-{get_column_letter(c)}{a+3}" for c in (2,3,4,5)],fmt={2:INR,3:INR,4:INR,5:INR})
    if res["warns"]:
        r+=1; block("Warnings")
        for w in res["warns"]: row_([w])
    widths(wm,[46,22,20,20,90,10])
    wb.move_sheet("Summary",-wb.index(wb["Summary"]))
    # conditional-free static colours already applied
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

# ─────────────────────────────────────────────────────────────────────────────
#  TAB 3 — 3B CHECK (GSTR-3B vs GSTR-1 output + 2B ITC)
# ─────────────────────────────────────────────────────────────────────────────
def build_3b_excel(data_3b,data_1,data_2b_itc,client,period):
    """
    data_3b: dict of 3B values entered by user
    data_1:  dict of GSTR-1 values
    data_2b_itc: dict of 2B ITC available
    """
    wb=openpyxl.Workbook()
    def mhdr(ws,txt,nc,r=1):
        ws.merge_cells(f"A{r}:{get_column_letter(nc)}{r}"); c=ws.cell(r,1); c.value=txt
        c.font=Font(name="Calibri",bold=True,size=12,color="FFFFFF"); c.fill=F(TITLE); c.alignment=CTR(); ws.row_dimensions[r].height=28
    def sub(ws,txt,nc,r=2):
        ws.merge_cells(f"A{r}:{get_column_letter(nc)}{r}"); c=ws.cell(r,1)
        c.value=f"Created by: Purva Doshi  |  linkedin.com/in/purvadoshi26  |  {txt}"
        c.font=Font(name="Calibri",size=8,italic=True,color="7F9FC0"); c.alignment=CTR(); ws.row_dimensions[r].height=14

    ws1=wb.active; ws1.title="3B Check"; ws1.sheet_view.showGridLines=False
    mhdr(ws1,f"{client}  —  GSTR-3B Audit Check  |  {period}",7); sub(ws1,f"3B Check | {period}",7)

    # Output liability: 3B vs GSTR-1
    ws1.merge_cells("A4:G4"); c=ws1.cell(4,1); c.value="A. OUTPUT LIABILITY — GSTR-3B vs GSTR-1"
    c.font=HF(10,NAVY); c.fill=F("E8EFF8"); c.alignment=CTR(); c.border=BD()
    for ci,v in enumerate(["Section","Description","3B (Filed)","GSTR-1 (Declared)","Diff","Status","Remark"],1):
        c=ws1.cell(5,ci); c.value=v; c.font=HF(8); c.fill=F(NAVY); c.alignment=CTR(); c.border=BD()

    out_rows=[
        ("3.1(a)","B2B Taxable Value",data_3b.get("out_taxable",0),data_1.get("b2b_taxable",0)),
        ("3.1(a)","IGST on Outward",data_3b.get("out_igst",0),data_1.get("b2b_igst",0)),
        ("3.1(a)","CGST on Outward",data_3b.get("out_cgst",0),data_1.get("b2b_cgst",0)),
        ("3.1(a)","SGST on Outward",data_3b.get("out_sgst",0),data_1.get("b2b_sgst",0)),
    ]
    for ri,(sec,desc,v3b,v1) in enumerate(out_rows,6):
        diff=v3b-v1; ok=abs(diff)<2
        st="✅ Match" if ok else ("⚠️ 3B > GSTR-1" if diff>0 else "⚠️ 3B < GSTR-1")
        rm="" if ok else ("Excess declared in 3B — check amendments" if diff>0 else "Short paid in 3B vs GSTR-1 — interest risk u/s 50")
        for ci,val in enumerate([sec,desc,v3b,v1,diff,st,rm],1):
            c=ws1.cell(ri,ci); c.value=val; c.border=BD(); c.font=DF(9,bold=(ci==1))
            if ci in(3,4,5) and isinstance(val,(int,float)): c.number_format=INR; c.alignment=RGT()
            elif ci==6:
                c.font=Font(name="Calibri",bold=True,size=9,color="1E5631" if ok else "C00000"); c.alignment=CTR()
            else: c.alignment=LFT()
            if not ok: c.fill=F(MIS if ci in(5,6) else "FFFFFF")
        ws1.row_dimensions[ri].height=17

    # ITC: 3B vs 2B
    r_start=12
    ws1.merge_cells(f"A{r_start}:G{r_start}"); c=ws1.cell(r_start,1)
    c.value="B. ITC CLAIMED — GSTR-3B vs GSTR-2B Available"
    c.font=HF(10,NAVY); c.fill=F("E8EFF8"); c.alignment=CTR(); c.border=BD()
    for ci,v in enumerate(["Section","Description","3B Claimed","2B Available","Diff","Status","Remark"],1):
        c=ws1.cell(r_start+1,ci); c.value=v; c.font=HF(8); c.fill=F(NAVY); c.alignment=CTR(); c.border=BD()
    itc_rows=[
        ("4(A)(1)","Import of Goods IGST",data_3b.get("itc_import",0),data_2b_itc.get("impg_igst",0)),
        ("4(A)(3)","RCM ITC",data_3b.get("itc_rcm",0),data_2b_itc.get("rcm_total",0)),
        ("4(A)(5)","All Other ITC — IGST",data_3b.get("itc_igst",0),data_2b_itc.get("b2b_igst",0)),
        ("4(A)(5)","All Other ITC — CGST",data_3b.get("itc_cgst",0),data_2b_itc.get("b2b_cgst",0)),
        ("4(A)(5)","All Other ITC — SGST",data_3b.get("itc_sgst",0),data_2b_itc.get("b2b_sgst",0)),
        ("4(C)","Net ITC Available",data_3b.get("net_itc",0),data_2b_itc.get("net_available",0)),
    ]
    for ri,(sec,desc,v3b,v2b) in enumerate(itc_rows,r_start+2):
        diff=v3b-v2b; ok=diff<=1  # claimed <= available is ok; claimed > available = risk
        st="✅ OK" if ok else "🔴 EXCESS CLAIM RISK"
        rm="" if ok else f"Claimed ₹{v3b:,.0f} but only ₹{v2b:,.0f} available in 2B — excess ITC reversal required u/s 42"
        for ci,val in enumerate([sec,desc,v3b,v2b,diff,st,rm],1):
            c=ws1.cell(ri,ci); c.value=val; c.border=BD(); c.font=DF(9,bold=(ci==1))
            if ci in(3,4,5) and isinstance(val,(int,float)): c.number_format=INR; c.alignment=RGT()
            elif ci==6: c.font=Font(name="Calibri",bold=True,size=9,color="1E5631" if ok else "C00000"); c.alignment=CTR()
            else: c.alignment=LFT()
            if not ok: c.fill=F(MIS if ci in(5,6) else "FFFFFF")
        ws1.row_dimensions[ri].height=17

    for col,w in zip("ABCDEFG",[12,32,16,16,16,22,55]): ws1.column_dimensions[col].width=w
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

# ─────────────────────────────────────────────────────────────────────────────
#  TAB 1 (v2) — ITC RECO, INVOICE-WISE  (GSTR-2B vs Purchase Register)
#  • 2B read by HEADER NAME from every sheet: B2B, B2BA, B2B-CDNR, B2B-CDNRA, ISD, IMPG, IMPGSEZ
#  • Books read by header name (any Tally / SAP / Excel layout) + legacy Tally sub-row/pivot fallbacks
#  • Matching tiers: GSTIN+Inv No → GSTIN+numeric core → GSTIN+tax amount → Inv No / amount with GSTIN typo
# ─────────────────────────────────────────────────────────────────────────────
import numpy as np
HEADS=("igst","cgst","sgst")
ITC_STAT=["Matched","Matched - Inv No differs","Matched - Rounding","Mismatch","Only in 2B","2B: ITC not available",
          "RCM - verify","Only in Books","Nil-tax / cancelled in books"]
ITC_BG={"Matched":MATCH,"Matched - Inv No differs":"E2EFDA","Matched - Rounding":"E2EFDA","Mismatch":MIS,"Only in 2B":ONLY2,
        "2B: ITC not available":"FCE4D6","RCM - verify":RCM_C,"Only in Books":ONLY1,"Nil-tax / cancelled in books":EXEMPT}
ITC_MEAN={"Matched":"Invoice in both, GSTIN + tax agree","Matched - Inv No differs":"Same GSTIN + tax; invoice number written differently",
 "Matched - Rounding":"Agrees within rounding tolerance","Mismatch":"Same invoice, tax / taxable / head differs",
 "Only in 2B":"Supplier filed, not booked by client","2B: ITC not available":"On 2B but 2B itself says ITC not available (Sec 16(4)/17(5)/POS etc.)",
 "RCM - verify":"Reverse-charge: tax payable by client first","Only in Books":"Booked, supplier has not reported it - ITC not claimable now",
 "Nil-tax / cancelled in books":"No tax in books and not on 2B - nothing to claim"}

def icore(s):
    m=re.findall(r"\d+",str(s or ""))
    return (m[-1].lstrip("0") or "0") if m else ""

# ── 2B (portal) ──────────────────────────────────────────────────────────────
TB_COLS={
 "gstin":[r"^gstin of supplier$",r"^gstin of isd$",r"^gstin/uin of supplier$",r"^gstin"],
 "name": [r"^trade/legal name$",r"^trade name",r"^legal name",r"^name of supplier"],
 "oinv": [r"^original (invoice|note|document) (number|no)"],
 "inv":  [r"^invoice number$",r"^note number$",r"^invoice no\.?$",r"^isd (invoice|credit note) (number|no)",r"^document number$"],
 "ntype":[r"^note type$",r"^document type$",r"^isd document type$"],
 "date": [r"^invoice date$",r"^note date$",r"^document date$"],
 "value":[r"^invoice value",r"^note value",r"^document value"],
 "rcm":  [r"^supply attract reverse charge",r"reverse charge"],
 "taxable":[r"^taxable value",r"^taxable amount"],
 "igst": [r"^integrated tax",r"^igst"], "cgst":[r"^central tax",r"^cgst"], "sgst":[r"^state/ut tax",r"^state tax",r"^sgst"], "cess":[r"^cess"],
 "avail":[r"^itc availability"], "reason":[r"^reason$"], "period":[r"gstr-?1.*period",r"^tax period"],
}
def _role(sn):
    t=sn.strip().lower()
    if t in ("read me","readme","itc summary","gstr-2b","gstr2b") or t.startswith("itc ") or t in ("eco","ecoa"): return None
    if t.startswith("b2ba"): return "inv_amd"
    if "cdnra" in t: return "note_amd"
    if "cdnr" in t: return "note"
    if t.startswith("isd"): return "isd"
    if t.startswith("impg"): return "imp"
    if t.startswith("b2b"): return "inv"
    return None

def parse_2b_full(fb):
    wb=load_wb(fb); inv=[]; notes=[]; isd=[]; imp=[]; layout=[]; warns=[]
    if not any(_role(s) for s in wb.sheetnames):
        raise ValueError("2B file: no B2B / B2B-CDNR / IMPG sheets found.\nDownload Excel from GST Portal → Returns → GSTR-2B → Download Excel.")
    for ws in wb.worksheets:
        role=_role(ws.title)
        if not role: continue
        rows=[list(r) for r in ws.iter_rows(values_only=True)]
        if role=="imp":
            hr=None
            for i,row in enumerate(rows[:15]):
                if any("port code" in _hn(v) for v in row): hr=i; break
            if hr is None: continue
            hs=[_hn(v) for v in row_ if True] if False else [_hn(v) for v in rows[hr]]
            def ix(rx):
                for i,h in enumerate(hs):
                    if re.search(rx,h): return i
            c_ref,c_port,c_no,c_dt,c_tx,c_ig=ix(r"reference date|ref"),ix(r"port"),ix(r"^number|be (no|number)|bill of entry"),ix(r"^date|be date"),ix(r"taxable"),ix(r"integrated")
            n0=len(imp)
            for r in rows[hr+1:]:
                if not any(v not in (None,"") for v in r): continue
                g=lambda c: r[c] if c is not None and c<len(r) else None
                if g(c_port) in (None,"") and g(c_no) in (None,""): continue
                imp.append({"Ref Date":str(g(c_ref) or ""),"Port":str(g(c_port) or ""),"BE No":str(g(c_no) or ""),"BE Date":str(to_date(g(c_dt)) or g(c_dt) or ""),
                            "Taxable":num(g(c_tx)),"IGST":num(g(c_ig)),"Sheet":ws.title})
            layout.append({"sheet":ws.title,"role":"Imports (IMPG)","header_row":hr+1,"rows":len(imp)-n0,"mapped":{}}); continue
        hr,cm=_find_hdr(rows,TB_COLS,["gstin","igst"] if role!="isd" else ["igst"],3)
        if hr is None:
            hr,cm=_find_hdr(rows,TB_COLS,["gstin"],3)
        if hr is None:
            warns.append(f"2B sheet '{ws.title}' skipped: header row not recognised."); continue
        hdr=rows[hr]; n0=len(inv)+len(notes)+len(isd)
        for r in rows[hr+1:]:
            g=lambda f: (r[cm[f]] if f in cm and cm[f]<len(r) else None)
            gs=cgstin(g("gstin"))
            if not gs or not re.match(r"^\d{2}[A-Z]{5}\d{4}[A-Z]\d[A-Z0-9][A-Z0-9]$",gs): continue
            ivn=str(g("inv") if g("inv") is not None else "").strip()
            rec={"gstin":gs,"party":str(g("name") or "").strip(),"inv":ivn,"oinv":str(g("oinv") or "").strip(),"date":to_date(g("date")),
                 "taxable":num(g("taxable")),"igst":num(g("igst")),"cgst":num(g("cgst")),"sgst":num(g("sgst")),"cess":num(g("cess")),
                 "rigst":0.0,"rcgst":0.0,"rsgst":0.0,
                 "rcm":str(g("rcm") or "").strip().lower().startswith("y"),"avail":str(g("avail") or "").strip(),"reason":str(g("reason") or "").strip(),
                 "period":str(g("period") or ""),"src":ws.title,"amend":role.endswith("_amd"),
                 "ntype":("D" if "debit" in str(g("ntype") or "").lower() else "C") if "note" in role or role=="isd" else "",
                 "cancel":False,"vch":""}
            if role=="isd": isd.append(rec)
            elif role.startswith("note"): notes.append(rec)
            else: inv.append(rec)
        layout.append({"sheet":ws.title,"role":{"inv":"Invoices","inv_amd":"Amended invoices (B2BA)","note":"Credit/Debit notes","note_amd":"Amended notes","isd":"ISD credit"}[role],
                       "header_row":hr+1,"rows":len(inv)+len(notes)+len(isd)-n0,"mapped":{k:str(hdr[v]).strip() for k,v in cm.items()}})
    for L in (inv,notes,isd):
        for r in L: r["inorm"]=inorm(r["inv"]); r["core"]=icore(r["inv"])
    return {"inv":inv,"notes":notes,"isd":isd,"imp":imp,"layout":layout,"warns":warns}

# ── BOOKS (purchase register) ────────────────────────────────────────────────
PB_COLS={
 "date": [r"^date$",r"^voucher date$",r"^posting date$"],
 "party":[r"^particulars$",r"^party( name)?$",r"^supplier( name)?$",r"^vendor( name)?$",r"^name of (the )?(party|supplier|vendor)$",r"^ledger( name)?$"],
 "vtype":[r"^voucher type$",r"^vch\.? type$",r"^doc(ument)? type$",r"^type$"],
 "vch":  [r"^voucher (no|number)\.?$",r"^vch\.? ?no\.?$",r"^doc(ument)? (no|number)\.?$"],
 "sinv": [r"^supplier inv(oice)?\.? ?(no|number|num)\.?$",r"^(party|vendor|bill|supplier) ?(bill|invoice|inv)\.? ?(no|number)\.?$",r"^ref(erence)?\.? ?(no|number)\.?$",r"^supplier ref",r"^invoice (no|number)\.?$",r"^bill (no|number)\.?$",r"^inv\.? ?(no|number)\.?$"],
 "sdate":[r"^supplier inv(oice)?\.? ?date$",r"^(party|vendor|bill|supplier) ?(bill|invoice|inv)\.? ?date$",r"^ref(erence)?\.? date$",r"^invoice date$",r"^bill date$"],
 "gstin":[r"^gstin(/uin)?$",r"^gst ?(no|number|in)\.?$",r"^(supplier|party|vendor) gstin(/uin)?$",r"gstin"],
 "gross":[r"^gross total$",r"^invoice (value|amount|total)$",r"^total( amount| value)?$",r"^bill (value|amount)$",r"^grand total$"],
 "narr": [r"^narration$",r"^remarks?$"],
}
PB_INFO=[r"^supplier",r"^vendor",r"^party",r"^bill",r"inv(oice)?\.? ?(no|num|date)",r"^ref",r"^pos$",r"^place of"]
LEDGER_CLASSES_P=["Taxable","IGST","CGST","SGST","Cess","RCM IGST","RCM CGST","RCM SGST","Other","Ignore"]
def guess_pledger(h):
    t=_hn(h)
    if re.search(r"rcm|reverse charge|\brc\b",t):
        if re.search(r"\bigst\b|integrated",t): return "RCM IGST"
        if re.search(r"\bcgst\b|central",t): return "RCM CGST"
        if re.search(r"\bsgst\b|\butgst\b|state",t): return "RCM SGST"
    if re.search(r"round(ing)?[ -]?(off|up|down)?\b|\btcs\b|\btds\b|freight (paid|payable)",t) and not re.search(r"purchase",t): return "Other"
    if re.search(r"\bcess\b",t): return "Cess"
    if re.search(r"\bigst\b|integrated (gst|tax)",t): return "IGST" if not re.search(_VAL_KW.replace("sale|sales|",""),t) or re.search(r"input|credit|itc|paid",t) else "Taxable"
    if re.search(r"\bcgst\b|central (gst|tax)",t): return "CGST" if not re.search(_VAL_KW.replace("sale|sales|",""),t) or re.search(r"input|credit|itc|paid",t) else "Taxable"
    if re.search(r"\bsgst\b|\butgst\b|state (gst|tax)",t): return "SGST" if not re.search(_VAL_KW.replace("sale|sales|",""),t) or re.search(r"input|credit|itc|paid",t) else "Taxable"
    return "Taxable"
def guess_pvtype(v):
    t=_hn(v)
    return "Credit/Debit Note" if re.search(r"debit note|dr\.? ?note|credit note|cr\.? ?note|purchase return|\bcn\b|\bdn\b|return",t) else "Invoice"

def scan_pbooks(fb):
    """Header-driven flat purchase register. Returns None if no flat layout is found (→ legacy parsers)."""
    wb=load_wb(fb); recs=[]; layout=[]; warns=[]
    for ws in wb.worksheets:
        rows=[list(r) for r in ws.iter_rows(values_only=True)]
        if not rows: continue
        hr,cm=_find_hdr(rows,PB_COLS,[],3)
        if hr is None or not ("sinv" in cm or "vch" in cm) or "party" not in cm: continue
        hdr=rows[hr]; data=rows[hr+1:]; used=set(cm.values()); lcols=[]; seen={}
        colv=lambda ci:[r[ci] if ci<len(r) else None for r in data]
        for ci,h in enumerate(hdr):
            hs=_hn(h)
            if not hs or ci in used or _is_info(hs) or any(re.search(p,hs) for p in PB_INFO): continue
            if not any(is_num(v) and num(v)!=0 for v in colv(ci)): continue
            name=str(h).strip()
            if name in seen: seen[name]+=1; name=f"{name} ({seen[name]})"
            else: seen[name]=1
            lcols.append((ci,name))
        n0=len(recs)
        for r in data:
            g=lambda f:(r[cm[f]] if f in cm and cm[f]<len(r) else None)
            if not isinstance(to_date(g("date")),datetime.date) and "date" in cm: 
                if str(g("party") or "").strip().lower() not in ("(cancelled )","cancelled"): continue
            party=str(g("party") or "").strip()
            if party.lower() in ("total","grand total","sub total",""): continue
            sinv=str(g("sinv") if g("sinv") is not None else "").strip(); vch=str(g("vch") if g("vch") is not None else "").strip()
            key=sinv or vch
            if not key: continue
            rec={"_sheet":ws.title,"_inv":key,"_sinv":sinv,"_vch":vch,"_date":to_date(g("sdate")) or to_date(g("date")),"_bdate":to_date(g("date")),
                 "_party":party,"_gstin":cgstin(g("gstin")),"_vtype":str(g("vtype") or "").strip(),
                 "_gross":(abs(num(g("gross"))) if "gross" in cm else np.nan),"_narr":str(g("narr") or "")}
            for ci,name in lcols: rec[name]=num(r[ci]) if ci<len(r) else 0.0
            recs.append(rec)
        layout.append({"sheet":ws.title,"header_row":hr+1,"rows":len(recs)-n0,"mapped":{k:str(hdr[v]).strip() for k,v in cm.items()},"ledgers":[n for _,n in lcols]})
        if "sinv" not in cm: warns.append(f"Sheet '{ws.title}': no Supplier Invoice No column found - Voucher No used for matching (will not match 2B unless client books supplier invoice no as voucher no).")
    if not recs: return None
    bw=pd.DataFrame(recs); led=[c for c in bw.columns if not c.startswith("_")]
    for c in led: bw[c]=pd.to_numeric(bw[c],errors="coerce").fillna(0.0)
    return {"bw":bw,"ledger_guess":{c:guess_pledger(c) for c in led},"vtype_guess":{v:guess_pvtype(v) for v in bw["_vtype"].unique()},"layout":layout,"warns":warns,"kind":"flat"}

def scan_pbooks_legacy(fb):
    """Tally sub-row / pivot formats via the original parser → invoice records (no taxable value)."""
    bd,det=parse_purchase(fb)
    recs=[]
    for d in det:
        inv=str(d.get("SupplierInv") or d.get("VchNo") or "").strip()
        recs.append({"_sheet":"legacy","_inv":inv,"_sinv":inv,"_vch":str(d.get("VchNo") or ""),"_date":to_date(d.get("Date")),"_bdate":to_date(d.get("Date")),
                     "_party":d.get("Party",""),"_gstin":d.get("GSTIN",""),"_vtype":"","_gross":np.nan,"_narr":"",
                     "Input IGST":d.get("IGST",0.0),"Input CGST":d.get("CGST",0.0),"Input SGST":d.get("SGST",0.0),
                     "_cat":d.get("Category","")})
    bw=pd.DataFrame(recs); 
    return {"bw":bw,"ledger_guess":{"Input IGST":"IGST","Input CGST":"CGST","Input SGST":"SGST"},"vtype_guess":{"":"Invoice"},
            "layout":[{"sheet":"(legacy Tally layout)","header_row":0,"rows":len(recs),"mapped":{},"ledgers":["Input IGST","Input CGST","Input SGST"]}],
            "warns":["Purchase register is in the older Tally sub-row/pivot layout: taxable value is not available, so only tax is compared. A flat export (one row per voucher with Supplier Invoice No + GSTIN columns) gives a stronger reco."],"kind":"legacy"}

def scan_pbooks_any(fb):
    s=scan_pbooks(fb)
    if s is not None: return s
    try: return scan_pbooks_legacy(fb)
    except Exception as e: raise ValueError("Could not read the purchase register.\nNeeded headers: Date · Particulars · Voucher No · Supplier Invoice No · GSTIN/UIN · Input IGST / CGST / SGST ledger columns.\n("+str(e)[:120]+")")

def build_pbooks(bw,ledger_class,vtype_class):
    led=[c for c in bw.columns if not c.startswith("_")]; d=bw.copy()
    def ssum(cls):
        cols=[c for c in led if ledger_class.get(c,"Taxable")==cls]
        return d[cols].sum(axis=1) if cols else pd.Series(0.0,index=d.index)
    for k,cl in (("taxable","Taxable"),("igst","IGST"),("cgst","CGST"),("sgst","SGST"),("cess","Cess"),("rigst","RCM IGST"),("rcgst","RCM CGST"),("rsgst","RCM SGST")): d[k]=ssum(cl)
    d["vclass"]=d["_vtype"].map(lambda v: vtype_class.get(v,"Invoice")); d=d[d.vclass!="Exclude"].copy()
    tot=d[led].abs().sum(axis=1)
    d["cancel"]=[("cancel" in p.lower()) or (t==0 and "cancel" in n.lower()) for p,t,n in zip(d._party,tot,d._narr)]
    return d

def _recs_from_books(d):
    out={"Invoice":[],"Credit/Debit Note":[]}
    for rr in d.to_dict("records"):
        out[rr["vclass"]].append({"gstin":rr["_gstin"],"party":rr["_party"],"inv":rr["_inv"],"inorm":inorm(rr["_inv"]),"core":icore(rr["_inv"]),
            "date":rr["_date"],"bdate":rr["_bdate"],"taxable":abs(rr["taxable"]),"igst":abs(rr["igst"]),"cgst":abs(rr["cgst"]),"sgst":abs(rr["sgst"]),"cess":abs(rr["cess"]),
            "rigst":abs(rr["rigst"]),"rcgst":abs(rr["rcgst"]),"rsgst":abs(rr["rsgst"]),"cancel":bool(rr["cancel"]),"vch":rr["_vch"],"vtype":rr["_vtype"],"gross":rr["_gross"],
            "src":rr["_sheet"],"cat":rr.get("_cat","")})
    return out

def _name_gstin_map(books,p2b):
    """fill blank GSTIN in books from (1) same party elsewhere in books (2) best word-overlap with 2B supplier names"""
    known={}
    for r in books:
        if r["gstin"]: known.setdefault(cname(r["party"]),r["gstin"])
    sup={}
    for r in p2b:
        sup.setdefault(r["gstin"],r["party"])
    sw={g:words(n) for g,n in sup.items()}
    GEN={"transport","carrier","carriers","roadlines","logistics","trading","enterprises","services","industries","pvt","ltd","private","limited"}
    filled=0
    for r in books:
        if r["gstin"]: continue
        k=cname(r["party"])
        if k in known: r["gstin"]=known[k]; r["gstin_src"]="same party in books"; filled+=1; continue
        bw=words(r["party"]); best=(0.0,None)
        for g,w in sw.items():
            ol=bw&w
            if not ol or (len(ol-GEN)==0 and len(ol)<=1): continue
            sc=len(ol)/len(bw|w)
            if sc>best[0]: best=(sc,g)
        if best[1] and best[0]>=0.5: r["gstin"]=best[1]; r["gstin_src"]="name match to 2B supplier"; known[k]=best[1]; filled+=1
    return filled

def _agg(recs):
    d={}
    for r in recs:
        k=(r["gstin"] or "~"+cname(r["party"]),r["inorm"])
        if k in d:
            a=d[k]
            for f in ("taxable","igst","cgst","sgst","cess","rigst","rcgst","rsgst"): a[f]+=r.get(f,0.0)
            a["n"]+=1; a["cancel"]=a["cancel"] and r.get("cancel",False); a["rcm"]=a.get("rcm",False) or r.get("rcm",False)
        else: a=dict(r); a["n"]=1; d[k]=a
    return list(d.values())

def _ttax(r): return r["igst"]+r["cgst"]+r["sgst"]+r["rigst"]+r["rcgst"]+r["rsgst"]

def match_itc(B,P):
    B=_agg(B); P=_agg(P); pairs=[]; ub=set(range(len(B))); up=set(range(len(P)))
    from collections import defaultdict
    def tier(kb,kp,how,fn_ok=None):
        mb=defaultdict(list); mp=defaultdict(list)
        for i in ub:
            k=kb(B[i]); 
            if k: mb[k].append(i)
        for j in up:
            ks=kp(P[j])
            for k in (ks if isinstance(ks,list) else [ks]):
                if k: mp[k].append(j)
        for k,bl in mb.items():
            pl=[j for j in mp.get(k,[]) if j in up]
            if len(bl)==1 and len(pl)==1 and bl[0] in ub:
                if fn_ok and not fn_ok(B[bl[0]],P[pl[0]]): continue
                pairs.append((bl[0],pl[0],how)); ub.discard(bl[0]); up.discard(pl[0])
    tier(lambda r:(r["gstin"],r["inorm"]) if r["gstin"] and r["inorm"] else None,
         lambda r:[(r["gstin"],r["inorm"])]+([(r["gstin"],inorm(r["oinv"]))] if r.get("oinv") else []),"exact")
    tier(lambda r:(r["gstin"],r["core"]) if r["gstin"] and r["core"] else None, lambda r:(r["gstin"],r["core"]) if r["core"] else None,"core")
    tier(lambda r:(r["gstin"],round(_ttax(r))) if r["gstin"] and _ttax(r)>0 else None, lambda r:(r["gstin"],round(_ttax(r))) if _ttax(r)>0 else None,"amount")
    tier(lambda r:(r["inorm"],round(_ttax(r))) if r["inorm"] and _ttax(r)>0 else None, lambda r:(r["inorm"],round(_ttax(r))) if r["inorm"] and _ttax(r)>0 else None,"gstin_typo")
    tier(lambda r:r["inorm"] if len(r["inorm"])>=4 else None, lambda r:r["inorm"] if len(r["inorm"])>=4 else None,"inv_only",
         lambda b,p: abs(_ttax(b)-_ttax(p))<=max(1.0,0.05*max(_ttax(b),_ttax(p))))
    return B,P,pairs,sorted(ub),sorted(up)

def _fmt(x): return f"{abs(x):,.0f}"
def reco_itc(Braw,Praw,tol,mat,has_taxable,doc="Invoice"):
    B,P,pairs,ubl,upl=match_itc(Braw,Praw); rows=[]
    z={"taxable":0.0,"igst":0.0,"cgst":0.0,"sgst":0.0,"cess":0.0,"rigst":0.0,"rcgst":0.0,"rsgst":0.0}
    def mk(b,p,how):
        b=b or {}; p=p or {}
        r={"B_Inv":b.get("inv",""),"P_Inv":p.get("inv",""),"Date":p.get("date") or b.get("date"),"B_Date":b.get("date"),"P_Date":p.get("date"),
           "Party":b.get("party") or p.get("party",""),"P_Party":p.get("party",""),"B_GSTIN":b.get("gstin",""),"P_GSTIN":p.get("gstin",""),"GSTIN":p.get("gstin") or b.get("gstin",""),
           "B_Tax":b.get("taxable",0.0),"P_Tax":p.get("taxable",0.0),"how":how,"src":p.get("src",""),"ntype":p.get("ntype",""),"avail":p.get("avail",""),"reason":p.get("reason",""),
           "B_VType":b.get("vtype",""),"Period":p.get("period",""),"has_b":bool(b),"has_p":bool(p),"amend":p.get("amend",False)}
        for h in HEADS: r["B_"+h.upper()]=b.get(h,0.0); r["P_"+h.upper()]=p.get(h,0.0)
        r["B_RCM"]=b.get("rigst",0.0)+b.get("rcgst",0.0)+b.get("rsgst",0.0); r["P_Cess"]=p.get("cess",0.0); r["B_Cess"]=b.get("cess",0.0)
        if r["B_RCM"]>0:   # RCM ledgers are compared head-wise with 2B tax
            for h,k in zip(HEADS,("rigst","rcgst","rsgst")): r["B_"+h.upper()]+=b.get(k,0.0)
        bg=r["B_IGST"]+r["B_CGST"]+r["B_SGST"]; pg=r["P_IGST"]+r["P_CGST"]+r["P_SGST"]
        r["B_GST"]=bg; r["P_GST"]=pg; r["elig"]=[0.0,0.0,0.0]
        pr="";rm="";st=""
        P_rcm=bool(p.get("rcm")); b_rcm=r["B_RCM"]>0
        big=lambda amt:"High" if abs(amt)>=mat else "Medium"
        if b and not p:
            if b.get("cancel") or (bg+r["B_RCM"]==0):
                st="Nil-tax / cancelled in books";pr="Low";rm="Cancelled / nil-tax voucher in books - no ITC involved"
            elif b_rcm:
                st="RCM - verify";pr="Medium";rm=f"RCM entry in books (Rs.{_fmt(r['B_RCM'])}) - confirm self-invoice, tax paid in 3B 3.1(d) in cash, ITC in 4(A)(3). Supplier is unregistered/not on 2B"
            else:
                st="Only in Books";pr=big(bg);rm=f"Booked (ITC Rs.{_fmt(bg)}) but NOT in 2B - ITC not claimable now (Sec 16(2)(aa)/Rule 36(4)); if already claimed in 3B reverse it; follow up with supplier / check next-month 2B"
        elif p and not b:
            if P_rcm:
                st="RCM - verify";pr=big(pg);rm=f"Supplier reports reverse charge (tax Rs.{_fmt(pg)}) - client must book self-invoice, pay tax in cash 3.1(d), then claim 4(A)(3). Not in books"
            elif str(p.get("avail","")).strip().lower().startswith("n"):
                st="2B: ITC not available";pr="Low";rm="2B shows ITC NOT available"+(f": {p['reason']}" if p.get("reason") else "")+" - not in books, nothing to claim"
            else:
                st="Only in 2B";pr=big(pg);rm=f"Supplier filed (ITC Rs.{_fmt(pg)}) but not in books - book if goods/services received; check cut-off / other register / GSTIN typo in books"
        else:
            iss=[];bigdiff=False;over=False
            ndiff=[(h.upper(),r["B_"+h.upper()]-r["P_"+h.upper()]) for h in HEADS]
            for h,dv in ndiff:
                if abs(dv)>tol: iss.append(f"{h} differs by Rs.{_fmt(dv)} ({'books>2B' if dv>0 else '2B>books'})"); bigdiff=True; over=over or dv>0
            if abs(bg-pg)<=tol and bigdiff: iss=[f"Tax HEAD differs (Books IGST/CGST/SGST = {r['B_IGST']:,.0f}/{r['B_CGST']:,.0f}/{r['B_SGST']:,.0f} vs 2B {r['P_IGST']:,.0f}/{r['P_CGST']:,.0f}/{r['P_SGST']:,.0f}) - place-of-supply / wrong ledger; total tax agrees"]
            if has_taxable and b.get("taxable",0)>0 and p.get("taxable",0)>0 and abs(b["taxable"]-p["taxable"])>tol:
                iss.append(f"Taxable differs by Rs.{_fmt(b['taxable']-p['taxable'])}"); bigdiff=True
            if b.get("gstin") and p.get("gstin") and b["gstin"]!=p["gstin"]: iss.append(f"GSTIN differs (Books {b['gstin']} / 2B {p['gstin']})"); bigdiff=True
            bd_,pd_=b.get("date"),p.get("date")
            if b.get("bdate") and p.get("date") and b.get("bdate")!=b.get("date") and bd_ and pd_ and bd_!=pd_: iss.append(f"Invoice date differs (Books {bd_:%d-%b-%y} / 2B {pd_:%d-%b-%y})")
            if P_rcm and not b_rcm: st="RCM - verify";pr="High";rm="Supplier reports REVERSE CHARGE but client booked it as regular ITC - ITC not allowed unless RCM tax is paid"
            elif b.get("cancel") or (bg+r["B_RCM"]==0 and pg>0):
                st="Mismatch";pr="High";rm=f"Books show nil/cancelled but 2B has tax Rs.{_fmt(pg)}"+(" | Inv No: Books "+r["B_Inv"]+" / 2B "+r["P_Inv"] if how!="exact" else "")
            elif bigdiff:
                st="Mismatch";pr=big(bg-pg) if over else "Medium";rm="; ".join(iss)+(f" | Inv No: Books {r['B_Inv']} / 2B {r['P_Inv']}" if how not in ("exact",) else "")
                if p.get("amend"): rm+=f" | 2B amendment (B2BA) of {p.get('oinv','')} - supplier revised after booking"
            elif iss: st="Matched - Rounding" if False else "Matched"; pr="Low"; rm="; ".join(iss)
            else:
                small=abs(bg-pg)>0.005 or any(abs(r["B_"+h.upper()]-r["P_"+h.upper()])>0.005 for h in HEADS)
                st="Matched - Rounding" if small else ("Matched" if how=="exact" else "Matched - Inv No differs")
                if how!="exact": 
                    st="Matched - Inv No differs"; pr="Low"; rm=f"Matched on {'GSTIN + numeric part of invoice no' if how=='core' else ('GSTIN + tax amount' if how=='amount' else 'invoice no + amount (GSTIN differs/typo)')}: Books '{r['B_Inv']}' vs 2B '{r['P_Inv']}'"
                    if how=="gstin_typo": st="Mismatch"; pr="Medium"; rm=f"GSTIN in books ({b.get('gstin') or 'blank'}) differs from 2B ({p.get('gstin')}) - correct supplier GSTIN in books"
                elif small: pr="Low"; rm="Within rounding"
                if P_rcm and b_rcm: st="RCM - verify";pr="Medium";rm="RCM booked correctly - confirm tax paid in 3.1(d) and ITC taken in 4(A)(3)"
                if str(p.get("avail","")).strip().lower().startswith("n") and st.startswith("Matched"):
                    st="2B: ITC not available";pr="High";rm="Booked in books but 2B says ITC NOT available"+(f": {p['reason']}" if p.get("reason") else "")+" - do not claim"
            if st in ("Matched","Matched - Inv No differs","Matched - Rounding","Mismatch") and not (str(p.get("avail","")).strip().lower().startswith("n")):
                r["elig"]=[min(r["B_IGST"],r["P_IGST"]),min(r["B_CGST"],r["P_CGST"]),min(r["B_SGST"],r["P_SGST"])]
        r["Status"]=st;r["Priority"]=pr;r["Remark"]=rm
        return r
    for bi,pj,how in pairs: rows.append(mk(B[bi],P[pj],how))
    for bi in ubl: rows.append(mk(B[bi],None,""))
    for pj in upl: rows.append(mk(None,P[pj],""))
    df=pd.DataFrame(rows)
    if df.empty: return df
    od={s:i for i,s in enumerate(ITC_STAT)}; df["_o"]=df.Status.map(od).fillna(99)
    df["_d"]=[d if isinstance(d,datetime.date) else datetime.date(2100,1,1) for d in df.Date]
    return df.sort_values(["_o","Party","_d"]).drop(columns=["_o","_d"]).reset_index(drop=True)

def run_itc_v2(pr_bytes,b2b_bytes,ledger_class=None,vtype_class=None,tol=1.0,mat=10000.0,scan=None,pt=None):
    pt=pt or parse_2b_full(b2b_bytes); scan=scan or scan_pbooks_any(pr_bytes)
    lc=dict(scan["ledger_guess"]); lc.update(ledger_class or {}); vc=dict(scan["vtype_guess"]); vc.update(vtype_class or {})
    d=build_pbooks(scan["bw"],lc,vc); rec=_recs_from_books(d)
    allp=pt["inv"]+pt["notes"]
    filled=_name_gstin_map(rec["Invoice"]+rec["Credit/Debit Note"],allp)
    has_tx=(scan["kind"]=="flat") and any(r["taxable"]>0 for r in rec["Invoice"])
    inv=reco_itc(rec["Invoice"],pt["inv"],tol,mat,has_tx,"Invoice"); notes=reco_itc(rec["Credit/Debit Note"],pt["notes"],tol,mat,has_tx,"Note")
    imp_books=[r for r in rec["Invoice"] if "custom" in str(r.get("cat","")).lower() or re.search(r"custom|boe|bill of entry|import",str(r["party"]).lower()) and not r["gstin"]]
    warns=list(scan["warns"])+list(pt["warns"])
    if filled: warns.append(f"{filled} book voucher(s) had no GSTIN - supplier GSTIN was inferred from the party name (listed in Mapping & Checks; verify).")
    if not rec["Credit/Debit Note"] and len(pt["notes"]): warns.append(f"{len(pt['notes'])} credit/debit note(s) on 2B but no debit-note vouchers found in the purchase register - provide the Debit Note (purchase return) register to reconcile.")
    return {"inv":inv,"notes":notes,"pt":pt,"scan":scan,"ledger_class":lc,"vtype_class":vc,"books":d,"rec":rec,"tol":tol,"mat":mat,"warns":warns,"imp_books":imp_books,"filled":filled}

# ── ITC v2 Excel output ──────────────────────────────────────────────────────
def _xl_kit(wb):
    def title(ws,txt,nc,r=1):
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=nc); c=ws.cell(r,1,txt)
        c.font=Font(name="Calibri",bold=True,size=12,color="FFFFFF"); c.fill=F(TITLE); c.alignment=CTR(); ws.row_dimensions[r].height=26
    def sub(ws,txt,nc,r=2):
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=nc)
        c=ws.cell(r,1,f"Created by: Purva Doshi  |  linkedin.com/in/purvadoshi26  |  {txt}"); c.font=Font(name="Calibri",size=8,italic=True,color="7F9FC0"); c.alignment=CTR()
    def head(ws,r,hs,color=NAVY):
        for ci,h in enumerate(hs,1):
            c=ws.cell(r,ci,h); c.font=HF(9); c.fill=F(color); c.alignment=CTR(); c.border=BD()
        ws.row_dimensions[r].height=28
    def widths(ws,w):
        for i,x in enumerate(w,1): ws.column_dimensions[get_column_letter(i)].width=x
    return title,sub,head,widths

def build_itc_excel_v2(res,client,period):
    wb=openpyxl.Workbook(); title,sub,head,widths=_xl_kit(wb)
    inv=res["inv"]; notes=res["notes"]; pt=res["pt"]; tol=res["tol"]
    dv=lambda x: None if (x is None or (not isinstance(x,datetime.date) and pd.isna(x))) else x
    # ---------- Invoice Reco ----------
    wi=wb.active; wi.title="Invoice Reco"; wi.sheet_view.showGridLines=False
    H=["#","Inv No (Books)","Inv No (2B)","Date","Party","GSTIN","Books Taxable","2B Taxable","Books IGST","2B IGST","Books CGST","2B CGST","Books SGST","2B SGST",
       "Books GST","2B GST","GST Diff (Books-2B)","Eligible IGST","Eligible CGST","Eligible SGST","Status","Priority","Match basis","2B ITC avail.","Remark"]
    title(wi,f"{client}  -  GSTR-2B vs Purchase Register  |  Invoice-wise ITC Reco  |  {period}",len(H)); sub(wi,"Diffs are live formulas. Eligible = lower of Books and 2B, head-wise",len(H)); head(wi,4,H)
    r0=5; basis={"exact":"GSTIN + Inv No","core":"GSTIN + inv no (numeric part)","amount":"GSTIN + tax amount","gstin_typo":"Inv No + amount (GSTIN differs)","inv_only":"Inv No only",
                 "":"-"}
    for i,r in inv.iterrows():
        x=r0+i; el=r["elig"]
        vals=[i+1,r.B_Inv,r.P_Inv,dv(r.Date),r.Party,r.GSTIN,r.B_Tax,r.P_Tax,r.B_IGST,r.P_IGST,r.B_CGST,r.P_CGST,r.B_SGST,r.P_SGST,f"=I{x}+K{x}+M{x}",f"=J{x}+L{x}+N{x}",f"=O{x}-P{x}",
              el[0],el[1],el[2],r.Status,r.Priority,basis.get(r.how,r.how) if r.has_b and r.has_p else "-",r.avail,r.Remark]
        for ci,v in enumerate(vals,1):
            c=wi.cell(x,ci,v); c.font=DF(9); c.border=BD(); c.fill=F(ITC_BG.get(r.Status,"FFFFFF"))
            if 7<=ci<=20: c.number_format=INR; c.alignment=RGT()
            elif ci==4: c.number_format="DD-MMM-YY"; c.alignment=CTR()
            elif ci in (1,21,22,24): c.alignment=CTR()
        wi.cell(x,21).font=DF(9,bold=True)
    last=r0+max(len(inv),1)-1; tr=last+1
    wi.cell(tr,5,"TOTAL").font=DF(9,bold=True)
    for ci in range(7,21):
        c=wi.cell(tr,ci,f"=SUM({get_column_letter(ci)}{r0}:{get_column_letter(ci)}{last})"); c.font=DF(9,bold=True); c.number_format=INR; c.fill=F("D9E1F2"); c.border=BD()
    widths(wi,[5,18,18,11,34,18,14,14,12,12,12,12,12,12,13,13,13,12,12,12,26,9,24,10,80]); wi.freeze_panes="D5"; wi.auto_filter.ref=f"A4:{get_column_letter(len(H))}{last}"
    RG=lambda col:f"'Invoice Reco'!${col}${r0}:${col}${last}"
    # ---------- Credit / Debit notes ----------
    wn=wb.create_sheet("Credit-Debit Notes"); wn.sheet_view.showGridLines=False
    NH=["#","Note No (Books)","Note No (2B)","Date","Party","GSTIN","2B Type (C/D)","Books Taxable","2B Taxable","Books GST","2B GST","GST Diff","Status","Priority","Remark"]
    title(wn,f"{client}  -  Credit / Debit Notes (2B CDNR vs Debit-Note vouchers)  |  {period}",len(NH)); sub(wn,"Absolute values compared. 2B credit note = ITC must be reduced (Sec 34(2))",len(NH)); head(wn,4,NH)
    n0=5
    for i,r in notes.iterrows():
        x=n0+i
        vals=[i+1,r.B_Inv,r.P_Inv,dv(r.Date),r.Party,r.GSTIN,r.ntype,r.B_Tax,r.P_Tax,r.B_GST,r.P_GST,f"=J{x}-K{x}",r.Status,r.Priority,r.Remark]
        for ci,v in enumerate(vals,1):
            c=wn.cell(x,ci,v); c.font=DF(9); c.border=BD(); c.fill=F(ITC_BG.get(r.Status,"FFFFFF"))
            if 8<=ci<=12: c.number_format=INR; c.alignment=RGT()
            elif ci==4: c.number_format="DD-MMM-YY"; c.alignment=CTR()
            elif ci in (1,7,13,14): c.alignment=CTR()
    nl=n0+max(len(notes),1)-1
    if len(notes):
        wn.cell(nl+1,5,"TOTAL").font=DF(9,bold=True)
        for ci in range(8,13):
            c=wn.cell(nl+1,ci,f"=SUM({get_column_letter(ci)}{n0}:{get_column_letter(ci)}{nl})"); c.font=DF(9,bold=True); c.number_format=INR; c.fill=F("D9E1F2")
    widths(wn,[5,18,18,11,34,18,8,14,14,13,13,13,26,9,80]); wn.freeze_panes="D5"
    # ---------- Summary ----------
    sm=wb.create_sheet("Summary",0); sm.sheet_view.showGridLines=False
    title(sm,f"{client}  -  GSTR-2B vs Books  |  ITC Reco Summary  |  {period}",7); sub(sm,f"ITC Reco | {period} | tolerance Rs.{tol:g} | high-priority threshold Rs.{res['mat']:,.0f}",7)
    head(sm,4,["A. Invoices - status","Count","Books GST","2B GST","Diff (Books-2B)","Eligible ITC (lower of both)","Meaning"])
    for i,s in enumerate(ITC_STAT):
        x=5+i; sm.cell(x,1,s); sm.cell(x,2,f'=COUNTIF({RG("U")},A{x})'); sm.cell(x,3,f'=SUMIF({RG("U")},A{x},{RG("O")})'); sm.cell(x,4,f'=SUMIF({RG("U")},A{x},{RG("P")})')
        sm.cell(x,5,f"=C{x}-D{x}"); sm.cell(x,6,f'=SUMIF({RG("U")},A{x},{RG("R")})+SUMIF({RG("U")},A{x},{RG("S")})+SUMIF({RG("U")},A{x},{RG("T")})'); sm.cell(x,7,ITC_MEAN[s])
    xt=5+len(ITC_STAT); sm.cell(xt,1,"TOTAL")
    for ci in range(2,7): sm.cell(xt,ci,f"=SUM({get_column_letter(ci)}5:{get_column_letter(ci)}{xt-1})")
    for x in range(5,xt+1):
        for ci in range(1,8):
            c=sm.cell(x,ci); c.border=BD(); c.font=DF(9,bold=(x==xt or ci==1))
            if 3<=ci<=6: c.number_format=INR; c.alignment=RGT()
            if ci==2: c.alignment=CTR()
            if x==xt: c.fill=F("D9E1F2")
            elif ci==1: c.fill=F(ITC_BG.get(sm.cell(x,1).value,"FFFFFF"))
    q=xt+2
    head(sm,q,["B. Tie-out (all tax heads)","","Amount (Rs.)","","","",""])
    tie=[("Total ITC per Books (invoices)",f"='Invoice Reco'!O{tr}"),("Total ITC per 2B (invoices incl. RCM / not-available rows)",f"='Invoice Reco'!P{tr}"),("Difference (Books - 2B)",f"=C{q+1}-C{q+2}"),
         ("Explained by status rows above",f"=E{xt}"),("Unexplained (must be 0)",f"=C{q+3}-C{q+4}")]
    for i,(a,b) in enumerate(tie):
        sm.cell(q+1+i,1,a); sm.cell(q+1+i,3,b).number_format=INR
        for ci in (1,2,3): sm.cell(q+1+i,ci).border=BD(); sm.cell(q+1+i,ci).font=DF(9,bold=(i in (2,4)))
    # indicative ITC for 3B
    q2=q+len(tie)+2
    head(sm,q2,["C. Indicative ITC for GSTR-3B Table 4(A)","IGST","CGST","SGST","Total","","Basis"])
    cn=pt["notes"]; cnC=[r for r in cn if r["ntype"]!="D"]; cnD=[r for r in cn if r["ntype"]=="D"]
    s_=lambda L,h: float(sum(r[h] for r in L))
    isd=pt["isd"]; impg=float(sum(r["IGST"] for r in pt["imp"]))
    rcm_rows=inv[(inv.Status=="RCM - verify")&inv.has_b]
    lines=[("Eligible ITC on invoices (lower of Books & 2B)",[f"=SUM({RG('R')})",f"=SUM({RG('S')})",f"=SUM({RG('T')})"],"Only invoices present in BOTH books and 2B (Rule 36(4)), excl. ITC-not-available & RCM rows"),
           ("Less: credit notes on 2B",[-s_(cnC,"igst"),-s_(cnC,"cgst"),-s_(cnC,"sgst")],"Supplier credit notes - recipient must reduce ITC"),
           ("Add: debit notes on 2B",[s_(cnD,"igst"),s_(cnD,"cgst"),s_(cnD,"sgst")],"Supplier debit notes"),
           ("Add: ISD credit (4(A)(4))",[s_(isd,"igst"),s_(isd,"cgst"),s_(isd,"sgst")],"From 2B ISD sheet"),
           ("Add: Import of goods (4(A)(1))",[impg,0,0],"2B IMPG sheet (Bill of Entry IGST)"),
           ("RCM ITC (4(A)(3)) - only after tax paid in cash",[float(rcm_rows.B_IGST.sum()),float(rcm_rows.B_CGST.sum()),float(rcm_rows.B_SGST.sum())],"Books RCM entries - confirm payment through 3.1(d) first")]
    for i,(a,hs,b) in enumerate(lines):
        x=q2+1+i; sm.cell(x,1,a)
        for k,v in enumerate(hs): sm.cell(x,2+k,v)
        sm.cell(x,5,f"=SUM(B{x}:D{x})"); sm.cell(x,7,b)
        for ci in range(1,8):
            c=sm.cell(x,ci); c.border=BD(); c.font=DF(9,bold=(ci==1))
            if 2<=ci<=5: c.number_format=INR; c.alignment=RGT()
    x=q2+len(lines)+1; sm.cell(x,1,"Indicative net ITC (excl. RCM line)")
    for ci in range(2,6):
        L=get_column_letter(ci); sm.cell(x,ci,f"=SUM({L}{q2+1}:{L}{q2+5})")
    for ci in range(1,8): c=sm.cell(x,ci); c.border=BD(); c.font=DF(9,bold=True); c.fill=F("D9E1F2"); 
    for ci in range(2,6): sm.cell(x,ci).number_format=INR
    sm.cell(x,7,"Compare with 3B 4(A)(1)+(3)+(4)+(5) less 4(B) reversals. Verify Sec 16(4), 17(5), Rule 37, 42, 43 separately.")
    # notes status
    q3=x+2
    head(sm,q3,["D. Credit / Debit notes","Count","Books GST","2B GST","Diff","",""])
    nst=[s for s in ITC_STAT if len(notes) and (notes.Status==s).any()]
    NR=lambda col:f"'Credit-Debit Notes'!${col}${n0}:${col}${nl}"
    for i,s in enumerate(nst):
        x=q3+1+i; sm.cell(x,1,s); sm.cell(x,2,f'=COUNTIF({NR("M")},A{x})'); sm.cell(x,3,f'=SUMIF({NR("M")},A{x},{NR("J")})'); sm.cell(x,4,f'=SUMIF({NR("M")},A{x},{NR("K")})'); sm.cell(x,5,f"=C{x}-D{x}")
        for ci in range(1,8):
            c=sm.cell(x,ci); c.border=BD(); c.font=DF(9,bold=(ci==1))
            if 3<=ci<=5: c.number_format=INR; c.alignment=RGT()
    if not nst: sm.cell(q3+1,1,"No credit/debit notes on 2B or in books").font=DF(9,c="777777")
    q4=q3+max(len(nst),1)+2
    sm.cell(q4,1,"E. Key observations").font=Font(name="Calibri",bold=True,size=11,color=NAVY)
    vc=inv.Status.value_counts().to_dict(); obs=[]
    ok=vc.get("Matched",0)+vc.get("Matched - Inv No differs",0)+vc.get("Matched - Rounding",0)
    obs.append(f"{ok} of {len(inv)} invoice rows reconcile on GSTIN + number + tax.")
    for s,t in [("Mismatch","invoice(s) with tax / taxable / head / GSTIN differences - see Exceptions"),("Only in 2B","invoice(s) on 2B not booked - book before claiming ITC"),
                ("Only in Books","invoice(s) booked but missing in 2B - ITC not claimable, reverse if already taken"),
                ("2B: ITC not available","invoice(s) where 2B itself says ITC not available - do not claim"),("RCM - verify","reverse-charge item(s) - self-invoice + cash payment first")]:
        if vc.get(s,0):
            amt=inv[inv.Status==s]; obs.append(f"{vc[s]} {t} (Books GST Rs.{amt.B_GST.sum():,.0f} / 2B GST Rs.{amt.P_GST.sum():,.0f}).")
    hd=inv[inv.Priority=="High"]
    if len(hd): obs.append(f"{len(hd)} HIGH-priority row(s) (>= Rs.{res['mat']:,.0f} or ITC-risk items).")
    if len(pt["imp"]): obs.append(f"IMPG: {len(pt['imp'])} Bill(s) of Entry, IGST Rs.{impg:,.0f} - reconcile with import / custom-duty ledger.")
    for w in res["warns"]: obs.append("WARNING: "+w)
    for i,t in enumerate(obs):
        x=q4+1+i; c=sm.cell(x,1,t); c.font=DF(9); c.alignment=Alignment(wrap_text=True,vertical="top"); sm.merge_cells(start_row=x,start_column=1,end_row=x,end_column=7); sm.row_dimensions[x].height=28
    widths(sm,[52,14,16,16,16,22,70])
    # ---------- Auditor Actions ----------
    wa=wb.create_sheet("Auditor Actions",1); wa.sheet_view.showGridLines=False
    title(wa,f"AUDITOR ACTION ITEMS - {client} | {period}",5); head(wa,3,["Priority","Action required","Detail","Amount (Rs.)","Legal ref"])
    acts=[]
    g=lambda s:inv[inv.Status==s]
    m=g("Mismatch")
    if len(m): acts.append(("HIGH",f"Resolve {len(m)} mismatched invoice(s)","Tax / taxable / head / GSTIN differ between books and 2B. Claim lower of the two until fixed",m[["B_GST","P_GST"]].min(axis=1).sum()*0+abs(m.B_GST-m.P_GST).sum(),"Sec 16(2)(aa)"))
    o2=g("Only in 2B")
    if len(o2): acts.append(("HIGH",f"Book {len(o2)} invoice(s) that suppliers filed",f"On 2B, not in books - verify receipt of goods/services, then book (ITC available only on booking)",o2.P_GST.sum(),"Sec 16(2)(b)"))
    ob=g("Only in Books")
    if len(ob): acts.append(("HIGH",f"Reverse / hold ITC on {len(ob)} invoice(s) not in 2B","If already claimed in 3B, reverse with interest; follow up supplier GSTR-1 filing",ob.B_GST.sum(),"Rule 36(4), Sec 41"))
    na=g("2B: ITC not available")
    if len(na): acts.append(("MED",f"{len(na)} invoice(s) flagged 'ITC not available' on 2B","Check reason (POS, Sec 16(4), 17(5), supplier non-compliance); do not claim",na.P_GST.sum(),"Sec 16/17(5)"))
    rc=g("RCM - verify")
    if len(rc): acts.append(("MED",f"Verify {len(rc)} RCM item(s)","Self-invoice raised? tax paid in cash in 3.1(d)? ITC in 4(A)(3) only after payment",rc[["B_GST","P_GST"]].max(axis=1).sum(),"Sec 9(3)/9(4), 16"))
    if len(notes) and (notes.Status!="Matched").any(): acts.append(("MED","Credit/debit notes not reconciled","Supplier credit notes must reduce ITC; debit notes issued by client need supplier-side reporting",abs(notes.B_GST-notes.P_GST).sum(),"Sec 34(2)"))
    if len(pt["imp"]): acts.append(("LOW","Reconcile imports (IMPG) with Bills of Entry","IGST on imports is claimed from 2B IMPG; match to import ledger / customs duty entries",impg,"Sec 16, 3(7) CTA"))
    acts.append(("OK",f"{ok} invoice(s) matched","Eligible after Sec 16(2)(a)-(d), 16(4), 17(5) checks",inv[inv.Status.str.startswith("Matched")].P_GST.sum(),"Sec 16"))
    pb={"HIGH":MIS,"MED":CUST_C,"LOW":ONLY2,"OK":MATCH}
    for i,it in enumerate(acts):
        x=4+i
        for ci,v in enumerate(it,1):
            c=wa.cell(x,ci,v); c.border=BD(); c.font=DF(9,bold=(ci<=2)); c.alignment=CTR() if ci in(1,4,5) else Alignment(wrap_text=True,vertical="center")
            if ci==4: c.number_format=INR
            if ci<=2: c.fill=F(pb[it[0]])
        wa.row_dimensions[x].height=30
    wa.cell(5+len(acts),1,"Also verify: Sec 16(4) time limit, Sec 17(5) blocked credits, Rule 37 (180-day payment), Rule 42/43 reversals, ITC on capital goods, 3B Table 4(B)/(D).").font=Font(name="Calibri",size=8,italic=True,color="666666")
    widths(wa,[10,46,70,16,20])
    # ---------- Party summary ----------
    wp=wb.create_sheet("Party Summary"); wp.sheet_view.showGridLines=False
    title(wp,f"{client}  -  Supplier-wise ITC summary  |  {period}",9); head(wp,3,["Party","GSTIN","Docs","Books GST","2B GST","Diff","Eligible ITC","Exceptions","Worst status"])
    gp=inv.groupby("GSTIN",dropna=False)
    rows_=[]
    for gs,d in gp:
        w=min(d.Status,key=lambda s:ITC_STAT.index(s) if s in ITC_STAT else 99) if False else None
        sev=[s for s in ["Mismatch","Only in Books","Only in 2B","RCM - verify","2B: ITC not available"] if (d.Status==s).any()]
        rows_.append((d.Party.iloc[0],gs,len(d),sev[0] if sev else "Matched"))
    rows_.sort(key=lambda t:(t[3]=="Matched",t[0].lower()))
    for i,(pn,gs,n,w) in enumerate(rows_):
        x=4+i; wp.cell(x,1,pn); wp.cell(x,2,gs); wp.cell(x,3,n)
        wp.cell(x,4,f'=SUMIF({RG("F")},B{x},{RG("O")})'); wp.cell(x,5,f'=SUMIF({RG("F")},B{x},{RG("P")})'); wp.cell(x,6,f"=D{x}-E{x}")
        wp.cell(x,7,f'=SUMIF({RG("F")},B{x},{RG("R")})+SUMIF({RG("F")},B{x},{RG("S")})+SUMIF({RG("F")},B{x},{RG("T")})')
        wp.cell(x,8,f'=COUNTIFS({RG("F")},B{x},{RG("U")},"<>Matched*")-COUNTIFS({RG("F")},B{x},{RG("U")},"Nil*")'); wp.cell(x,9,w)
        for ci in range(1,10):
            c=wp.cell(x,ci); c.border=BD(); c.font=DF(9)
            if 4<=ci<=7: c.number_format=INR; c.alignment=RGT()
            if w!="Matched": c.fill=F(ITC_BG.get(w,"FFFFFF")) if ci==9 else PatternFill()
    widths(wp,[38,18,7,15,15,15,15,11,22]); wp.freeze_panes="A4"
    # ---------- Exceptions ----------
    we=wb.create_sheet("Exceptions"); we.sheet_view.showGridLines=False
    EH=["Priority","Doc","Status","Inv / Note No","Date","Party","GSTIN","Books GST","2B GST","Remark"]
    title(we,f"{client}  -  Exceptions needing action  |  {period}",len(EH)); head(we,3,EH)
    ex=[]
    for df_,doc in ((inv,"Invoice"),(notes,"Note")):
        for r in df_.itertuples():
            if r.Status in ("Matched","Matched - Inv No differs","Matched - Rounding","Nil-tax / cancelled in books"): continue
            ex.append((r.Priority,doc,r.Status,r.B_Inv or r.P_Inv,dv(r.Date),r.Party,r.GSTIN,r.B_GST,r.P_GST,r.Remark))
    po={"High":0,"Medium":1,"Low":2,"":3}; ex.sort(key=lambda t:(po.get(t[0],3),t[1],t[2]))
    for i,rw in enumerate(ex):
        for ci,v in enumerate(rw,1):
            c=we.cell(4+i,ci,v); c.font=DF(9,bold=(ci==1),c=("C00000" if ci==1 and rw[0]=="High" else "222222")); c.border=BD()
            if ci in (8,9): c.number_format=INR; c.alignment=RGT()
            if ci==5: c.number_format="DD-MMM-YY"
    widths(we,[9,9,26,20,11,34,18,14,14,100]); we.freeze_panes="A4"; we.auto_filter.ref=f"A3:J{3+max(len(ex),1)}"
    # ---------- IMPG / ISD ----------
    if pt["imp"] or pt["isd"]:
        wg=wb.create_sheet("IMPG & ISD"); wg.sheet_view.showGridLines=False
        title(wg,f"{client}  -  Imports (IMPG) and ISD credit per 2B  |  {period}",7); r=3
        if pt["imp"]:
            head(wg,r,["Ref date","Port","BE No","BE Date","Taxable","IGST","Sheet"]); 
            for i,d in enumerate(pt["imp"]):
                for ci,k in enumerate(["Ref Date","Port","BE No","BE Date","Taxable","IGST","Sheet"],1):
                    c=wg.cell(r+1+i,ci,d[k]); c.font=DF(9); c.border=BD()
                    if k in("Taxable","IGST"): c.number_format=INR
            r+=len(pt["imp"])+3
        if pt["isd"]:
            head(wg,r,["ISD GSTIN","Name","Doc No","Date","IGST","CGST","SGST"])
            for i,d in enumerate(pt["isd"]):
                for ci,v in enumerate([d["gstin"],d["party"],d["inv"],dv(d["date"]),d["igst"],d["cgst"],d["sgst"]],1):
                    c=wg.cell(r+1+i,ci,v); c.font=DF(9); c.border=BD()
                    if ci>=5: c.number_format=INR
        widths(wg,[20,30,18,14,16,16,16])
    # ---------- Mapping & Checks ----------
    wm=wb.create_sheet("Mapping & Checks"); wm.sheet_view.showGridLines=False
    title(wm,f"{client}  -  What the app detected + integrity checks  |  {period}",6); r=[3]
    def block(t): c=wm.cell(r[0],1,t); c.font=Font(name="Calibri",bold=True,size=11,color=NAVY); r[0]+=1
    def row_(vals,hdr=False,fmt=None):
        for ci,v in enumerate(vals,1):
            c=wm.cell(r[0],ci,v); c.border=BD()
            if hdr: c.font=HF(9); c.fill=F(NAVY); c.alignment=CTR()
            else:
                c.font=DF(9)
                if fmt and ci in fmt: c.number_format=fmt[ci]; c.alignment=RGT()
        r[0]+=1
    sc=res["scan"]; bw=sc["bw"]
    block("1. Purchase register - columns found (by header name)"); row_(["Sheet","Header row","Voucher rows","Field -> header used"],True)
    for L_ in sc["layout"]: row_([L_["sheet"],L_["header_row"],L_["rows"],"; ".join(f"{k}: {v}" for k,v in L_["mapped"].items())])
    r[0]+=1; block("2. Ledger columns - class used (Taxable / IGST / CGST / SGST / Cess / RCM IGST-CGST-SGST / Other / Ignore)"); row_(["Ledger (column header)","Class used","Auto-guess","Total in file (Rs.)"],True)
    for k,v in res["ledger_class"].items(): row_([k,v,sc["ledger_guess"].get(k,""),float(bw[k].sum()) if k in bw else 0.0],fmt={4:INR})
    r[0]+=1; block("3. Voucher types"); row_(["Voucher type","Treated as","Vouchers"],True)
    for k,v in res["vtype_class"].items(): row_([k or "(blank)",v,int((bw._vtype==k).sum())])
    r[0]+=1; block("4. GSTR-2B file - sheets and columns found"); row_(["Sheet","Role","Header row","Rows","Field -> header used"],True)
    for L_ in pt["layout"]: row_([L_["sheet"],L_["role"],L_["header_row"],L_["rows"],"; ".join(f"{k}: {v}" for k,v in L_["mapped"].items())])
    r[0]+=1; block("5. Books internal checks"); row_(["Check","Result","Detail"],True)
    d=res["books"]; d=d[d.vclass=="Invoice"]
    bad_g=[x for x in d._gstin.unique() if x and not re.match(r"^\d{2}[A-Z]{5}\d{4}[A-Z]\d[A-Z0-9][A-Z0-9]$",x)]
    row_(["GSTIN format valid (15-char pattern)",f"{len(bad_g)} invalid","; ".join(bad_g[:8]) if bad_g else "OK"])
    row_(["Vouchers with blank GSTIN",f"{int((d._gstin=='').sum())}","GSTIN inferred from party name for "+str(res['filled'])+" of them" if res['filled'] else ("OK" if not (d._gstin=='').sum() else "Unmatched to 2B - may be unregistered / RCM / imports")])
    allp=set(x["gstin"] for x in pt["inv"]+pt["notes"])
    notin=sorted(set(x for x in d._gstin.unique() if x and x not in allp))
    row_(["Books suppliers with no document on 2B at all",f"{len(notin)}","Supplier hasn't filed / wrong GSTIN in books: "+"; ".join(notin[:6])+(" ..." if len(notin)>6 else "") if notin else "OK"])
    dupb=d.groupby(["_gstin","_inv"]).size(); dupb=dupb[dupb>1]
    row_(["Duplicate supplier invoice no (same GSTIN) in books",f"{len(dupb)}","Merged for matching - verify double booking" if len(dupb) else "OK"])
    if "gross" in sc["layout"][0]["mapped"] if sc["layout"] else False:
        g=d.copy(); g["calc"]=g.taxable+g.igst+g.cgst+g.sgst+g.cess+g.rigst+g.rcgst+g.rsgst
        chk=g[g._gross.notna()&(g._gross>0)]; fail=chk[(chk._gross-chk.calc.abs()).abs()>max(1,tol)]
        row_(["Gross total = value + tax (per voucher)",f"{len(chk)-len(fail)} of {len(chk)} pass","OK" if not len(fail) else "Check ledger classes: "+"; ".join(fail._inv.astype(str).head(6))])
    if res["filled"]:
        r[0]+=1; block("6. GSTIN inferred from party name - please verify"); row_(["Party","Inferred GSTIN","Basis"],True)
        for x in (res["rec"]["Invoice"]+res["rec"]["Credit/Debit Note"]):
            if x.get("gstin_src"): row_([x["party"],x["gstin"],x["gstin_src"]])
    if res["warns"]:
        r[0]+=1; block("Warnings")
        for w in res["warns"]: row_([w])
    widths(wm,[48,22,20,20,100,10])
    # ---------- Legend ----------
    wl=wb.create_sheet("Legend"); wl.sheet_view.showGridLines=False
    title(wl,"Status legend & ITC audit checklist",3)
    for i,s in enumerate(ITC_STAT):
        wl.cell(3+i,1,s).fill=F(ITC_BG[s]); wl.cell(3+i,1).font=DF(9,bold=True); wl.merge_cells(start_row=3+i,start_column=2,end_row=3+i,end_column=3); wl.cell(3+i,2,ITC_MEAN[s]).font=DF(9)
    b=4+len(ITC_STAT); wl.cell(b,1,"ITC AUDIT CHECKLIST").font=HF(10,NAVY)
    for i,t in enumerate(["Sec 16(2)(a) tax invoice / debit note held","16(2)(b) goods/services received (GRN, LR, e-way bill)","16(2)(aa) invoice reflected in GSTR-2B","16(2)(c) tax actually paid by supplier","16(2)(d) return filed",
        "16(4) time limit - 30 Nov of next FY or annual return date","17(5) blocked credits - motor vehicles, food, club, personal use, works contract on immovable property","Rule 37 - 180 days payment to supplier or reverse","Rule 42/43 - exempt / partly business use reversal","Credit notes - reduce ITC (Sec 34(2))","RCM - pay in cash first, claim 4(A)(3)","Imports - IGST via IMPG / Bill of Entry"]):
        wl.cell(b+1+i,1,"[ ]").font=DF(9); wl.merge_cells(start_row=b+1+i,start_column=2,end_row=b+1+i,end_column=3); wl.cell(b+1+i,2,t).font=DF(9)
    widths(wl,[28,50,50])
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

# ─────────────────────────────────────────────────────────────────────────────
#  TAB 3 (v2) — GSTR-3B CHECK  : 3B vs GSTR-1 vs Books (output)  ·  3B vs 2B vs ITC-reco (input)
# ─────────────────────────────────────────────────────────────────────────────
COLS3=["Taxable","IGST","CGST","SGST"]
ROWS_3B=["3.1(a) Outward taxable (other than zero-rated, nil, exempt)","3.1(b) Outward zero-rated (exports / SEZ)","3.1(c) Other outward nil-rated / exempt",
         "3.1(d) Inward supplies liable to reverse charge","4(A)(1) Import of goods","4(A)(3) Inward supplies liable to RCM (ITC)","4(A)(4) Inward supplies from ISD",
         "4(A)(5) All other ITC","4(B) ITC reversed (Rule 38/42/43 + others)","4(D) Ineligible ITC","4(C) Net ITC available (as per 3B)"]
ROWS_G1=["Total taxable outward supplies - net of credit/debit notes (GSTR-1 Total Liability)","Of which: exports / SEZ with payment (zero-rated)","Nil-rated / exempt / non-GST (GSTR-1 Table 8)"]
ROWS_BK=["Sales register - taxable outward (net of credit notes)"]
ROWS_2B=["2B: B2B invoices - ITC available (net of credit notes)","2B: Import of goods (IMPG)","2B: ISD credit","2B: Reverse-charge supplies (tax liability)","2B: ITC NOT available (info)"]
ROWS_IB=["ITC reco: eligible ITC on invoices (lower of books & 2B) net of CN","ITC reco: RCM entries in books"]

def zero_tbl(rows): return pd.DataFrame({"Line":rows,**{c:[0.0]*len(rows) for c in COLS3}})
def tget(df,label,col): 
    m=df[df.Line==label]; return float(m[col].iloc[0]) if len(m) and col in m else 0.0

def utilise(out,itc):
    I,C,S=out; i,c,s=itc; cash={}
    u=min(i,I); i-=u; I-=u
    u=min(i,C); i-=u; C-=u
    u=min(i,S); i-=u; S-=u
    u=min(c,C); c-=u; C-=u
    u=min(c,I); c-=u; I-=u
    u=min(s,S); s-=u; S-=u
    u=min(s,I); s-=u; I-=u
    return (I,C,S),(i,c,s)

def twob_refs(pt):
    """2B reference table from parse_2b_full output (ITC available only)"""
    ok=lambda r: not str(r.get("avail","")).strip().lower().startswith("n") and not r.get("rcm")
    t={h:0.0 for h in ("taxable","igst","cgst","sgst")}
    inv_=[r for r in pt["inv"] if ok(r)]; cn=[r for r in pt["notes"] if ok(r)]
    b2b={h:sum(r[h] for r in inv_)-sum(r[h] for r in cn if r["ntype"]!="D")+sum(r[h] for r in cn if r["ntype"]=="D") for h in t}
    rcm={h:sum(r[h] for r in pt["inv"]+pt["notes"] if r.get("rcm")) for h in t}
    na={h:sum(r[h] for r in pt["inv"]+pt["notes"] if str(r.get("avail","")).strip().lower().startswith("n") and not r.get("rcm")) for h in t}
    isd={h:sum(r[h] for r in pt["isd"]) for h in t}
    imp={"taxable":sum(r["Taxable"] for r in pt["imp"]),"igst":sum(r["IGST"] for r in pt["imp"]),"cgst":0.0,"sgst":0.0}
    df=zero_tbl(ROWS_2B)
    for lab,d in zip(ROWS_2B,(b2b,imp,isd,rcm,na)):
        for c in COLS3: df.loc[df.Line==lab,c]=d[c.lower()]
    return df

def books_sales_net(res):
    bk=res["books"]; inv=bk["inv"]; nt=bk["notes"]; out={c:float(inv[c].sum()) for c in COLS3}
    if len(nt):
        for c in COLS3:
            sgn=sum((1 if "debit" in str(v).lower() else -1)*abs(x) for v,x in zip(nt["_vtype"],nt[c]))
            out[c]+=sgn
    return out

def itc_reco_refs(res):
    inv=res["inv"]; notes=res["notes"]; pt=res["pt"]
    e=[float(sum(x[k] for x in inv.elig)) for k in (0,1,2)] if len(inv) else [0,0,0]
    cn=[0.0,0.0,0.0]
    for r in pt["notes"]:
        if str(r.get("avail","")).lower().startswith("n") or r.get("rcm"): continue
        s=-1 if r["ntype"]!="D" else 1
        for k,h in enumerate(HEADS): cn[k]+=s*r[h]
    rcm=inv[(inv.Status=="RCM - verify")&inv.has_b]
    return [e[k]+cn[k] for k in range(3)],[float(rcm.B_IGST.sum()),float(rcm.B_CGST.sum()),float(rcm.B_SGST.sum())]

def run_3b_v2(t3b,tg1,tbk,t2b,tib,tol=2.0):
    R=[]   # dict rows
    def add(sec,desc,v3,ref,refname,kind,note=""):
        d=v3-ref; ok=abs(d)<=tol
        if ok: st,rm="Match",""
        elif kind=="out":
            if d<0: st,rm="3B LESS than "+refname,f"Short payment Rs.{abs(d):,.0f} vs {refname} - pay with interest u/s 50 (18% p.a.) in next 3B; check amendments / DRC-03"
            else:   st,rm="3B MORE than "+refname,f"Excess Rs.{d:,.0f} - check advances, amendments / timing; refund or carry via next period if wrongly paid"
        elif kind=="itc":
            if d>0: st,rm="EXCESS CLAIM vs "+refname,f"ITC claimed Rs.{d:,.0f} above {refname} - reverse in 4(B)(2) with interest u/s 50(3) (24% p.a.)"
            else:   st,rm="Claimed LESS than "+refname,f"Rs.{abs(d):,.0f} available but not claimed - can be taken later if within Sec 16(4); check if invoices not booked / ineligible"
        else:
            st,rm=("3B higher" if d>0 else "3B lower"),f"Difference Rs.{abs(d):,.0f}"
        if ok and note: rm=note
        R.append({"Section":sec,"Description":desc,"3B":v3,"Reference":ref,"Ref source":refname,"Diff":d,"Status":st,"Remark":rm,"ok":ok})
    a=lambda lab,c: tget(t3b,lab,c)
    out_taxable=a(ROWS_3B[0],"Taxable")+a(ROWS_3B[1],"Taxable")
    # A. Output vs GSTR-1
    g_tot=ROWS_G1[0]
    add("A","Taxable value: 3.1(a)+3.1(b)",out_taxable,tget(tg1,g_tot,"Taxable"),"GSTR-1","out")
    for c in COLS3[1:]: add("A",f"{c} on outward: 3.1(a)+3.1(b)",a(ROWS_3B[0],c)+a(ROWS_3B[1],c),tget(tg1,g_tot,c),"GSTR-1","out")
    add("A","Zero-rated taxable: 3.1(b)",a(ROWS_3B[1],"Taxable"),tget(tg1,ROWS_G1[1],"Taxable"),"GSTR-1","out")
    add("A","Nil / exempt: 3.1(c)",a(ROWS_3B[2],"Taxable"),tget(tg1,ROWS_G1[2],"Taxable"),"GSTR-1","out")
    # A2 vs books
    if tbk is not None and tbk[COLS3].abs().sum().sum()>0:
        for c in COLS3: add("A","3B vs Books - "+(c if c!="Taxable" else "Taxable value")+" (3.1(a)+(b))",a(ROWS_3B[0],c)+a(ROWS_3B[1],c),tget(tbk,ROWS_BK[0],c),"Books","out")
        for c in COLS3: add("A","GSTR-1 vs Books - "+c,tget(tg1,g_tot,c),tget(tbk,ROWS_BK[0],c),"Books","out","")
    # B. RCM liability
    for c in COLS3[1:]: add("B",f"RCM liability 3.1(d) {c} vs 2B reverse-charge supplies",a(ROWS_3B[3],c),tget(t2b,ROWS_2B[3],c),"2B RCM","gen","")
    # C. ITC
    add("C","4(A)(1) Import IGST",a(ROWS_3B[4],"IGST"),tget(t2b,ROWS_2B[1],"IGST"),"2B IMPG","itc")
    for c in COLS3[1:]: add("C",f"4(A)(4) ISD {c}",a(ROWS_3B[6],c),tget(t2b,ROWS_2B[2],c),"2B ISD","itc")
    for c in COLS3[1:]: add("C",f"4(A)(5) All other ITC {c}",a(ROWS_3B[7],c),tget(t2b,ROWS_2B[0],c),"2B","itc")
    if tib is not None:
        for c in COLS3[1:]: add("C",f"4(A)(5) {c} vs ITC-reco eligible (books∩2B)",a(ROWS_3B[7],c),tget(tib,ROWS_IB[0],c),"ITC reco","itc")
    for c in COLS3[1:]: add("C",f"4(A)(3) RCM ITC {c} vs RCM tax paid 3.1(d)",a(ROWS_3B[5],c),a(ROWS_3B[3],c),"3.1(d) payment","itc","")
    # D. arithmetic of 4(C)
    for c in COLS3[1:]:
        calc=sum(a(ROWS_3B[k],c) for k in (4,5,6,7))-a(ROWS_3B[8],c)
        add("D",f"4(C) Net ITC {c} = 4(A) - 4(B)",a(ROWS_3B[10],c),calc,"3B arithmetic","gen","")
    df=pd.DataFrame(R)
    # utilisation
    out=[a(ROWS_3B[0],c)+a(ROWS_3B[1],c) for c in COLS3[1:]]; rcm=[a(ROWS_3B[3],c) for c in COLS3[1:]]
    itc=[a(ROWS_3B[10],c) for c in COLS3[1:]]
    cash,carry=utilise(out,itc)
    util={"out":out,"rcm":rcm,"itc":itc,"cash":cash,"carry":carry}
    return df,util

def build_3b_excel_v2(df,util,t3b,tg1,tbk,t2b,tib,client,period):
    wb=openpyxl.Workbook(); title,sub,head,widths=_xl_kit(wb)
    ws=wb.active; ws.title="3B Check"; ws.sheet_view.showGridLines=False
    H=["Sec","Description","3B (Filed)","Reference","Reference source","Diff (3B - Ref)","Status","Remark"]
    title(ws,f"{client}  -  GSTR-3B Audit Check  |  {period}",8); sub(ws,"A: output vs GSTR-1 / books · B: RCM · C: ITC vs 2B / ITC reco · D: arithmetic",8); head(ws,4,H)
    secname={"A":"A. OUTPUT LIABILITY","B":"B. REVERSE CHARGE","C":"C. ITC CLAIMED","D":"D. 3B ARITHMETIC"}; x=5; last=None
    for i,r in df.iterrows():
        if r.Section!=last:
            ws.merge_cells(start_row=x,start_column=1,end_row=x,end_column=8); c=ws.cell(x,1,secname[r.Section]); c.font=HF(10,NAVY); c.fill=F("E8EFF8"); x+=1; last=r.Section
        vals=[r.Section,r.Description,r["3B"],r.Reference,r["Ref source"],f"=C{x}-D{x}",r.Status,r.Remark]
        for ci,v in enumerate(vals,1):
            c=ws.cell(x,ci,v); c.border=BD(); c.font=DF(9,bold=(ci==7)); 
            if ci in (3,4,6): c.number_format=INR; c.alignment=RGT()
            if not r.ok and ci in (6,7): c.fill=F(MIS); 
            if ci==7: c.font=Font(name="Calibri",bold=True,size=9,color="1E5631" if r.ok else "C00000")
        x+=1
    x+=1; ws.merge_cells(start_row=x,start_column=1,end_row=x,end_column=8); c=ws.cell(x,1,"E. INDICATIVE TAX UTILISATION (Rule 88A order) - based on 3B output and 4(C) net ITC"); c.font=HF(10,NAVY); c.fill=F("E8EFF8"); x+=1
    head(ws,x,["","","IGST","CGST","SGST","Total","",""]); x+=1
    for lab,v in (("Output tax 3.1(a)+(b)",util["out"]),("RCM tax (cash only) 3.1(d)",util["rcm"]),("Net ITC 4(C)",util["itc"]),("Cash payable on outward (indicative)",util["cash"]),("ITC carried forward (indicative)",util["carry"])):
        ws.cell(x,2,lab).font=DF(9,bold=True)
        for k,val in enumerate(v): c=ws.cell(x,3+k,val); c.number_format=INR; c.border=BD()
        c=ws.cell(x,6,f"=SUM(C{x}:E{x})"); c.number_format=INR; c.border=BD(); x+=1
    widths(ws,[6,52,16,16,18,16,32,90]); ws.freeze_panes="A5"
    wi=wb.create_sheet("Inputs used"); wi.sheet_view.showGridLines=False; title(wi,f"{client}  -  Values used in this check  |  {period}",5); r=3
    for nm,t in (("GSTR-3B (filed) - typed",t3b),("GSTR-1 (filed) - typed",tg1),("Books sales (from Sales Reco / typed)",tbk),("GSTR-2B (from file)",t2b),("ITC Reco (from run / typed)",tib)):
        if t is None: continue
        wi.cell(r,1,nm).font=Font(name="Calibri",bold=True,size=11,color=NAVY); r+=1; head(wi,r,["Line"]+COLS3); r+=1
        for _,row in t.iterrows():
            wi.cell(r,1,row.Line).font=DF(9)
            for k,cn in enumerate(COLS3): c=wi.cell(r,2+k,float(row[cn])); c.number_format=INR; c.border=BD()
            r+=1
        r+=1
    widths(wi,[80,16,16,16,16])
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

# ─────────────────────────────────────────────────────────────────────────────
#  UI
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hdr">
  <h1>📊 GST Audit Reconciliation Tool</h1>
  <div class="sub">ITC Reco &nbsp;·&nbsp; Sales Reco &nbsp;·&nbsp; GSTR-3B Check &nbsp;·&nbsp; Works for any company</div>
  <div class="cred">Built by <a href="https://linkedin.com/in/purvadoshi26" target="_blank">Purva Doshi</a></div>
</div>""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### Settings")
    client=st.text_input("Client Name",placeholder="e.g. ABC Chemicals Pvt Ltd")
    period=st.text_input("Period",placeholder="e.g. January 2026")
    st.markdown("---")
    st.markdown("**Legend**")
    st.markdown("🟢 Matched · 🔴 Mismatch\n🟡 Only Portal/2B · 🔵 Only Books\n🟠 Custom/Import · ⚪ Exempt")
    st.markdown("---")
    st.caption("Built by [Purva Doshi](https://linkedin.com/in/purvadoshi26)")

if not client.strip(): client="Client"
if not period.strip(): period="Period"

tab1, tab2, tab3 = st.tabs(["🔵 ITC Reco (2B vs Purchase)", "🟢 Sales Reco (GSTR-1 vs Books)", "📋 3B Check"])

# ── TAB 1: ITC (invoice-wise) ─────────────────────────────────────────────────
with tab1:
    st.markdown("### GSTR-2B vs Purchase Register — invoice-wise ITC reco")
    c1,c2=st.columns(2)
    with c1:
        st.markdown("**File 1 — GSTR-2B Excel (portal)**")
        st.markdown('<div class="fmtbox">Portal → Returns → GSTR-2B → Download Excel<br>All sheets are read by <b>header name</b>: B2B · B2BA (amendments) · B2B-CDNR (credit/debit notes) · ISD · IMPG. Gets ITC-availability + reverse-charge flags per invoice.</div>',unsafe_allow_html=True)
        f_2b=st.file_uploader("GSTR-2B",type=["xlsx","xls"],key="itc_2b",label_visibility="collapsed")
    with c2:
        st.markdown("**File 2 — Purchase Register (Tally / SAP / Excel)**")
        st.markdown('<div class="fmtbox"><b>Best:</b> flat export — one row per voucher with <code>Voucher No</code> · <code>Supplier Invoice No</code> · <code>GSTIN/UIN</code> · <code>Date</code> + purchase / Input IGST / CGST / SGST ledger columns (any order, any extra columns). Debit-note vouchers in the same file are picked up.<br><b>Also accepted:</b> older Tally sub-row layout · pivot summary (party-level only)</div>',unsafe_allow_html=True)
        f_pr=st.file_uploader("Purchase Register",type=["xlsx","xls"],key="itc_pr",label_visibility="collapsed")
    if f_2b and f_pr:
        eb=f_2b.getvalue(); pbb=f_pr.getvalue(); sig=hashlib.md5(pbb+eb).hexdigest()[:10]; scan=None; pt=None
        try:
            if st.session_state.get("itc_sig")!=sig:
                st.session_state["itc_pt"]=parse_2b_full(eb); st.session_state["itc_scan"]=scan_pbooks_any(pbb)
                st.session_state["itc_sig"]=sig; st.session_state.pop("itc_out",None)
            scan=st.session_state["itc_scan"]; pt=st.session_state["itc_pt"]
        except Exception as e:
            st.error(f"❌ {e}")
            with st.expander("Details"): import traceback; st.code(traceback.format_exc())
        if scan is not None:
            pivot=(scan["kind"]=="legacy" and (scan["bw"]["_inv"].astype(str).str.strip()=="").all())
            for w in scan["warns"]+pt["warns"]: st.warning(w)
            if pivot:
                st.warning("Purchase file is a party-level pivot (no invoice numbers) → running the party-wise reco. Export the flat Purchase Register for invoice-wise matching.")
                if st.button("▶ Run ITC Reconciliation (party-wise)",type="primary",width="stretch",key="run_itc_old"):
                    try:
                        df,det,impg,impg_r=run_itc_reco(pbb,eb); st.session_state["itc_old"]=(df,build_itc_excel(df,det,impg_r,impg,client,period).getvalue(),impg)
                    except Exception as e: st.error(f"❌ {e}")
                if st.session_state.get("itc_old"):
                    df,xb,impg=st.session_state["itc_old"]; cnt=df["Status"].value_counts().to_dict()
                    st.success("✅ Party-wise ITC reco complete"); cols=st.columns(6)
                    for col,lbl,key in zip(cols,["✅ Matched","⚠️ Mismatch","🟡 Only 2B","🔵 Only Books","🔵 RCM","🟠 Custom"],["Matched","Mismatch","Only in 2B","Only in Books","RCM — Verify Books","Custom Duty / Import"]): col.metric(lbl,cnt.get(key,0))
                    st.download_button("⬇️ Download ITC Reco Excel",data=xb,file_name=f"{client.replace(' ','_')}_ITC_Reco_{period.replace(' ','_')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",width="stretch")
            else:
                bw=scan["bw"]; led=[c for c in bw.columns if not c.startswith("_")]
                with st.expander(f"🔧 Detected columns — review once ({len(led)} ledger columns, {len(bw)} voucher rows)",expanded=False):
                    st.caption("**Taxable** = purchase/expense value · **IGST/CGST/SGST** = input-tax ledgers · **RCM IGST/CGST/SGST** = reverse-charge tax ledgers · **Other** = round-off/TDS/TCS · **Ignore** = leave out.")
                    led_df=pd.DataFrame({"Ledger column":led,"Class":[scan["ledger_guess"][k] for k in led],"Total in file (₹)":[float(bw[k].sum()) for k in led]})
                    led_ed=st.data_editor(led_df,hide_index=True,width="stretch",disabled=["Ledger column","Total in file (₹)"],key=f"pled_{sig}",
                        column_config={"Class":st.column_config.SelectboxColumn("Class",options=LEDGER_CLASSES_P,required=True),"Total in file (₹)":st.column_config.NumberColumn(format="₹ %,.0f")})
                    vk=list(scan["vtype_guess"].keys())
                    vt_df=pd.DataFrame({"Voucher type":[v or "(blank)" for v in vk],"Treat as":[scan["vtype_guess"][k] for k in vk],"Vouchers":[int((bw._vtype==k).sum()) for k in vk]})
                    vt_ed=st.data_editor(vt_df,hide_index=True,width="stretch",disabled=["Voucher type","Vouchers"],key=f"pvt_{sig}",
                        column_config={"Treat as":st.column_config.SelectboxColumn("Treat as",options=["Invoice","Credit/Debit Note","Exclude"],required=True)})
                    cA,cB=st.columns(2)
                    tol=cA.number_input("Match tolerance (₹)",min_value=0.0,max_value=100.0,value=1.0,step=0.5,key="itc_tol")
                    mat=cB.number_input("High-priority threshold (₹ GST per invoice)",min_value=0.0,value=10000.0,step=1000.0,key="itc_mat")
                if st.button("▶ Run ITC Reconciliation",type="primary",width="stretch",key="run_itc"):
                    with st.spinner("Running invoice-wise ITC reco..."):
                        try:
                            lc=dict(zip(led_ed["Ledger column"],led_ed["Class"])); vc=dict(zip(vk,vt_ed["Treat as"]))
                            res=run_itc_v2(pbb,eb,lc,vc,tol,mat,scan=scan,pt=pt)
                            st.session_state["itc_out"]=(res,build_itc_excel_v2(res,client,period).getvalue())
                        except Exception as e:
                            st.session_state.pop("itc_out",None); st.error(f"❌ {e}")
                            with st.expander("Details"): import traceback; st.code(traceback.format_exc())
                if st.session_state.get("itc_out"):
                    res,xb=st.session_state["itc_out"]; inv=res["inv"]; cnt=inv.Status.value_counts().to_dict()
                    st.success("✅ ITC Reconciliation complete!"); cols=st.columns(6)
                    for col,lbl,keys in zip(cols,["✅ Matched","⚠️ Mismatch","🟡 Only 2B","🔵 Only Books","⛔ ITC N/A on 2B","🟢 RCM"],
                            [["Matched","Matched - Inv No differs","Matched - Rounding"],["Mismatch"],["Only in 2B"],["Only in Books"],["2B: ITC not available"],["RCM - verify"]]):
                        col.metric(lbl,sum(cnt.get(k,0) for k in keys))
                    el=sum(sum(x) for x in inv.elig) if len(inv) else 0
                    st.info(f"**Books ITC: ₹{inv.B_GST.sum():,.0f}** · 2B ITC: ₹{inv.P_GST.sum():,.0f} · **Eligible (lower of both): ₹{el:,.0f}** · Credit/debit notes: {len(res['notes'])} · IMPG: ₹{sum(r['IGST'] for r in res['pt']['imp']):,.0f}")
                    for w in res["warns"]: st.warning(w)
                    hi=inv[inv.Priority=="High"]
                    if len(hi):
                        with st.expander(f"🔴 High-priority items ({len(hi)})",expanded=True):
                            d=hi[["B_Inv","P_Inv","Party","B_GST","P_GST","Status","Remark"]].copy(); d.columns=["Books Inv","2B Inv","Party","Books GST (₹)","2B GST (₹)","Status","Remark"]
                            st.dataframe(d,width="stretch",height=260,hide_index=True,column_config={k:st.column_config.NumberColumn(format="₹ %,.0f") for k in ["Books GST (₹)","2B GST (₹)"]})
                    with st.expander("View invoice-wise reco"):
                        d=inv[["B_Inv","P_Inv","Party","GSTIN","B_GST","P_GST","Status","Remark"]].copy(); d["Diff (₹)"]=d["B_GST"]-d["P_GST"]
                        d.columns=["Books Inv","2B Inv","Party","GSTIN","Books GST (₹)","2B GST (₹)","Status","Remark","Diff (₹)"]
                        st.dataframe(d,width="stretch",height=380,hide_index=True,column_config={k:st.column_config.NumberColumn(format="₹ %,.0f") for k in ["Books GST (₹)","2B GST (₹)","Diff (₹)"]})
                    if len(res["notes"]):
                        with st.expander(f"Credit / debit notes ({len(res['notes'])})"):
                            d=res["notes"][["B_Inv","P_Inv","Party","B_GST","P_GST","Status","Remark"]].copy(); d.columns=["Books Note","2B Note","Party","Books GST (₹)","2B GST (₹)","Status","Remark"]
                            st.dataframe(d,width="stretch",hide_index=True)
                    st.download_button("⬇️ Download ITC Reco Excel",data=xb,file_name=f"{client.replace(' ','_')}_ITC_Reco_{period.replace(' ','_')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",width="stretch")
                    st.caption("Excel: Summary · Auditor Actions · Invoice Reco · Credit-Debit Notes · Party Summary · Exceptions · IMPG & ISD · Mapping & Checks · Legend")
    else:
        st.info("Upload both files above to run.")

# ── TAB 2: SALES ──────────────────────────────────────────────────────────────
with tab2:
    st.markdown("### E-Invoice / GSTR-1 vs Books (Sales Register)")
    c1,c2=st.columns(2)
    with c1:
        st.markdown("**File 1 — E-Invoice / GSTR-1 Portal Excel**")
        st.markdown('<div class="fmtbox">GST Portal → E-Invoice → Download Excel<br>Sheets (invoices, credit/debit notes, HSN) are found by <b>header names</b> — sheet names and column order can be anything.<br>Needs headers: <code>Invoice number</code> · <code>Taxable Value</code> · <code>GSTIN</code> · tax columns</div>',unsafe_allow_html=True)
        f_einv=st.file_uploader("E-Invoice Excel",type=["xlsx","xls"],key="sal_inv",label_visibility="collapsed")
    with c2:
        st.markdown("**File 2 — Sales Register (Tally / SAP / Excel)**")
        st.markdown('<div class="fmtbox">Any company format — columns are found by <b>header name</b>, in any order.<br>Needs headers: <code>Voucher No / Invoice No</code> · <code>Date</code> · <code>GSTIN</code> · <code>Gross Total</code> (optional) + your sales / IGST / CGST / SGST ledger columns. Ledger columns are auto-classified — review below.</div>',unsafe_allow_html=True)
        f_sal=st.file_uploader("Sales Register",type=["xlsx","xls"],key="sal_bk",label_visibility="collapsed")
    if f_einv and f_sal:
        eb=f_einv.getvalue(); bb=f_sal.getvalue(); sig=hashlib.md5(bb).hexdigest()[:10]
        scan=None
        try:
            if st.session_state.get("sal_sig")!=sig:
                st.session_state["sal_scan"]=scan_books(bb); st.session_state["sal_sig"]=sig; st.session_state.pop("sal_out",None)
            scan=st.session_state["sal_scan"]
        except Exception as e:
            st.error(f"❌ {e}")
            with st.expander("Details"): import traceback; st.code(traceback.format_exc())
        if scan:
            for w in scan["warns"]: st.warning(w)
            bw=scan["bw"]
            with st.expander(f"🔧 Detected columns — review once ({len(scan['ledger_guess'])} ledger columns, {len(bw)} voucher rows)",expanded=False):
                st.caption("Auto-detected from header names. If a ledger is wrongly classified, change it here before running. **Taxable** = sales/charges value · **IGST/CGST/SGST** = tax ledgers · **Other** = round-off/TCS · **Ignore** = leave out.")
                lkeys=list(scan["ledger_guess"].keys())
                led_df=pd.DataFrame({"Ledger column":lkeys,"Class":[scan["ledger_guess"][k] for k in lkeys],"Total in file (₹)":[float(bw[k].sum()) for k in lkeys]})
                led_ed=st.data_editor(led_df,hide_index=True,width="stretch",disabled=["Ledger column","Total in file (₹)"],key=f"led_{sig}",
                    column_config={"Class":st.column_config.SelectboxColumn("Class",options=LEDGER_CLASSES,required=True),
                                   "Total in file (₹)":st.column_config.NumberColumn(format="₹ %,.0f")})
                vkeys=list(scan["vtype_guess"].keys())
                vt_df=pd.DataFrame({"Voucher type":[v or "(blank)" for v in vkeys],"Treat as":[scan["vtype_guess"][k] for k in vkeys],"Vouchers":[int((bw._vtype==k).sum()) for k in vkeys]})
                vt_ed=st.data_editor(vt_df,hide_index=True,width="stretch",disabled=["Voucher type","Vouchers"],key=f"vt_{sig}",
                    column_config={"Treat as":st.column_config.SelectboxColumn("Treat as",options=["Invoice","Credit/Debit Note","Exclude"],required=True)})
                tol=st.number_input("Match tolerance (₹)",min_value=0.0,max_value=100.0,value=1.0,step=0.5,key="sal_tol")
            if st.button("▶ Run Sales Reconciliation",type="primary",width="stretch",key="run_sal"):
                with st.spinner("Running sales reco..."):
                    try:
                        lc=dict(zip(led_ed["Ledger column"],led_ed["Class"])); vc=dict(zip(vkeys,vt_ed["Treat as"]))
                        res=run_sales_reco(eb,bb,lc,vc,tol,scan=scan)
                        buf=build_sales_excel(res,client,period)
                        st.session_state["sal_out"]=(res,buf.getvalue())
                    except Exception as e:
                        st.session_state.pop("sal_out",None); st.error(f"❌ {e}")
                        with st.expander("Details"): import traceback; st.code(traceback.format_exc())
            if st.session_state.get("sal_out"):
                res,xbytes=st.session_state["sal_out"]; reco=res["reco"]; cnt=reco["Status"].value_counts().to_dict(); bk=res["books"]
                st.success("✅ Sales Reconciliation complete!")
                cols=st.columns(6)
                for col,lbl,keys in zip(cols,["✅ Matched","⚠️ Mismatch","🟡 Only Portal","🔵 Only Books","🔴 Cancelled but live","⚪ Cancelled / B2C"],
                        [["Matched","Matched - Inv No differs"],["Mismatch"],["Only in Portal"],["Only in Books"],["Cancelled in books - LIVE on portal","LIVE in books - cancelled on portal"],["Cancelled in books","No GSTIN (B2C / unregistered)"]]):
                    col.metric(lbl,sum(cnt.get(k,0) for k in keys))
                st.info(f"**Books taxable: ₹{reco['B_Tax'].sum():,.0f}** · Portal taxable: ₹{reco['P_Tax'].sum():,.0f} · Difference: ₹{reco['B_Tax'].sum()-reco['P_Tax'].sum():,.0f}")
                if len(bk["gross_fail"]) or len(bk["rate_fail"]):
                    st.warning(f"Books checks: {len(bk['gross_fail'])} voucher(s) where Gross ≠ value + tax, {len(bk['rate_fail'])} with unusual GST rate — a ledger may be mis-classified. Review 'Detected columns'.")
                for w in res["warns"]: st.warning(w)
                hi=reco[reco.Priority=="High"]
                if len(hi):
                    with st.expander(f"🔴 High-priority items ({len(hi)})",expanded=True):
                        d=hi[["B_Inv","P_Inv","Party","B_Tax","P_Tax","Status","Remark"]].copy(); d.columns=["Books Inv","Portal Inv","Party","Books (₹)","Portal (₹)","Status","Remark"]
                        st.dataframe(d,width="stretch",height=260,hide_index=True,column_config={"Books (₹)":st.column_config.NumberColumn(format="₹ %,.0f"),"Portal (₹)":st.column_config.NumberColumn(format="₹ %,.0f")})
                with st.expander("View invoice-wise reco"):
                    d=reco[["B_Inv","P_Inv","Party","B_Tax","P_Tax","Status","Remark"]].copy(); d["Diff (₹)"]=d["B_Tax"]-d["P_Tax"]
                    d.columns=["Books Inv","Portal Inv","Party","Books (₹)","Portal (₹)","Status","Remark","Diff (₹)"]
                    st.dataframe(d,width="stretch",height=380,hide_index=True,column_config={k:st.column_config.NumberColumn(format="₹ %,.0f") for k in ["Books (₹)","Portal (₹)","Diff (₹)"]})
                st.download_button("⬇️ Download Sales Reco Excel",data=xbytes,file_name=f"{client.replace(' ','_')}_Sales_Reco_{period.replace(' ','_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",width="stretch")
                st.caption("Excel: Summary · Invoice Reco · Credit-Debit Notes · Exceptions · Mapping & Checks")
    else:
        st.info("Upload both files above to run.")

# ── TAB 3: 3B CHECK (v2) ──────────────────────────────────────────────────────
with tab3:
    st.markdown("### GSTR-3B Audit Check — 3B vs GSTR-1 vs Books · 3B vs 2B vs ITC reco")
    st.markdown('<div class="fmtbox">Type the filed <b>GSTR-3B</b> and <b>GSTR-1</b> figures (from the portal PDF/summary). <b>2B</b>, <b>books sales</b> and <b>ITC-reco</b> references fill automatically when you have uploaded the 2B here or run the Sales / ITC tabs in this session. Every cell stays editable.</div>',unsafe_allow_html=True)
    st.markdown("**Step 1 — GSTR-2B Excel** *(optional if the ITC tab already has it)*")
    f_2b3=st.file_uploader("GSTR-2B Excel",type=["xlsx","xls"],key="tb3_2b",label_visibility="collapsed")
    pt3=None; src2b=""
    if f_2b3:
        try: pt3=parse_2b_full(f_2b3.getvalue()); src2b="uploaded 2B"
        except Exception as e: st.warning(f"Could not read 2B: {e}")
    elif st.session_state.get("itc_pt") is not None: pt3=st.session_state["itc_pt"]; src2b="2B from ITC tab"
    t2b_def=twob_refs(pt3) if pt3 else zero_tbl(ROWS_2B)
    sal=st.session_state.get("sal_out"); itcx=st.session_state.get("itc_out")
    tbk_def=zero_tbl(ROWS_BK); tib_def=zero_tbl(ROWS_IB); srcs=[]
    if sal:
        bs=books_sales_net(sal[0]); 
        for c in COLS3: tbk_def.loc[0,c]=bs[c]
        srcs.append("books sales ← Sales Reco tab")
    if itcx:
        e,rc=itc_reco_refs(itcx[0])
        for k,c in enumerate(COLS3[1:]): tib_def.loc[0,c]=e[k]; tib_def.loc[1,c]=rc[k]
        srcs.append("ITC reference ← ITC Reco tab")
    if src2b: srcs.append(src2b)
    if srcs: st.caption("Auto-filled: "+" · ".join(srcs))
    ssig=hashlib.md5((str(t2b_def.values.tolist())+str(tbk_def.values.tolist())+str(tib_def.values.tolist())).encode()).hexdigest()[:8]
    cfg={c:st.column_config.NumberColumn(c,format="%.2f",step=1.0) for c in COLS3}
    st.markdown("---"); st.markdown("**Step 2 — GSTR-3B (filed) — type from the return**")
    e3b=st.data_editor(zero_tbl(ROWS_3B),hide_index=True,width="stretch",disabled=["Line"],column_config=cfg,key="e3b")
    st.markdown("**Step 3 — GSTR-1 (filed)** — use *Total Liability* line of the GSTR-1 summary (net of credit/debit notes & amendments)")
    eg1=st.data_editor(zero_tbl(ROWS_G1),hide_index=True,width="stretch",disabled=["Line"],column_config=cfg,key="eg1")
    with st.expander("References (auto-filled — edit if needed)",expanded=bool(srcs)):
        ebk=st.data_editor(tbk_def,hide_index=True,width="stretch",disabled=["Line"],column_config=cfg,key=f"ebk_{ssig}")
        e2b=st.data_editor(t2b_def,hide_index=True,width="stretch",disabled=["Line"],column_config=cfg,key=f"e2b_{ssig}")
        eib=st.data_editor(tib_def,hide_index=True,width="stretch",disabled=["Line"],column_config=cfg,key=f"eib_{ssig}")
    tol3=st.number_input("Tolerance (₹)",min_value=0.0,value=2.0,step=1.0,key="tb3_tol")
    if st.button("▶ Run 3B Check",type="primary",width="stretch",key="run_3b"):
        try:
            use_bk=ebk if ebk[COLS3].abs().sum().sum()>0 else None; use_ib=eib if eib[COLS3].abs().sum().sum()>0 else None
            df3,util=run_3b_v2(e3b,eg1,use_bk,e2b,use_ib,tol3)
            st.session_state["tb3_out"]=(df3,util,build_3b_excel_v2(df3,util,e3b,eg1,use_bk,e2b,use_ib,client,period).getvalue())
        except Exception as e:
            st.session_state.pop("tb3_out",None); st.error(f"❌ {e}")
            with st.expander("Details"): import traceback; st.code(traceback.format_exc())
    if st.session_state.get("tb3_out"):
        df3,util,xb=st.session_state["tb3_out"]; bad=df3[~df3.ok]
        if len(bad): st.warning(f"⚠️ {len(bad)} difference(s) found out of {len(df3)} checks")
        else: st.success("✅ 3B agrees with all references")
        d=df3[["Section","Description","3B","Reference","Ref source","Diff","Status","Remark"]]
        st.dataframe(d,width="stretch",hide_index=True,height=420,column_config={k:st.column_config.NumberColumn(format="₹ %,.2f") for k in ["3B","Reference","Diff"]})
        st.caption(f"Indicative cash payable on outward (IGST/CGST/SGST): ₹{util['cash'][0]:,.0f} / ₹{util['cash'][1]:,.0f} / ₹{util['cash'][2]:,.0f} · RCM tax payable in cash: ₹{sum(util['rcm']):,.0f}")
        st.download_button("⬇️ Download 3B Check Excel",data=xb,file_name=f"{client.replace(' ','_')}_3B_Check_{period.replace(' ','_')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",width="stretch")
    with st.expander("📋 Where to find the values"):
        st.markdown("""
**GSTR-3B** — Returns → GSTR-3B → View filed return (PDF): Table 3.1 (a)(b)(c)(d), Table 4 (A)(B)(C)(D).
**GSTR-1** — Returns → GSTR-1 → View summary: *Total Liability* row, Table 8 nil/exempt, exports rows.
**2B** — upload the Excel above (or run the ITC tab first). **Books sales** — run the Sales tab first. **ITC reference** — run the ITC tab first.
        """)

st.markdown("---")
st.markdown('<div style="text-align:center;font-size:.74rem;color:#888">GST Audit Reconciliation Tool · Built by <a href="https://linkedin.com/in/purvadoshi26" style="color:#5585B5">Purva Doshi</a></div>',unsafe_allow_html=True)
