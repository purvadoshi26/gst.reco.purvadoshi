"""
GST Audit Reconciliation Tool
Created by: Purva Doshi | linkedin.com/in/purvadoshi26
Tabs: ITC Reco | Sales Reco | 3B vs GSTR-1 Check
"""
import streamlit as st
import io, re, datetime
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
#  TAB 2 — SALES RECO (GSTR-1 / E-Invoice vs Books)
# ─────────────────────────────────────────────────────────────────────────────
TALLY_SALES_SH=["GST Local","GST SALES (Local)","IGST SALES (INTEGRATED)","IGST (Sales)",
                "Handling Charges","Handling charges","Storage Charges ","STORAGE CHARGES",
                "SALES TO SEZ (IGST)","STORAGE CHARGES - SEZ UNIT","HANDLING CHARGES - SEZ UNIT",
                "Sales Exempt","SALES-EXEMPT","Detention Chgs","SHORTAGE IN TRANSIT",
                "Freight Charges (Income)","FREIGHT CHARGES(SALES)","Storage Income",
                "Sale Exempt-CN","Shortage in transit-CN"]
EXEMPT_SH={"Sales Exempt","SALES-EXEMPT","Detention Chgs","SHORTAGE IN TRANSIT","Sale Exempt-CN","Shortage in transit-CN"}

def parse_sales_books(fb):
    wb=load_wb(fb); rows=[]
    found_tally=any(s in wb.sheetnames for s in TALLY_SALES_SH)
    sheets=[s for s in TALLY_SALES_SH if s in wb.sheetnames] if found_tally else wb.sheetnames
    for sname in sheets:
        ws=wb[sname]; hr=None
        for i,row in enumerate(ws.iter_rows(min_row=1,max_row=20,values_only=True),1):
            rl=[str(v or "").lower() for v in row]
            if any(k in " ".join(rl) for k in ["voucher no","invoice no","bill no"]):
                hr=i; break
        if not hr: continue
        hdrs=[ws.cell(hr,c).value for c in range(1,ws.max_column+1)]; hl=[str(h or "").lower() for h in hdrs]
        def fc2(keys):
            for k in keys:
                for i,h in enumerate(hl):
                    if k==h or (len(k)>4 and k in h): return i
            return None
        inv_ci=fc2(["voucher no.","voucher no","invoice no","bill no","doc no"])
        pty_ci=fc2(["particulars","party","customer","party name"])
        dat_ci=fc2(["date","invoice date","bill date","voucher date"])
        gst_ci=fc2(["gstin/uin","gstin"])
        # Smart value column: prefer specific charge col over generic "value"
        val_ci=None
        for i2,h in enumerate(hl):
            if i2>5 and any(kw in h for kw in ["gst sales","igst sales","storage charges","handling charges","sales to sez","shortage","detention","freight charges","storage income"]): val_ci=i2; break
        if val_ci is None: val_ci=fc2(["value","taxable value","taxable","basic","amount"])
        if inv_ci is None: continue
        for row in ws.iter_rows(min_row=hr+1,max_row=ws.max_row,values_only=True):
            inv=str(row[inv_ci] or "").strip() if inv_ci<len(row) else ""
            if not inv: continue
            p=str(row[pty_ci] or "").strip() if pty_ci is not None and pty_ci<len(row) else ""
            if p.lower() in ("total","grand total","(cancelled )","cancelled"): continue
            try: val=float(row[val_ci] or 0) if val_ci is not None and val_ci<len(row) else 0
            except: val=0
            dt=row[dat_ci] if dat_ci is not None and dat_ci<len(row) else (row[0] if row else None)
            gstin=cgstin(row[gst_ci]) if gst_ci is not None and gst_ci<len(row) else ""
            rows.append({"Date":dt,"Party":p,"GSTIN":gstin,"InvNo":inv,"Taxable":val,"Category":sname.strip()})
    if not rows: raise ValueError("No data in sales/books file.\nNeeds: Invoice No / Voucher No + Value columns.")
    df=pd.DataFrame(rows); df["inv_norm"]=df["InvNo"].apply(norm_inv)
    return df.groupby("inv_norm").agg(Date=("Date","first"),Party=("Party","first"),GSTIN=("GSTIN","first"),
        InvNo=("InvNo","first"),Taxable=("Taxable","sum"),Category=("Category","first")).reset_index()

def parse_sales_portal(fb):
    wb=load_wb(fb)
    # Try E-Invoice format (b2b, sez, de sheet, header row 4, data row 5)
    for sn in ["b2b, sez, de","b2b,sez,de","B2B, SEZ, DE"]:
        if sn in wb.sheetnames:
            ws=wb[sn]; rows=[]
            # Detect header: row 4 should have GSTIN column
            hr=4
            for i,row in enumerate(ws.iter_rows(min_row=1,max_row=7,values_only=True),1):
                if row[0] and "gstin" in str(row[0]).lower(): hr=i; break
            for row in ws.iter_rows(min_row=hr+1,max_row=ws.max_row,values_only=True):
                if not row[0] or not str(row[0]).strip(): continue
                # E-Invoice format: GSTIN|Receiver|InvNo|Date|InvVal|PoS|RC|Rate|InvType|ECO|Rate2|Taxable|IGST|CGST|SGST|Cess|...
                rows.append({"GSTIN":cgstin(row[0]),"Party":str(row[1] or "").strip(),
                             "InvNo":str(row[2] or "").strip(),"Date":str(row[3] or ""),
                             "Taxable":safe(str(row[11] or "0").replace(",","")),
                             "IGST":safe(str(row[12] or "0").replace(",","")),
                             "CGST":safe(str(row[13] or "0").replace(",","")),
                             "SGST":safe(str(row[14] or "0").replace(",",""))})
            if rows:
                df=pd.DataFrame(rows); df["inv_norm"]=df["InvNo"].apply(norm_inv); return df
    raise ValueError("Portal file needs sheet 'b2b, sez, de'.\nDownload E-Invoice Excel from GST Portal → E-Invoice → Download.")

def run_sales_reco(portal_bytes,books_bytes):
    df_b=parse_sales_books(books_bytes); df_p=parse_sales_portal(portal_bytes)
    reco=pd.merge(df_b,df_p,on="inv_norm",how="outer",suffixes=("_b","_p"))
    reco["B_Tax"]=reco["Taxable_b"].fillna(0); reco["P_Tax"]=reco["Taxable_p"].fillna(0); reco["Diff"]=reco["B_Tax"]-reco["P_Tax"]
    def status(r):
        cat=str(r.get("Category",""))
        if cat in EXEMPT_SH and r["P_Tax"]==0: return "Exempt / Bond Transfer"
        if r["B_Tax"]==0 and r["P_Tax"]==0: return "Skip"
        if r["B_Tax"]==0: return "Only in Portal"
        if r["P_Tax"]==0: return "Only in Books"
        if abs(r["Diff"])<1: return "Matched"
        return "Value Mismatch"
    reco["Status"]=reco.apply(status,axis=1)
    return reco[reco["Status"]!="Skip"].sort_values(["Status","inv_norm"]).reset_index(drop=True),df_b,df_p

def build_sales_excel(reco,df_b,df_p,client,period):
    wb=openpyxl.Workbook(); cnt=reco["Status"].value_counts().to_dict()
    b_tot=df_b["Taxable"].sum(); p_tot=df_p["Taxable"].sum()
    SAL_BG={"Matched":MATCH,"Value Mismatch":MIS,"Only in Portal":ONLY2,"Only in Books":ONLY1,"Exempt / Bond Transfer":EXEMPT}
    SAL_FC={"Matched":"1E5631","Value Mismatch":"C00000","Only in Portal":"7A5C00","Only in Books":"0D3055","Exempt / Bond Transfer":"555555"}
    SAL_RM={"Matched":"Books = Portal ✓","Value Mismatch":"Same invoice, taxable differs — verify",
            "Only in Portal":"Filed in GSTR-1 but not found in books — check other ledgers",
            "Only in Books":"In books but not in portal — verify if filed or exempt","Exempt / Bond Transfer":"Exempt / Bond Transfer — not e-invoiced (correct)"}

    def mhdr(ws,txt,nc,r=1):
        ws.merge_cells(f"A{r}:{get_column_letter(nc)}{r}"); c=ws.cell(r,1); c.value=txt
        c.font=Font(name="Calibri",bold=True,size=12,color="FFFFFF"); c.fill=F(TITLE); c.alignment=CTR(); ws.row_dimensions[r].height=28
    def sub(ws,txt,nc,r=2):
        ws.merge_cells(f"A{r}:{get_column_letter(nc)}{r}"); c=ws.cell(r,1)
        c.value=f"Created by: Purva Doshi  |  linkedin.com/in/purvadoshi26  |  {txt}"
        c.font=Font(name="Calibri",size=8,italic=True,color="7F9FC0"); c.alignment=CTR(); ws.row_dimensions[r].height=14

    ws1=wb.active; ws1.title="Summary"; ws1.sheet_view.showGridLines=False
    mhdr(ws1,f"{client}  —  GSTR-1 vs Books Sales Reco  |  {period}",5); sub(ws1,f"Sales Reco | {period}",5)
    for ci,v in enumerate(["","Books","Portal","Diff","Remark"],1):
        c=ws1.cell(4,ci); c.value=v; c.font=Font(name="Calibri",bold=True,size=9,color=NAVY)
        c.fill=F("E8EFF8"); c.alignment=CTR(); c.border=BD()
    for ri,rd in enumerate([("Total Invoices",len(df_b),len(df_p),len(df_b)-len(df_p),""),
        ("✅ Matched",cnt.get("Matched",0),cnt.get("Matched",0),0,"Clean"),
        ("⚠️ Value Mismatch",cnt.get("Value Mismatch",0),cnt.get("Value Mismatch",0),0,"Investigate"),
        ("🔵 Only in Books",cnt.get("Only in Books",0),0,cnt.get("Only in Books",0),"Not filed"),
        ("🟡 Only in Portal",0,cnt.get("Only in Portal",0),-cnt.get("Only in Portal",0),"Check books"),
        ("⚪ Exempt/Bond",cnt.get("Exempt / Bond Transfer",0),0,"N/A","Not e-invoiced (correct)")],5):
        for ci,val in enumerate(rd,1):
            c=ws1.cell(ri,ci); c.value=val; c.border=BD(); c.font=DF(9,bold=(ci==1))
            if ci in(2,3,4) and isinstance(val,(int,float)): c.number_format="#,##0"; c.alignment=RGT()
            else: c.alignment=LFT()
        ws1.row_dimensions[ri].height=16
    for ci,v in enumerate(["","Books (₹)","Portal (₹)","Diff (₹)",""],1):
        c=ws1.cell(12,ci); c.value=v; c.font=HF(9); c.fill=F(NAVY); c.alignment=CTR(); c.border=BD()
    for ri,rd in enumerate([("Total Taxable",b_tot,p_tot,b_tot-p_tot),
        ("Matched",reco[reco.Status=="Matched"]["B_Tax"].sum(),reco[reco.Status=="Matched"]["P_Tax"].sum(),0),
        ("Exempt/Bond",reco[reco.Status=="Exempt / Bond Transfer"]["B_Tax"].sum(),0,reco[reco.Status=="Exempt / Bond Transfer"]["B_Tax"].sum()),
        ("Only in Books",reco[reco.Status=="Only in Books"]["B_Tax"].sum(),0,reco[reco.Status=="Only in Books"]["B_Tax"].sum()),
        ("Only in Portal",0,reco[reco.Status=="Only in Portal"]["P_Tax"].sum(),-reco[reco.Status=="Only in Portal"]["P_Tax"].sum())],13):
        for ci,val in enumerate(rd,1):
            c=ws1.cell(ri,ci); c.value=val; c.border=BD(); c.font=DF(9,bold=(ci==1))
            if ci in(2,3,4) and isinstance(val,(int,float)): c.number_format=INR; c.alignment=RGT()
            else: c.alignment=LFT()
    for col,w in zip("ABCDE",[32,18,18,18,28]): ws1.column_dimensions[col].width=w

    ws2=wb.create_sheet("Invoice Reco"); ws2.sheet_view.showGridLines=False
    mhdr(ws2,f"{client}  —  GSTR-1 vs Books Invoice Reco  |  {period}",11); sub(ws2,f"Sales Reco | {period}",11)
    rc2=[("#",4),("Invoice No (Books)",22),("Invoice No (Portal)",22),("Date",12),("Party",38),
         ("Category",20),("Books (₹)",15),("Portal (₹)",15),("Diff (₹)",14),("Status",24),("Remarks",44)]
    for ci,(hdr,w) in enumerate(rc2,1):
        c=ws2.cell(4,ci); c.value=hdr; c.font=HF(8); c.fill=F(NAVY); c.alignment=CTR(); c.border=BD()
        ws2.column_dimensions[get_column_letter(ci)].width=w
    ws2.row_dimensions[4].height=20
    for ri,row in reco.iterrows():
        r=ri+5; st=row["Status"]; bg=SAL_BG.get(st,"FFFFFF")
        b_inv=str(row.get("InvNo_b") or ""); p_inv=str(row.get("InvNo") or row.get("InvNo_p") or "")
        party=str(row.get("Party_b") or row.get("Party_p") or row.get("Party") or "")
        date=row.get("Date_b") or row.get("Date_p") or row.get("Date") or ""
        vals=[ri+1,b_inv,p_inv,date,party,str(row.get("Category","")),row["B_Tax"],row["P_Tax"],row["Diff"],st,SAL_RM.get(st,"")]
        for ci,val in enumerate(vals,1):
            c=ws2.cell(r,ci); c.value=val; c.fill=F(bg); c.border=BD()
            if ci in(7,8,9): c.font=DF(9); c.alignment=RGT(); c.number_format=INR
            elif ci==4: c.font=DF(9); c.number_format="DD-MMM-YY"; c.alignment=CTR()
            elif ci==1: c.font=DF(9); c.alignment=CTR()
            elif ci==10: c.font=Font(name="Calibri",bold=True,size=9,color=SAL_FC.get(st,"333")); c.alignment=CTR()
            else: c.font=DF(9); c.alignment=LFT()
        ws2.row_dimensions[r].height=14
    ws2.freeze_panes="B5"

    def det2(df_in,cols,title,hc):
        ws=wb.create_sheet(title[:31]); ws.sheet_view.showGridLines=False
        ws.merge_cells(f"A1:{get_column_letter(len(cols))}1"); h=ws.cell(1,1); h.value=title
        h.font=Font(name="Calibri",bold=True,size=11,color="FFFFFF"); h.fill=F(hc); h.alignment=CTR()
        for ci,(hdr,w,fld) in enumerate(cols,1):
            c=ws.cell(2,ci); c.value=hdr; c.font=HF(9); c.fill=F(hc); c.alignment=CTR(); c.border=BD()
            ws.column_dimensions[get_column_letter(ci)].width=w
        for ri,row in df_in.iterrows():
            r2=ri+3
            for ci,(_,_,fld) in enumerate(cols,1):
                c=ws.cell(r2,ci); c.value=row.get(fld,""); c.border=BD(); c.font=DF(9)
                if fld=="Date": c.number_format="DD-MMM-YY"; c.alignment=CTR()
                elif fld in("Taxable","Taxable Value","IGST","CGST","SGST"): c.number_format=INR; c.alignment=RGT()
                else: c.alignment=LFT()
            ws.row_dimensions[r2].height=14
        ws.freeze_panes="A3"

    df_b2=df_b.rename(columns={"InvNo":"Invoice No","Taxable":"Taxable Value"})
    df_p2=df_p.rename(columns={"InvNo":"Invoice No","Taxable":"Taxable Value"})
    det2(df_b2,[("Date",12,"Date"),("Party",40,"Party"),("Invoice No",22,"Invoice No"),("GSTIN",22,"GSTIN"),("Category",22,"Category"),("Taxable Value",16,"Taxable Value")],f"Books Detail | {period}",BLUE)
    det2(df_p2,[("GSTIN",22,"GSTIN"),("Party",40,"Party"),("Invoice No",22,"Invoice No"),("Date",12,"Date"),("Taxable Value",16,"Taxable Value"),("IGST",14,"IGST"),("CGST",14,"CGST"),("SGST",14,"SGST")],f"Portal Detail | {period}",GREEN)
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

# ── TAB 1: ITC ────────────────────────────────────────────────────────────────
with tab1:
    st.markdown("### GSTR-2B vs Tally Purchase Register")
    c1,c2=st.columns(2)
    with c1:
        st.markdown("**File 1 — GSTR-2B Excel**")
        st.markdown('<div class="fmtbox">Portal → Returns → GSTR-2B → Download Excel<br>Must have sheet: <code>B2B</code> · Old and new format supported · IMPG auto-detected</div>',unsafe_allow_html=True)
        f_2b=st.file_uploader("GSTR-2B",type=["xlsx","xls"],key="itc_2b",label_visibility="collapsed")
    with c2:
        st.markdown("**File 2 — Tally Purchase Register**")
        st.markdown('<div class="fmtbox"><b>Preferred:</b> Detailed Register → Tally → Purchase Register → Alt+F1 → Alt+E → Excel<br><b>Also accepted:</b> Pivot summary (Party | Sum of CGST/SGST/IGST) · Sub-ledger format · <code>.xls</code> files</div>',unsafe_allow_html=True)
        f_pr=st.file_uploader("Purchase Register",type=["xlsx","xls"],key="itc_pr",label_visibility="collapsed")
    if f_2b and f_pr:
        if st.button("▶ Run ITC Reconciliation",type="primary",use_container_width=True,key="run_itc"):
            with st.spinner("Running ITC reco..."):
                try:
                    df,det,impg,impg_r=run_itc_reco(f_pr.read(),f_2b.read())
                    cnt=df["Status"].value_counts().to_dict()
                    buf=build_itc_excel(df,det,impg_r,impg,client,period)
                    st.success("✅ ITC Reconciliation complete!")
                    cols=st.columns(6)
                    for col,lbl,key in zip(cols,["✅ Matched","⚠️ Mismatch","🟡 Only 2B","🔵 Only Books","🔵 RCM","🟠 Custom"],
                                           ["Matched","Mismatch","Only in 2B","Only in Books","RCM — Verify Books","Custom Duty / Import"]):
                        col.metric(lbl,cnt.get(key,0))
                    claimable=df[df.Status=="Matched"]["G_Tax"].sum()
                    st.info(f"**ITC Claimable (Matched): ₹{claimable:,.2f}** · Total 2B: ₹{df['G_Tax'].sum():,.2f} · Total Books: ₹{df['T_Tax'].sum():,.2f} · IMPG: ₹{impg:,.2f}")
                    with st.expander("View party-wise reco"):
                        disp=df[["T_Party","G_Party","T_IGST","G_IGST","T_CGST","G_CGST","D_Tax","Status","Remark"]].copy()
                        disp.columns=["Books Party","2B Party","Books IGST","2B IGST","Books CGST","2B CGST","Diff Tax","Status","Remark"]
                        st.dataframe(disp,use_container_width=True,height=380,
                            column_config={"Books IGST":st.column_config.NumberColumn(format="₹ %,.0f"),"2B IGST":st.column_config.NumberColumn(format="₹ %,.0f"),
                                           "Books CGST":st.column_config.NumberColumn(format="₹ %,.0f"),"2B CGST":st.column_config.NumberColumn(format="₹ %,.0f"),"Diff Tax":st.column_config.NumberColumn(format="₹ %,.0f")})
                    st.download_button("⬇️ Download ITC Reco Excel",data=buf,file_name=f"{client.replace(' ','_')}_ITC_Reco_{period.replace(' ','_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
                    st.caption("Excel: Summary · Auditor Actions · ITC Reco · Tally Detail · IMPG · Legend")
                except Exception as e:
                    st.error(f"❌ {e}")
                    with st.expander("Details"): import traceback; st.code(traceback.format_exc())
    else:
        st.info("Upload both files above to run.")

# ── TAB 2: SALES ──────────────────────────────────────────────────────────────
with tab2:
    st.markdown("### GSTR-1 vs Books (Sales Register)")
    c1,c2=st.columns(2)
    with c1:
        st.markdown("**File 1 — E-Invoice / GSTR-1 Portal Excel**")
        st.markdown('<div class="fmtbox">GST Portal → E-Invoice → Download Excel<br>Must have sheet: <code>b2b, sez, de</code></div>',unsafe_allow_html=True)
        f_einv=st.file_uploader("E-Invoice Excel",type=["xlsx","xls"],key="sal_inv",label_visibility="collapsed")
    with c2:
        st.markdown("**File 2 — Tally Sales Register / GSTR-1 Working**")
        st.markdown('<div class="fmtbox">Tally Sales Working File (multi-sheet: GST Local, IGST Sales, Handling Charges etc.)<br>Needs: <code>Voucher No</code> + <code>Value</code> columns in each sheet</div>',unsafe_allow_html=True)
        f_sal=st.file_uploader("Sales Register",type=["xlsx","xls"],key="sal_bk",label_visibility="collapsed")
    if f_einv and f_sal:
        if st.button("▶ Run Sales Reconciliation",type="primary",use_container_width=True,key="run_sal"):
            with st.spinner("Running sales reco..."):
                try:
                    reco,df_b,df_p=run_sales_reco(f_einv.read(),f_sal.read())
                    cnt=reco["Status"].value_counts().to_dict()
                    buf=build_sales_excel(reco,df_b,df_p,client,period)
                    st.success("✅ Sales Reconciliation complete!")
                    cols=st.columns(5)
                    for col,lbl,key in zip(cols,["✅ Matched","⚠️ Mismatch","🟡 Only Portal","🔵 Only Books","⚪ Exempt"],
                                           ["Matched","Value Mismatch","Only in Portal","Only in Books","Exempt / Bond Transfer"]):
                        col.metric(lbl,cnt.get(key,0))
                    with st.expander("View invoice-wise reco"):
                        disp=reco[["inv_norm","B_Tax","P_Tax","Diff","Status"]].copy()
                        disp.columns=["Invoice No","Books (₹)","Portal (₹)","Diff (₹)","Status"]
                        st.dataframe(disp,use_container_width=True,height=380,
                            column_config={"Books (₹)":st.column_config.NumberColumn(format="₹ %,.0f"),"Portal (₹)":st.column_config.NumberColumn(format="₹ %,.0f"),"Diff (₹)":st.column_config.NumberColumn(format="₹ %,.0f")})
                    st.download_button("⬇️ Download Sales Reco Excel",data=buf,file_name=f"{client.replace(' ','_')}_Sales_Reco_{period.replace(' ','_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
                    st.caption("Excel: Summary · Invoice Reco · Books Detail · Portal Detail")
                except Exception as e:
                    st.error(f"❌ {e}")
                    with st.expander("Details"): import traceback; st.code(traceback.format_exc())
    else:
        st.info("Upload both files above to run.")

# ── TAB 3: 3B CHECK ───────────────────────────────────────────────────────────
with tab3:
    st.markdown("### GSTR-3B Audit Check")
    st.markdown('<div class="fmtbox">Upload GSTR-2B Excel → 2B ITC values auto-fill · Then type 3B and GSTR-1 values from filed PDFs · Click Run</div>',unsafe_allow_html=True)

    # ── Step 1: 2B Upload for auto-fill ──────────────────────────────────────
    st.markdown("**Step 1 — Upload GSTR-2B Excel (auto-fills ITC available values)**")
    f_2b_3b = st.file_uploader("GSTR-2B Excel",type=["xlsx","xls"],key="tb3_2b",label_visibility="collapsed")

    # Parse 2B ITC Available sheet
    af = {"b2b_igst":0.,"b2b_cgst":0.,"b2b_sgst":0.,"impg":0.,"rcm_cgst":0.,"rcm_sgst":0.,"isd_cgst":0.,"isd_sgst":0.,"cn_cgst":0.,"cn_sgst":0.}
    if f_2b_3b:
        try:
            wb_af = load_wb(f_2b_3b.read())
            if "ITC Available" in wb_af.sheetnames:
                ws_af = wb_af["ITC Available"]
                for row in ws_af.iter_rows(min_row=6, max_row=ws_af.max_row, values_only=True):
                    if not row[2]: continue
                    ref = str(row[2] or "").strip()
                    igst = safe(row[3]); cgst = safe(row[4]); sgst = safe(row[5] if len(row)>5 else 0)
                    if "4(A)(5)" in ref and "4(A)(3)" not in ref and "4(A)(4)" not in ref:
                        if row[0] in ("I","Details") and igst+cgst>0:
                            if igst > 0: af["b2b_igst"] += igst
                            if cgst > 0: af["b2b_cgst"] += cgst
                            if sgst > 0: af["b2b_sgst"] += sgst
                    if "4(A)(3)" in ref and cgst > 0: af["rcm_cgst"] += cgst; af["rcm_sgst"] += sgst
                    if "4(A)(4)" in ref and cgst > 0: af["isd_cgst"] += cgst; af["isd_sgst"] += sgst
                    if "4(A)(1)" in ref and igst > 0: af["impg"] += igst
                    if "4(A)(5)" in ref and cgst < 0: af["cn_cgst"] += abs(cgst); af["cn_sgst"] += abs(sgst)
                # Also parse credit notes (Part B) — net off
                for row in ws_af.iter_rows(min_row=6, max_row=ws_af.max_row, values_only=True):
                    if row[0] == "Details" and row[1] and "credit note" in str(row[1]).lower():
                        af["b2b_cgst"] = max(0, af["b2b_cgst"] - safe(row[4]))
                        af["b2b_sgst"] = max(0, af["b2b_sgst"] - safe(row[5] if len(row)>5 else 0))
                st.success(f"✅ 2B auto-filled — B2B IGST: ₹{af['b2b_igst']:,.2f} · CGST: ₹{af['b2b_cgst']:,.2f} · IMPG: ₹{af['impg']:,.2f} · RCM: ₹{af['rcm_cgst']:,.2f}")
        except Exception as e:
            st.warning(f"Could not parse ITC Available sheet: {e}")

    st.markdown("---")
    st.markdown("**Step 2 — Enter values from filed returns (type from PDF)**")

    col_a, col_b, col_c = st.columns(3)

    with col_a:
        st.markdown("**From GSTR-3B (Filed PDF)**")
        b3_out_taxable = st.number_input("3.1(a) Outward Taxable",min_value=0.,value=0.,step=1000.,format="%.2f",key="b3_ot")
        b3_out_igst    = st.number_input("3.1(a) IGST on Outward",min_value=0.,value=0.,step=100.,format="%.2f",key="b3_oi")
        b3_out_cgst    = st.number_input("3.1(a) CGST on Outward",min_value=0.,value=0.,step=100.,format="%.2f",key="b3_oc")
        b3_out_sgst    = st.number_input("3.1(a) SGST on Outward",min_value=0.,value=0.,step=100.,format="%.2f",key="b3_os")
        st.markdown("---")
        b3_itc_import  = st.number_input("4(A)(1) Import of Goods ITC",min_value=0.,value=0.,step=1000.,format="%.2f",key="b3_ii")
        b3_itc_rcm_c   = st.number_input("4(A)(3) RCM ITC — CGST",min_value=0.,value=0.,step=100.,format="%.2f",key="b3_rc")
        b3_itc_rcm_s   = st.number_input("4(A)(3) RCM ITC — SGST",min_value=0.,value=0.,step=100.,format="%.2f",key="b3_rs")
        b3_itc_igst    = st.number_input("4(A)(5) Other ITC — IGST",min_value=0.,value=0.,step=1000.,format="%.2f",key="b3_ig")
        b3_itc_cgst    = st.number_input("4(A)(5) Other ITC — CGST",min_value=0.,value=0.,step=1000.,format="%.2f",key="b3_cg")
        b3_itc_sgst    = st.number_input("4(A)(5) Other ITC — SGST",min_value=0.,value=0.,step=1000.,format="%.2f",key="b3_sg")
        b3_net_itc     = st.number_input("4(C) Net ITC Available",min_value=0.,value=0.,step=1000.,format="%.2f",key="b3_net")

    with col_b:
        st.markdown("**From GSTR-1 (Filed PDF)**")
        g1_b2b_taxable = st.number_input("4A B2B Taxable Value",min_value=0.,value=0.,step=1000.,format="%.2f",key="g1_bt")
        g1_b2b_igst    = st.number_input("4A IGST",min_value=0.,value=0.,step=100.,format="%.2f",key="g1_bi")
        g1_b2b_cgst    = st.number_input("4A CGST (incl. B2CS)",min_value=0.,value=0.,step=100.,format="%.2f",key="g1_bc")
        g1_b2b_sgst    = st.number_input("4A SGST (incl. B2CS)",min_value=0.,value=0.,step=100.,format="%.2f",key="g1_bs")

    with col_c:
        st.markdown("**2B ITC Available** *(auto-filled from upload above)*")
        v2b_b2b_igst = st.number_input("4(A)(5) B2B IGST",min_value=0.,value=af["b2b_igst"],step=1000.,format="%.2f",key="v2b_ii")
        v2b_b2b_cgst = st.number_input("4(A)(5) B2B CGST",min_value=0.,value=af["b2b_cgst"],step=1000.,format="%.2f",key="v2b_ic")
        v2b_b2b_sgst = st.number_input("4(A)(5) B2B SGST",min_value=0.,value=af["b2b_sgst"],step=1000.,format="%.2f",key="v2b_is")
        v2b_impg     = st.number_input("4(A)(1) IMPG IGST",min_value=0.,value=af["impg"],step=1000.,format="%.2f",key="v2b_im")
        v2b_rcm_c    = st.number_input("4(A)(3) RCM CGST",min_value=0.,value=af["rcm_cgst"],step=100.,format="%.2f",key="v2b_rc")
        v2b_rcm_s    = st.number_input("4(A)(3) RCM SGST",min_value=0.,value=af["rcm_sgst"],step=100.,format="%.2f",key="v2b_rs")
        v2b_isd_c    = st.number_input("4(A)(4) ISD CGST",min_value=0.,value=af["isd_cgst"],step=100.,format="%.2f",key="v2b_dc")
        v2b_isd_s    = st.number_input("4(A)(4) ISD SGST",min_value=0.,value=af["isd_sgst"],step=100.,format="%.2f",key="v2b_ds")
        if af["b2b_igst"]+af["b2b_cgst"]+af["impg"] > 0:
            st.caption("✅ Values auto-filled from 2B upload. Edit if needed.")

    if st.button("▶ Run 3B Check",type="primary",use_container_width=True,key="run_3b"):
        checks=[]
        def chk(label,v3b,vref,higher_ok=False,allow_reclaim_note=""):
            diff=v3b-vref
            if not higher_ok: ok=abs(diff)<2
            else: ok=diff<=1
            icon="✅" if ok else "⚠️"
            if ok: status="Match"
            elif diff>0 and higher_ok: status=f"3B > 2B — excess claim ₹{diff:,.0f}"
            elif diff<0 and not higher_ok: status=f"3B < GSTR-1 — short paid ₹{abs(diff):,.0f} — interest risk"
            else: status=f"3B > GSTR-1 — variance ₹{diff:,.0f}"
            if allow_reclaim_note and ok is False and diff>0: status+=f" (may be reclaim — verify)"
            checks.append((icon,label,v3b,vref,diff,status))

        # Output checks
        chk("IGST on Outward (3B vs GSTR-1)",b3_out_igst,g1_b2b_igst)
        chk("CGST on Outward (3B vs GSTR-1)",b3_out_cgst,g1_b2b_cgst)
        chk("SGST on Outward (3B vs GSTR-1)",b3_out_sgst,g1_b2b_sgst)
        # ITC checks
        chk("4(A)(5) Other ITC IGST vs 2B",b3_itc_igst,v2b_b2b_igst,higher_ok=True,allow_reclaim_note="yes")
        chk("4(A)(5) Other ITC CGST vs 2B",b3_itc_cgst,v2b_b2b_cgst,higher_ok=True,allow_reclaim_note="yes")
        chk("4(A)(5) Other ITC SGST vs 2B",b3_itc_sgst,v2b_b2b_sgst,higher_ok=True,allow_reclaim_note="yes")
        chk("4(A)(1) Import IGST vs IMPG",b3_itc_import,v2b_impg,higher_ok=True,allow_reclaim_note="yes")
        chk("4(A)(3) RCM CGST vs 2B",b3_itc_rcm_c,v2b_rcm_c,higher_ok=True)
        chk("4(A)(3) RCM SGST vs 2B",b3_itc_rcm_s,v2b_rcm_s,higher_ok=True)

        issues=[c for c in checks if c[0]=="⚠️"]
        if issues: st.warning(f"⚠️ {len(issues)} issue(s) found")
        else: st.success("✅ 3B appears clean")

        df_check=pd.DataFrame(checks,columns=["","Description","3B (Filed)","2B / GSTR-1 Ref","Diff","Status"])
        st.dataframe(df_check,use_container_width=True,hide_index=True,
            column_config={"3B (Filed)":st.column_config.NumberColumn(format="₹ %,.2f"),
                           "2B / GSTR-1 Ref":st.column_config.NumberColumn(format="₹ %,.2f"),
                           "Diff":st.column_config.NumberColumn(format="₹ %,.2f")})

        data_3b={"out_taxable":b3_out_taxable,"out_igst":b3_out_igst,"out_cgst":b3_out_cgst,"out_sgst":b3_out_sgst,
                 "itc_import":b3_itc_import,"itc_rcm":b3_itc_rcm_c+b3_itc_rcm_s,"itc_igst":b3_itc_igst,
                 "itc_cgst":b3_itc_cgst,"itc_sgst":b3_itc_sgst,"net_itc":b3_net_itc}
        data_1={"b2b_taxable":g1_b2b_taxable,"b2b_igst":g1_b2b_igst,"b2b_cgst":g1_b2b_cgst,"b2b_sgst":g1_b2b_sgst}
        data_2b_={"b2b_igst":v2b_b2b_igst,"b2b_cgst":v2b_b2b_cgst,"b2b_sgst":v2b_b2b_sgst,
                  "impg_igst":v2b_impg,"rcm_total":v2b_rcm_c+v2b_rcm_s,"net_available":b3_net_itc}
        buf=build_3b_excel(data_3b,data_1,data_2b_,client,period)
        st.download_button("⬇️ Download 3B Check Excel",data=buf,
            file_name=f"{client.replace(' ','_')}_3B_Check_{period.replace(' ','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)

    with st.expander("📋 Where to find 3B and GSTR-1 values"):
        st.markdown("""
**GSTR-3B values** — portal → Returns → GSTR-3B → View filed return:
- `3.1(a)`: Outward taxable supplies (Taxable, IGST, CGST, SGST)
- `4(A)(1)`: Import of goods ITC (IGST)
- `4(A)(3)`: RCM ITC (CGST + SGST)
- `4(A)(5)`: All other ITC (IGST + CGST + SGST)
- `4(C)`: Net ITC available

**GSTR-1 values** — portal → Returns → GSTR-1 → View filed return:
- `Table 4A`: B2B Taxable + IGST + CGST + SGST
- `Table 7 (B2CS)`: Add CGST/SGST to Table 4A totals

**2B ITC** — auto-filled when you upload 2B Excel above ✅
        """)

st.markdown("---")
st.markdown('<div style="text-align:center;font-size:.74rem;color:#888">GST Audit Reconciliation Tool · Built by <a href="https://linkedin.com/in/purvadoshi26" style="color:#5585B5">Purva Doshi</a></div>',unsafe_allow_html=True)
