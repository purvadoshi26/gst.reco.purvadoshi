"""
GST Audit Reconciliation Tool - Streamlit UI
Created by: Purva Doshi | linkedin.com/in/purvadoshi26
Tabs: ITC Reco | Sales Reco | 3B vs GSTR-1 Check
All parsing / matching / Excel logic lives in gst_core.py (testable without Streamlit).
"""
import streamlit as st
import io, re, datetime, hashlib
import pandas as pd
import openpyxl
from collections import defaultdict
import importlib
import gst_core as _core
_core = importlib.reload(_core)
# bring every core name (including _private helpers) into this module's namespace
globals().update({k: v for k, v in vars(_core).items() if not k.startswith("__") and k != "st"})

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
    st.markdown("---")
    st.warning("🔒 Client data: run this app locally / on firm-approved infrastructure. "
               "Do not upload real client files to a public cloud deployment without the firm's approval.")

if not client.strip(): client="Client"
if not period.strip(): period="Period"

tab1, tab2, tab3 = st.tabs(["🔵 ITC Reco (2B vs Purchase)", "🟢 Sales Reco (GSTR-1 vs Books)", "📋 3B Check"])

def _meta_line(m):
    return " · ".join(x for x in (m.get("Legal name",""),m.get("GSTIN",""),(str(m.get("Period",""))+" "+str(m.get("FY",""))).strip()) if x)
def _check_meta(items):
    """items = [(source label, meta dict)]; show an error if GSTIN or tax period differ across files"""
    items=[(n,m) for n,m in items if m]
    g={n:m.get("GSTIN") for n,m in items if m.get("GSTIN")}; p={n:m.get("pstart") for n,m in items if m.get("pstart")}
    if len(set(g.values()))>1: st.error("⛔ These files belong to DIFFERENT GSTINs: "+" · ".join(f"{n}: {v}" for n,v in g.items())+" — do not mix clients.")
    if len(set(p.values()))>1: st.error("⛔ These files are for DIFFERENT tax periods: "+" · ".join(f"{n}: {v:%b-%Y}" for n,v in p.items()))
def _ups(files):
    """uploaded file(s) (one, many or None) -> [(name, bytes)]"""
    if files is None: return []
    if not isinstance(files,list): files=[files]
    return [(f.name,f.getvalue()) for f in files]
def _sig(*groups):
    h=hashlib.md5()
    for g_ in groups:
        for n_,b_ in g_: h.update(n_.encode()); h.update(b_)
        h.update(b"|")
    return h.hexdigest()[:10]
def _fsig(*keys):
    """signature of the files CURRENTLY uploaded under these widget keys (None if the first two groups are empty)"""
    gs=[_ups(st.session_state.get(k)) for k in keys]
    return None if (len(gs)<2 or not gs[0] or not gs[1]) else _sig(*gs)
def _mlabel(m):
    ps=(m or {}).get("pstart"); return f"{ps:%b-%Y}" if ps else str((m or {}).get("Period",""))
def _ym(d): return (d.year,d.month)
def _chk_period(books_lbl,portal_lbl,dates,bmeta,pmeta,edu_n=0):
    """books must cover the SAME months as the portal files (checks first AND last date)"""
    lo=pmeta.get("pstart"); hi=pmeta.get("pend") or lo
    if not lo: return
    rng=f"{lo:%b-%Y}" if _ym(lo)==_ym(hi) else f"{lo:%b-%Y} to {hi:%b-%Y}"
    blo=bmeta.get("pstart"); bhi=bmeta.get("pend") or blo
    if blo and (_ym(blo)<_ym(lo) or _ym(bhi)>_ym(hi)):
        st.error(f"⛔ TAX PERIOD MISMATCH: {books_lbl} indicate period '{bmeta.get('period')}', but {portal_lbl} cover {rng} ({pmeta.get('FY','')})! Upload matching months on both sides - each month as its own file.")
    elif dates and edu_n==0:
        mn,mx=min(dates),max(dates)
        if _ym(mn)<_ym(lo) or _ym(mx)>_ym(hi):
            st.error(f"⛔ TAX PERIOD MISMATCH: {books_lbl} are dated {mn:%d-%b-%Y} to {mx:%d-%b-%Y}, but {portal_lbl} cover {rng} ({pmeta.get('FY','')})! Add the missing month's portal file (or remove the extra month from books) so both sides cover the same months.")
def _stamp(*objs):
    return hashlib.md5("|".join(repr(getattr(o,"values",o).tolist() if hasattr(getattr(o,"values",None),"tolist") else o) for o in objs).encode()).hexdigest()[:10]

def _recv_line(*groups):
    """One line showing what the app has ACTUALLY received from each uploader: (label, files, optional).
    If the browser lists a file above but this line says 'not received yet', the upload never reached the
    server - remove it (x) and add it again."""
    parts=[]
    for lbl,files,opt in groups:
        n=len(files or [])
        parts.append(f"{'✅' if n else ('➖' if opt else '⬜')} **{lbl}**: "+(f"{n} file{'s' if n!=1 else ''}" if n else ("none (optional)" if opt else "not received yet")))
    st.caption("Received by the app → "+"  ·  ".join(parts))
def _locked_run(label,key,need):
    """Run button that is ALWAYS on screen: greyed out until the required files have been received and read."""
    st.info(f"Add {need} above - the Run button unlocks as soon as they are received and read. "
            "If a file is listed in the box above but shows 'not received yet' here, the upload did not reach the app: click ✕ on it and add it again (or refresh the page).")
    st.button(label,type="primary",width="stretch",key=key,disabled=True,help="Unlocks when the required files are uploaded.")

# ── TAB 1: ITC (invoice-wise) ─────────────────────────────────────────────────
with tab1:
    st.markdown("### GSTR-2B vs Purchase Register — invoice-wise ITC reco")
    c1,c2=st.columns(2)
    with c1:
        st.markdown("**File 1 — GSTR-2B Excel(s) (portal)** — one file per month")
        st.markdown('<div class="fmtbox">Portal → Returns → GSTR-2B → Download Excel<br>All sheets are read by <b>header name</b>: B2B · B2BA (amendments) · B2B-CDNR (credit/debit notes) · ISD · IMPG. Gets ITC-availability + reverse-charge flags per invoice.<br><b>Several months?</b> Select all the monthly 2B files together (or add them one by one). Do <u>not</u> paste months into one sheet.</div>',unsafe_allow_html=True)
        f_2b=st.file_uploader("GSTR-2B",type=["xlsx","xls"],key="itc_2b",label_visibility="collapsed",accept_multiple_files=True)
    with c2:
        st.markdown("**File 2 — Purchase Register (Tally / SAP / Excel)** — one file per month, or one for all")
        st.markdown('<div class="fmtbox"><b>Best:</b> flat export — one row per voucher with <code>Voucher No</code> · <code>Supplier Invoice No</code> · <code>GSTIN/UIN</code> · <code>Date</code> + purchase / Input IGST / CGST / SGST ledger columns (any order, any extra columns). Debit-note vouchers in the same file are picked up.<br><b>Also accepted:</b> older Tally sub-row layout · pivot summary (party-level only)</div>',unsafe_allow_html=True)
        f_pr=st.file_uploader("Purchase Register",type=["xlsx","xls"],key="itc_pr",label_visibility="collapsed",accept_multiple_files=True)
        st.markdown("**File 2b — Debit Note Register(s)** *(optional — one or more files, if exported separately)*")
        f_dn=st.file_uploader("Debit Note Register (optional)",type=["xlsx","xls"],key="itc_dn",label_visibility="collapsed",accept_multiple_files=True)
    _recv_line(("GSTR-2B",f_2b,False),("Purchase Register",f_pr,False),("Debit notes",f_dn,True))
    _gate=st.empty()      # one fixed slot: locked Run button / 'reading files' state - replaced as a whole, so nothing stale is left behind
    if f_2b and f_pr:
        eb=_ups(f_2b); pbb=_ups(f_pr); dnb=_ups(f_dn) or None
        sig=_sig(pbb,eb,dnb or []); scan=None; pt=None
        try:
            if st.session_state.get("itc_sig")!=sig:
                with _gate.container():
                    st.button("▶ Run ITC Reconciliation",type="primary",width="stretch",key="run_itc_wait",disabled=True)
                    with st.spinner("⏳ Reading the uploaded files... the Run button unlocks as soon as this finishes"):
                        s2b=_sig(eb); sbk=_sig(pbb,dnb or [])      # parse each side only when ITS files changed (adding a debit-note file no longer re-reads the 2B)
                        if st.session_state.get("itc_s2b")!=s2b: st.session_state["itc_pt"]=parse_2b_multi(eb); st.session_state["itc_s2b"]=s2b
                        if st.session_state.get("itc_sbk")!=sbk: st.session_state["itc_scan"]=scan_pbooks_any(pbb,dn_fb=dnb); st.session_state["itc_sbk"]=sbk
                st.session_state["itc_sig"]=sig; st.session_state.pop("itc_out",None); st.session_state.pop("itc_run",None)
            scan=st.session_state["itc_scan"]; pt=st.session_state["itc_pt"]
        except Exception as e:
            st.session_state.pop("itc_sig",None)      # never leave half-updated results marked as current
            st.error(f"❌ {e}")
            with st.expander("Details"): import traceback; st.code(traceback.format_exc())
        _gate.empty()      # files are read -> the locked placeholder makes way for the real Run button below
        if scan is not None:
            pivot=(scan["kind"]=="legacy" and (scan["bw"]["_inv"].astype(str).str.strip()=="").all())
            if pt.get("meta"): st.caption(f"GSTR-2B ({len(pt['parts'])} file{'s' if len(pt['parts'])>1 else ''}): "+_meta_line(pt["meta"]))
            if len(pt["parts"])>1 or len(scan["layout"])>1:
                with st.expander("📁 Files read",expanded=len(pt["parts"])>1):
                    st.dataframe(pd.DataFrame([{"2B file":p_["name"],"Month":_mlabel(p_["meta"]),"Invoices":p_["n_inv"],"Notes (CN/DN)":p_["n_notes"],"ISD":p_["n_isd"],"IMPG":p_["n_imp"]} for p_ in pt["parts"]]),hide_index=True,width="stretch")
                    st.dataframe(pd.DataFrame([{"Books file / sheet":L_["sheet"],"Vouchers":L_["rows"]} for L_ in scan["layout"]]),hide_index=True,width="stretch")
            _bm = scan.get("meta", {})
            if _bm.get("company"): st.caption(f"Books: {_bm['company']} · {_bm.get('period','')}")
            pdates=[d for d in scan["bw"]["_date"] if pd.notna(d) and hasattr(d, "year")]
            edu_n = _bm.get("edu_count", 0)
            _chk_period("Purchase vouchers","GSTR-2B file(s)",pdates,_bm,pt.get("meta",{}),edu_n)
            if edu_n > 0:
                st.info(f"ℹ️ Tally Educational Mode detected: {edu_n} vouchers have 'Educational' date. Vouchers are reconciled on GSTIN, Invoice Number, and Tax amounts.")
            for w in scan["warns"]+pt["warns"]: st.warning(w)
            if pivot:
                if len(pbb)>1 or len(eb)>1: st.warning("Party-wise (pivot) mode reads only the FIRST purchase file and FIRST 2B file. Export the flat Purchase Register to reconcile several months invoice-wise.")
                st.warning("Purchase file is a party-level pivot (no invoice numbers) → running the party-wise reco. Export the flat Purchase Register for invoice-wise matching.")
                if st.button("▶ Run ITC Reconciliation (party-wise)",type="primary",width="stretch",key="run_itc_old"):
                    try:
                        df,det,impg,impg_r=run_itc_reco(pbb[0][1],eb[0][1]); st.session_state["itc_old"]=(df,build_itc_excel(df,det,impg_r,impg,client,period).getvalue(),impg)
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
                            res=run_itc_v2(pbb,eb,lc,vc,tol,mat,scan=scan,pt=pt,dn_bytes=dnb)
                            rep_client = client if client != "Client" else (scan.get("meta",{}).get("company") or pt.get("meta",{}).get("Legal name") or "Client")
                            rep_period = period if period != "Period" else (scan.get("meta",{}).get("period") or pt.get("meta",{}).get("Period") or "Period")
                            st.session_state["itc_out"]=(res,build_itc_excel_v2(res,rep_client,rep_period).getvalue()); st.session_state["itc_run"]=_stamp(led_ed,vt_ed,tol,mat)
                        except Exception as e:
                            st.session_state.pop("itc_out",None); st.error(f"❌ {e}")
                            with st.expander("Details"): import traceback; st.code(traceback.format_exc())
                if st.session_state.get("itc_out"):
                    res,xb=st.session_state["itc_out"]; inv=res["inv"]; cnt=inv.Status.value_counts().to_dict()
                    st.success("✅ ITC Reconciliation complete!")
                    if st.session_state.get("itc_run")!=_stamp(led_ed,vt_ed,tol,mat): st.warning("Ledger classes / voucher types / tolerance changed after the last run - press Run again to refresh these results.")
                    cols=st.columns(6)
                    for col,lbl,keys in zip(cols,["✅ Matched","⚠️ Mismatch","🟡 Only 2B","🔵 Only Books","⛔ ITC N/A on 2B","🟢 RCM"],
                            [["Matched","Matched - Inv No differs","Matched - Rounding","Matched - RCM"],["Mismatch"],["Only in 2B"],["Only in Books"],["2B: ITC not available"],["RCM - verify"]]):
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
        with _gate.container(): _locked_run("▶ Run ITC Reconciliation","run_itc_locked","the **GSTR-2B file(s)** and the **Purchase Register**")

# ── TAB 2: SALES ──────────────────────────────────────────────────────────────
with tab2:
    st.markdown("### E-Invoice / GSTR-1 vs Books (Sales Register)")
    c1,c2=st.columns(2)
    with c1:
        st.markdown("**File 1 — E-Invoice / GSTR-1 Portal Excel(s)** — one file per month")
        st.markdown('<div class="fmtbox">GST Portal → E-Invoice → Download Excel<br>Sheets (invoices, credit/debit notes, HSN) are found by <b>header names</b> — sheet names and column order can be anything.<br>Needs headers: <code>Invoice number</code> · <code>Taxable Value</code> · <code>GSTIN</code> · tax columns</div>',unsafe_allow_html=True)
        f_einv=st.file_uploader("E-Invoice Excel",type=["xlsx","xls"],key="sal_inv",label_visibility="collapsed",accept_multiple_files=True)
    with c2:
        st.markdown("**File 2 — Sales Register (Tally / SAP / Excel)** — one file per month, or one for all")
        st.markdown('<div class="fmtbox">Any company format — columns are found by <b>header name</b>, in any order.<br>Needs headers: <code>Voucher No / Invoice No</code> · <code>Date</code> · <code>GSTIN</code> · <code>Gross Total</code> (optional) + your sales / IGST / CGST / SGST ledger columns. Ledger columns are auto-classified — review below.</div>',unsafe_allow_html=True)
        f_sal=st.file_uploader("Sales Register",type=["xlsx","xls"],key="sal_bk",label_visibility="collapsed",accept_multiple_files=True)
        st.markdown("**File 2b — Credit Note Register(s)** *(optional — one or more files, if exported separately)*")
        f_cn=st.file_uploader("Credit Note Register (optional)",type=["xlsx","xls"],key="sal_cn",label_visibility="collapsed",accept_multiple_files=True)
    _recv_line(("E-invoice / GSTR-1",f_einv,False),("Sales Register",f_sal,False),("Credit notes",f_cn,True))
    _gate=st.empty()
    if f_einv and f_sal:
        eb=_ups(f_einv); bb=_ups(f_sal); cb=_ups(f_cn) or None
        sig=_sig(eb,bb,cb or [])
        scan=None
        try:
            if st.session_state.get("sal_sig")!=sig:
                with _gate.container():
                    st.button("▶ Run Sales Reconciliation",type="primary",width="stretch",key="run_sal_wait",disabled=True)
                    with st.spinner("⏳ Reading the uploaded files... the Run button unlocks as soon as this finishes"):
                        sbk=_sig(bb,cb or []); spt=_sig(eb)        # parse each side only when ITS files changed
                        if st.session_state.get("sal_sbk")!=sbk: st.session_state["sal_scan"]=scan_books(bb,cn_fb=cb); st.session_state["sal_sbk"]=sbk
                        if st.session_state.get("sal_spt")!=spt: st.session_state["sal_pt"]=parse_portal_multi(eb); st.session_state["sal_spt"]=spt
                st.session_state["sal_sig"]=sig; st.session_state.pop("sal_out",None); st.session_state.pop("sal_run",None)
            scan=st.session_state["sal_scan"]
        except Exception as e:
            st.session_state.pop("sal_sig",None)
            st.error(f"❌ {e}")
            with st.expander("Details"): import traceback; st.code(traceback.format_exc())
        _gate.empty()
        if scan:
            try:
                ptl=st.session_state["sal_pt"]; _pm=ptl.get("meta",{})
                if _pm.get("GSTIN") or _pm.get("Period"): st.caption(f"E-invoice ({len(ptl['parts'])} file{'s' if len(ptl['parts'])>1 else ''}): "+_meta_line(_pm))
                if len(ptl["parts"])>1 or len(scan["layout"])>1:
                    with st.expander("📁 Files read",expanded=len(ptl["parts"])>1):
                        st.dataframe(pd.DataFrame([{"Portal file":p_["name"],"Month":_mlabel(p_["meta"]),"Invoices":p_["n_inv"],"Notes (CN/DN)":p_["n_notes"]} for p_ in ptl["parts"]]),hide_index=True,width="stretch")
                        st.dataframe(pd.DataFrame([{"Books file / sheet":L_["sheet"],"Vouchers":L_["rows"]} for L_ in scan["layout"]]),hide_index=True,width="stretch")
                _bm = scan.get("meta", {})
                if _bm.get("company"): st.caption(f"Books: {_bm['company']} · {_bm.get('period','')}")
                bdates=[d for d in scan["bw"]["_date"] if pd.notna(d) and hasattr(d, "year")]
                edu_n = _bm.get("edu_count", 0)
                _chk_period("Books vouchers","E-Invoice portal file(s)",bdates,_bm,_pm,edu_n)
                if edu_n > 0:
                    st.info(f"ℹ️ Tally Educational Mode detected: {edu_n} vouchers have 'Educational' date. Vouchers are reconciled on GSTIN, Invoice Number, and Tax amounts.")
            except Exception: pass
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
                        res=run_sales_reco(None,None,lc,vc,tol,scan=scan,portal=st.session_state["sal_pt"])
                        rep_client = client if client != "Client" else (scan.get("meta",{}).get("company") or _pm.get("Legal name") or "Client")
                        rep_period = period if period != "Period" else (scan.get("meta",{}).get("period") or _pm.get("Period") or "Period")
                        buf=build_sales_excel(res,rep_client,rep_period)
                        st.session_state["sal_out"]=(res,buf.getvalue()); st.session_state["sal_run"]=_stamp(led_ed,vt_ed,tol)
                    except Exception as e:
                        st.session_state.pop("sal_out",None); st.error(f"❌ {e}")
                        with st.expander("Details"): import traceback; st.code(traceback.format_exc())
            if st.session_state.get("sal_out"):
                res,xbytes=st.session_state["sal_out"]; reco=res["reco"]; cnt=reco["Status"].value_counts().to_dict(); bk=res["books"]
                st.success("✅ Sales Reconciliation complete!")
                if st.session_state.get("sal_run")!=_stamp(led_ed,vt_ed,tol): st.warning("Ledger classes / voucher types / tolerance changed after the last run - press Run again to refresh these results.")
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
        with _gate.container(): _locked_run("▶ Run Sales Reconciliation","run_sal_locked","the **E-invoice / GSTR-1 file(s)** and the **Sales Register**")

# ── TAB 3: 3B CHECK (v2) ──────────────────────────────────────────────────────
with tab3:
    st.markdown("### GSTR-3B Audit Check — 3B vs GSTR-1 vs Books · 3B vs 2B vs ITC reco")
    st.markdown('<div class="fmtbox">Type the filed <b>GSTR-3B</b> and <b>GSTR-1</b> figures (from the portal PDF/summary). <b>2B</b>, <b>books sales</b> and <b>ITC-reco</b> references fill automatically when you have uploaded the 2B here or run the Sales / ITC tabs in this session. Every cell stays editable.</div>',unsafe_allow_html=True)
    st.markdown("**Step 1 — GSTR-2B Excel** *(optional if the ITC tab already has it)*")
    f_2b3=st.file_uploader("GSTR-2B Excel",type=["xlsx","xls"],key="tb3_2b",label_visibility="collapsed",accept_multiple_files=True)
    pt3=None; src2b=""
    if f_2b3:
        try: pt3=parse_2b_multi(_ups(f_2b3)); src2b="uploaded 2B"
        except Exception as e: st.warning(f"Could not read 2B: {e}")
    itc_valid=st.session_state.get("itc_sig") is not None and st.session_state.get("itc_sig")==_fsig("itc_pr","itc_2b","itc_dn")
    sal_valid=st.session_state.get("sal_sig") is not None and st.session_state.get("sal_sig")==_fsig("sal_inv","sal_bk","sal_cn")
    if (not f_2b3) and itc_valid and st.session_state.get("itc_pt") is not None: pt3=st.session_state["itc_pt"]; src2b="2B from ITC tab ("+", ".join(p_["name"] for p_ in pt3["parts"])+")"
    # a 3B covers ONE month: when several months were uploaded, pick the month(s) that match this 3B
    multi3=bool(pt3 and len(pt3["parts"])>1); sel_files=None; sel_months=None
    if multi3:
        lab3={p_["name"]:f"{p_['name']} ({_mlabel(p_['meta'])})" for p_ in pt3["parts"]}
        pick=st.multiselect("This 3B is for ONE month - which 2B month(s) should be used?",list(lab3),default=list(lab3)[:1],format_func=lambda n:lab3[n],key="tb3_pick")
        if not pick: pick=list(lab3)[:1]
        pt3=merge_2b_parts([p_ for p_ in pt3["parts"] if p_["name"] in pick]); sel_files=set(pick)
    if pt3: sel_months={(p_["meta"]["pstart"].year,p_["meta"]["pstart"].month) for p_ in pt3["parts"] if p_["meta"].get("pstart")} or None
    def _pick_portal(pe):
        """e-invoice / GSTR-1 files for the selected month(s) only"""
        if pe is None or len(pe.get("parts",[]))<2 or not sel_months: return pe
        ps=[p_ for p_ in pe["parts"] if p_["meta"].get("pstart") and (p_["meta"]["pstart"].year,p_["meta"]["pstart"].month) in sel_months]
        return merge_portal_parts(ps) if ps else pe
    t2b_def=twob_refs(pt3) if pt3 else zero_tbl(ROWS_2B)
    sal=st.session_state.get("sal_out") if sal_valid else None; itcx=st.session_state.get("itc_out") if itc_valid else None
    tbk_def=zero_tbl(ROWS_BK); tib_def=zero_tbl(ROWS_IB); tei_def=zero_tbl(ROWS_EI); srcs=[]
    st.markdown("**Step 1b — E-invoice / GSTR-1 portal Excel** *(optional — gives a B2B outward reference before you type GSTR-1)*")
    f_ei=st.file_uploader("E-invoice Excel",type=["xlsx","xls"],key="tb3_ei",label_visibility="collapsed",accept_multiple_files=True)
    metas=[]
    if pt3 and pt3.get("meta"): metas.append(("2B",pt3["meta"]))
    if f_ei:
        try: pe_=_pick_portal(parse_portal_multi(_ups(f_ei))); tei_def=einv_net(pe_); srcs.append("e-invoice reference ← uploaded file(s)"); metas.append(("E-invoice",pe_.get("meta",{})))
        except Exception as e: st.warning(f"Could not read e-invoice file: {e}")
    elif sal:
        try:
            pe_=_pick_portal(sal[0]["portal"]); tei_def=einv_net(pe_); metas.append(("E-invoice",pe_.get("meta",{}))); srcs.append("e-invoice reference ← Sales Reco tab ("+", ".join(n_ for n_,_ in _ups(st.session_state.get("sal_inv")))+")")
        except Exception as e:
            st.warning(f"Could not load e-invoice reference from Sales tab: {e}")
    if sal:
        try:
            bs=books_sales_net(sal[0],months=sel_months if multi3 else None)
            for c in COLS3: tbk_def.loc[0,c]=bs.get(c,0.0)
            srcs.append("books sales ← Sales Reco tab ("+", ".join(n_ for n_,_ in _ups(st.session_state.get("sal_bk")))+")")
        except Exception as e:
            st.warning(f"Could not load books sales reference from Sales tab: {e}")
    if itcx:
        _itc_names={p_["name"] for p_ in itcx[0]["pt"]["parts"]}
        _use_f=sel_files if (multi3 and sel_files and sel_files<=_itc_names) else None
        e,rc=itc_reco_refs(itcx[0],files=_use_f,months=sel_months if _use_f else None)
        if multi3 and not _use_f: st.warning("ITC reference covers ALL months run in the ITC tab (the 2B files chosen here are not the ones used there) - edit the references or run that month alone.")
        for k,c in enumerate(COLS3[1:]): tib_def.loc[0,c]=e[k]; tib_def.loc[1,c]=rc[k]
        srcs.append("ITC reference ← ITC Reco tab ("+", ".join(n_ for n_,_ in _ups(st.session_state.get("itc_pr")))+")")
    if src2b: srcs.append(src2b)
    if srcs: st.caption("Auto-filled: "+" · ".join(srcs))
    ssig=hashlib.md5((str(t2b_def.values.tolist())+str(tbk_def.values.tolist())+str(tib_def.values.tolist())+str(tei_def.values.tolist())).encode()).hexdigest()[:8]
    cfg={c:st.column_config.NumberColumn(c,format="%.2f",step=1.0) for c in COLS3}
    st.markdown("---"); st.markdown("**Step 2 — GSTR-3B (filed)** — upload the filed PDF from the GST portal (read automatically), or type the values")
    f_3b=st.file_uploader("Filed GSTR-3B PDF",type=["pdf"],key="tb3_pdf",label_visibility="collapsed")
    t3b_def=zero_tbl(ROWS_3B); pay3=None; nongst3=0.0
    if f_3b:
        try:
            t3b_def,m3,w3=parse_3b_pdf(f_3b.getvalue()); pay3=m3.get("pay"); nongst3=m3.get("nongst",0.0) or 0.0
            metas.append(("3B PDF",{"GSTIN":m3.get("GSTIN"),"pstart":_pstart(m3.get("Period"),m3.get("Year"))}))
            st.success(f"Read GSTR-3B · {m3.get('Legal name','')} · {m3.get('GSTIN','')} · {m3.get('Period','')} {m3.get('Year','')} · ARN {m3.get('ARN','')}")
            for w_ in w3: st.warning(w_)
        except Exception as e: st.error(f"Could not read the 3B PDF: {e} - type the values below instead.")
    _check_meta(metas)
    s3sig=hashlib.md5(str(t3b_def.values.tolist()).encode()).hexdigest()[:8]
    e3b=st.data_editor(t3b_def,hide_index=True,width="stretch",disabled=["Line"],column_config=cfg,key=f"e3b_{s3sig}")
    st.markdown("**Step 3 — GSTR-1 (filed)** *(optional)* — type the *Total Liability* line of the GSTR-1 summary (net of notes & amendments); leave 0 to skip")
    eg1=st.data_editor(zero_tbl(ROWS_G1),hide_index=True,width="stretch",disabled=["Line"],column_config=cfg,key="eg1")
    with st.expander("References (auto-filled — edit if needed)",expanded=bool(srcs)):
        ebk=st.data_editor(tbk_def,hide_index=True,width="stretch",disabled=["Line"],column_config=cfg,key=f"ebk_{ssig}")
        e2b=st.data_editor(t2b_def,hide_index=True,width="stretch",disabled=["Line"],column_config=cfg,key=f"e2b_{ssig}")
        eib=st.data_editor(tib_def,hide_index=True,width="stretch",disabled=["Line"],column_config=cfg,key=f"eib_{ssig}")
        eei=st.data_editor(tei_def,hide_index=True,width="stretch",disabled=["Line"],column_config=cfg,key=f"eei_{ssig}")
    tol3=st.number_input("Tolerance (₹)",min_value=0.0,value=2.0,step=1.0,key="tb3_tol")
    use_2b=e2b if (pt3 is not None or e2b[COLS3].abs().sum().sum()>0) else None
    if st.button("▶ Run 3B Check",type="primary",width="stretch",key="run_3b"):
        try:
            use_bk=ebk if ebk[COLS3].abs().sum().sum()>0 else None; use_ib=eib if eib[COLS3].abs().sum().sum()>0 else None
            use_ei=eei if eei[COLS3].abs().sum().sum()>0 else None; use_g1=eg1 if eg1[COLS3].abs().sum().sum()>0 else None
            df3,util=run_3b_v2(e3b,use_g1,use_bk,use_2b,use_ib,tol3,use_ei,pay3,nongst3)
            st.session_state["tb3_out"]=(df3,util,build_3b_excel_v2(df3,util,e3b,use_g1,use_bk,use_2b,use_ib,client,period,use_ei).getvalue(),pay3)
            st.session_state["tb3_run"]=_stamp(e3b,eg1,ebk,e2b,eib,eei,tol3)
        except Exception as e:
            st.session_state.pop("tb3_out",None); st.error(f"❌ {e}")
            with st.expander("Details"): import traceback; st.code(traceback.format_exc())
    if st.session_state.get("tb3_out"):
        df3,util,xb,pay_=st.session_state["tb3_out"]; bad=df3[~df3.ok]
        if st.session_state.get("tb3_run")!=_stamp(e3b,eg1,ebk,e2b,eib,eei,tol3): st.warning("Inputs changed after the last run - press Run again to refresh these results.")
        if use_2b is None: st.info("No GSTR-2B supplied - the ITC-vs-2B and RCM-vs-2B checks were skipped (upload the 2B above to run them).")
        if len(bad): st.warning(f"⚠️ {len(bad)} difference(s) found out of {len(df3)} checks")
        else: st.success("✅ 3B agrees with all references")
        d=df3[["Section","Description","3B","Reference","Ref source","Diff","Status","Remark"]]
        st.dataframe(d,width="stretch",hide_index=True,height=420,column_config={k:st.column_config.NumberColumn(format="₹ %,.2f") for k in ["3B","Reference","Diff"]})
        if pay_:
            A_=pay_["A"]; B_=pay_["B"]
            st.caption(f"As filed (Table 6.1): ITC used IGST ₹{sum(A_[r][3] for r in range(3)):,.0f} · CGST ₹{sum(A_[r][4] for r in range(3)):,.0f} · SGST ₹{sum(A_[r][5] for r in range(3)):,.0f} | cash on outward ₹{sum(A_[r][7] for r in range(3)):,.0f} | RCM paid in cash ₹{sum(B_[r][7] for r in range(3)):,.0f}")
        else:
            st.caption(f"Indicative only - opening credit-ledger balance not considered. Cash if only this month's ITC were available (IGST/CGST/SGST): ₹{util['cash'][0]:,.0f} / ₹{util['cash'][1]:,.0f} / ₹{util['cash'][2]:,.0f}")
        st.download_button("⬇️ Download 3B Check Excel",data=xb,file_name=f"{client.replace(' ','_')}_3B_Check_{period.replace(' ','_')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",width="stretch")
    with st.expander("📋 Where to find the values"):
        st.markdown("""
**GSTR-3B** — Returns → GSTR-3B → Download filed PDF (read automatically).
**GSTR-1** — Returns → GSTR-1 → View summary: *Total Liability* row, Table 8 nil/exempt, exports rows.
**2B** — upload the Excel above (or run the ITC tab first). **Books sales** — run the Sales tab first. **ITC reference** — run the ITC tab first.
        """)

st.markdown("---")
st.markdown('<div style="text-align:center;font-size:.74rem;color:#888">GST Audit Reconciliation Tool · Built by <a href="https://linkedin.com/in/purvadoshi26" style="color:#5585B5">Purva Doshi</a> · build 2026-11-19 r5 · always-on Run button</div>',unsafe_allow_html=True)
