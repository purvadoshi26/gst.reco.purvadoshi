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
import gst_core as _core
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
def _fsig(*keys):
    """signature of the files CURRENTLY uploaded under these widget keys (None if any is missing)"""
    h=hashlib.md5()
    for k in keys:
        f=st.session_state.get(k)
        if f is None: return None
        h.update(f.getvalue())
    return h.hexdigest()[:10]
def _stamp(*objs):
    return hashlib.md5("|".join(repr(getattr(o,"values",o).tolist() if hasattr(getattr(o,"values",None),"tolist") else o) for o in objs).encode()).hexdigest()[:10]

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
        st.markdown("**File 2b — Debit Note Register** *(optional — if exported as a separate file)*")
        f_dn=st.file_uploader("Debit Note Register (optional)",type=["xlsx","xls"],key="itc_dn",label_visibility="collapsed")
    if f_2b and f_pr:
        eb=f_2b.getvalue(); pbb=f_pr.getvalue(); dnb=f_dn.getvalue() if f_dn else None
        sig=hashlib.md5(pbb+eb+(dnb or b"")).hexdigest()[:10]; scan=None; pt=None
        try:
            if st.session_state.get("itc_sig")!=sig:
                st.session_state["itc_pt"]=parse_2b_full(eb); st.session_state["itc_scan"]=scan_pbooks_any(pbb,dn_fb=dnb)
                st.session_state["itc_sig"]=sig; st.session_state.pop("itc_out",None); st.session_state.pop("itc_run",None)
            scan=st.session_state["itc_scan"]; pt=st.session_state["itc_pt"]
        except Exception as e:
            st.error(f"❌ {e}")
            with st.expander("Details"): import traceback; st.code(traceback.format_exc())
        if scan is not None:
            pivot=(scan["kind"]=="legacy" and (scan["bw"]["_inv"].astype(str).str.strip()=="").all())
            if pt.get("meta"): st.caption("GSTR-2B file: "+_meta_line(pt["meta"]))
            pdates=[d for d in scan["bw"]["_date"] if pd.notna(d) and hasattr(d, "year")]
            if pdates and pt.get("meta",{}).get("pstart"):
                min_pd, max_pd = min(pdates), max(pdates)
                pst = pt["meta"]["pstart"]
                if min_pd.year != pst.year or min_pd.month != pst.month:
                    st.error(f"⛔ TAX PERIOD MISMATCH: Purchase vouchers are dated {min_pd:%d-%b-%Y} to {max_pd:%d-%b-%Y}, but GSTR-2B is for {pst:%b-%Y} ({pt['meta'].get('FY','')})! Please verify that your files are for the same tax period.")
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
                            res=run_itc_v2(pbb,eb,lc,vc,tol,mat,scan=scan,pt=pt,dn_bytes=dnb)
                            st.session_state["itc_out"]=(res,build_itc_excel_v2(res,client,period).getvalue()); st.session_state["itc_run"]=_stamp(led_ed,vt_ed,tol,mat)
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
        st.markdown("**File 2b — Credit Note Register** *(optional — if exported as a separate file)*")
        f_cn=st.file_uploader("Credit Note Register (optional)",type=["xlsx","xls"],key="sal_cn",label_visibility="collapsed")
    if f_einv and f_sal:
        eb=f_einv.getvalue(); bb=f_sal.getvalue(); cb=f_cn.getvalue() if f_cn else None
        sig=hashlib.md5(eb+bb+(cb or b"")).hexdigest()[:10]
        scan=None
        try:
            if st.session_state.get("sal_sig")!=sig:
                st.session_state["sal_scan"]=scan_books(bb,cn_fb=cb); st.session_state["sal_sig"]=sig; st.session_state.pop("sal_out",None); st.session_state.pop("sal_run",None)
            scan=st.session_state["sal_scan"]
        except Exception as e:
            st.error(f"❌ {e}")
            with st.expander("Details"): import traceback; st.code(traceback.format_exc())
        if scan:
            try:
                _pm=parse_portal(eb).get("meta",{})
                if _pm.get("GSTIN") or _pm.get("Period"): st.caption("E-invoice file: "+_meta_line(_pm))
                bdates=[d for d in scan["bw"]["_date"] if pd.notna(d) and hasattr(d, "year")]
                if bdates and _pm.get("pstart"):
                    min_d, max_d = min(bdates), max(bdates)
                    pst = _pm["pstart"]
                    if min_d.year != pst.year or min_d.month != pst.month:
                        st.error(f"⛔ TAX PERIOD MISMATCH: Books vouchers are dated {min_d:%d-%b-%Y} to {max_d:%d-%b-%Y}, but the E-Invoice Portal file is for {pst:%b-%Y} ({_pm.get('FY','')})! Comparing across different financial years will cause false mismatches. Please verify that your Books export matches the portal period.")
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
                        res=run_sales_reco(eb,bb,lc,vc,tol,scan=scan,cn_bytes=cb)
                        buf=build_sales_excel(res,client,period)
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
    itc_valid=st.session_state.get("itc_sig") is not None and st.session_state.get("itc_sig")==_fsig("itc_pr","itc_2b")
    sal_valid=st.session_state.get("sal_sig") is not None and st.session_state.get("sal_sig")==_fsig("sal_inv","sal_bk")
    if (not f_2b3) and itc_valid and st.session_state.get("itc_pt") is not None: pt3=st.session_state["itc_pt"]; src2b="2B from ITC tab ("+str(getattr(st.session_state.get("itc_2b"),"name",""))+")"
    t2b_def=twob_refs(pt3) if pt3 else zero_tbl(ROWS_2B)
    sal=st.session_state.get("sal_out") if sal_valid else None; itcx=st.session_state.get("itc_out") if itc_valid else None
    tbk_def=zero_tbl(ROWS_BK); tib_def=zero_tbl(ROWS_IB); tei_def=zero_tbl(ROWS_EI); srcs=[]
    st.markdown("**Step 1b — E-invoice / GSTR-1 portal Excel** *(optional — gives a B2B outward reference before you type GSTR-1)*")
    f_ei=st.file_uploader("E-invoice Excel",type=["xlsx","xls"],key="tb3_ei",label_visibility="collapsed")
    metas=[]
    if pt3 and pt3.get("meta"): metas.append(("2B",pt3["meta"]))
    if f_ei:
        try: pe_=parse_portal(f_ei.getvalue()); tei_def=einv_net(pe_); srcs.append("e-invoice reference ← uploaded file"); metas.append(("E-invoice",pe_.get("meta",{})))
        except Exception as e: st.warning(f"Could not read e-invoice file: {e}")
    elif sal: tei_def=einv_net(sal[0]["portal"]); metas.append(("E-invoice",sal[0]["portal"].get("meta",{}))); srcs.append("e-invoice reference ← Sales Reco tab ("+str(getattr(st.session_state.get("sal_inv"),"name",""))+")")
    if sal:
        bs=books_sales_net(sal[0]); 
        for c in COLS3: tbk_def.loc[0,c]=bs[c]
        srcs.append("books sales ← Sales Reco tab ("+str(getattr(st.session_state.get("sal_bk"),"name",""))+")")
    if itcx:
        e,rc=itc_reco_refs(itcx[0])
        for k,c in enumerate(COLS3[1:]): tib_def.loc[0,c]=e[k]; tib_def.loc[1,c]=rc[k]
        srcs.append("ITC reference ← ITC Reco tab ("+str(getattr(st.session_state.get("itc_pr"),"name",""))+")")
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
st.markdown('<div style="text-align:center;font-size:.74rem;color:#888">GST Audit Reconciliation Tool · Built by <a href="https://linkedin.com/in/purvadoshi26" style="color:#5585B5">Purva Doshi</a> · build 2026-10-18 r2</div>',unsafe_allow_html=True)
