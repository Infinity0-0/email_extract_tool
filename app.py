# -*- coding: utf-8 -*-
"""
PDF EMAIL EXTRACTOR - Web Tool
Chalane ke liye:   streamlit run app.py

Ye file email_extractor.py ko use karti hai, dono ek hi folder me rakho.
"""

import io
import os
import re
import tempfile
from collections import OrderedDict
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import streamlit as st

import email_extractor as ee

URL_RE = re.compile(r"https?://[^\s<>\"'\[\],;]+", re.I)


def parse_links(text):
    """Kaise bhi paste karo (quotes, commas, alag lines) """
    urls = [u.rstrip(").,") for u in URL_RE.findall(text or "")]
    return list(OrderedDict.fromkeys(urls))


def run_job(links, workers, on_progress):
    """Saare links process karo. Return: results list (email_extractor wala format)."""
    ee.WORKERS = workers
    os.makedirs(ee.PDF_FOLDER, exist_ok=True)
    session = ee.make_session()
    idx = {l: i for i, l in enumerate(links)}
    results, jobs = [], []

    # ---- Step A: har link se PDF ka address nikalo (0% - 40%)
    done = 0
    with ThreadPoolExecutor(workers) as ex:
        futs = {ex.submit(ee.discover, session, l): l for l in links}
        for f in as_completed(futs):
            link = futs[f]
            try:
                found = f.result()
                if not found:
                    raise ValueError("Is link par PDF nahi mila")
                jobs += [(link, u, t) for u, t in found]
            except Exception as e:
                results.append({"idx": idx[link], "link": link, "pdf_url": "", "title": "",
                                "file": "", "status": "Link failed", "note": str(e),
                                "total_pages": 0, "rows": []})
            done += 1
            on_progress(0.4 * done / len(links), f"Links check ho rahe hain ({done}/{len(links)})")

    # ---- Step B: download + emails (40% - 100%)
    done = 0
    with ThreadPoolExecutor(workers) as ex:
        futs = {ex.submit(ee.download, session, u): (l, u, t) for l, u, t in jobs}
        for f in as_completed(futs):
            link, url, title = futs[f]
            base = {"idx": idx[link], "link": link, "pdf_url": url, "title": title}
            try:
                path = f.result()
                res = ee.process_pdf(path)
                ee.maybe_delete(path)
                results.append({**base, "file": os.path.basename(path), "status": res["status"],
                                "note": res["note"], "total_pages": res["total_pages"],
                                "rows": res["rows"]})
            except Exception as e:
                results.append({**base, "file": "", "status": "Download failed",
                                "note": str(e), "total_pages": 0, "rows": []})
            done += 1
            on_progress(0.4 + 0.6 * done / max(len(jobs), 1),
                        f"PDF padh rahe hain ({done}/{len(jobs)})")
    on_progress(1.0, "Ho gaya!")
    return results


def to_excel(results):
    """Results -> (excel_bytes, {sheet: DataFrame}, total_rows, unique_emails)"""
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp.close()
    old = ee.OUTPUT_FILE
    ee.OUTPUT_FILE = tmp.name
    try:
        total, unique, path = ee.save_excel(results)
        with open(path, "rb") as fh:
            data = fh.read()
    finally:
        ee.OUTPUT_FILE = old
        try:
            os.remove(tmp.name)
        except OSError:
            pass
    sheets = pd.read_excel(io.BytesIO(data), sheet_name=None)
    return data, sheets, total, unique


def ui():
    st.set_page_config(page_title="PDF Email Extractor", page_icon="📧", layout="wide")

    # Optional password (company use): set env var APP_PASSWORD
    pwd = os.environ.get("APP_PASSWORD")
    if pwd:
        if st.text_input("Password", type="password") != pwd:
            st.info("Password daalo.")
            st.stop()

    st.title("📧 PDF Email Extractor")
    st.write("Bahut saare PDF / article links paste karo → saare emails Excel me.")

    with st.sidebar:
        st.header("Settings")
        workers = st.slider("Ek saath kitne git add app.py email_extractor.py
git commit -m "kya badla uska chhota sa message"
git pushdownload", 1, 8, 3,
                            help="Zyada karoge to website block kar sakti hai")
        max_pages = st.number_input("Har PDF ke kitne pages padhne hain (0 = sab)", 0, 500, 0)
        sab = st.checkbox("Incomplted Emails", value=True)
        delete_pdf = st.checkbox("delete after complete", value=True)

    text = st.text_area(
        "Links yaha paste karo (quotes, commas, alag lines - kuch bhi chalega)",
        height=260,
        placeholder='"https://www.jidmis.org/index.php/jidmis/article/view/4043/1792",\n'
                    '"https://www.jidmis.org/index.php/jidmis/article/view/4044/1793",',
    )
    links = parse_links(text)
    st.caption(f"{len(links)} links mile")

    if st.button("🚀 Start", type="primary", disabled=not links):
        ee.MAX_PAGES = int(max_pages)
        ee.SAB_KUCH_DO = bool(sab)
        ee.DELETE_PDF_AFTER = bool(delete_pdf)
        bar = st.progress(0.0)
        status = st.empty()

        def cb(frac, msg):
            bar.progress(min(max(frac, 0.0), 1.0))
            status.text(msg)

        results = run_job(links, workers, cb)
        data, sheets, total, unique = to_excel(results)
        failed = sum(1 for r in results if r["status"] not in ("OK", "No text"))
        st.session_state["out"] = {"data": data, "sheets": sheets, "total": total,
                                   "unique": unique, "failed": failed, "n": len(results)}

    out = st.session_state.get("out")
    if out:
        c1, c2, c3 = st.columns(3)
        c1.metric("Email rows", out["total"])
        c2.metric("Alag-alag emails", out["unique"])
        c3.metric("Fail hue links", out["failed"])

        st.download_button("⬇️ Excel download karo", out["data"], file_name="emails_output.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        tab1, tab2 = st.tabs(["Emails (sequence me)", "Article-wise"])
        with tab1:
            st.dataframe(out["sheets"]["Emails"], use_container_width=True, hide_index=True)
        with tab2:
            st.dataframe(out["sheets"]["Article_Wise"], use_container_width=True, hide_index=True)


if __name__ == "__main__":
    ui()