# -*- coding: utf-8 -*-
"""
账单拆分底稿生成器（网页版 / Streamlit）

上传：若干 Italent 账单 PDF + 一份 HC 人员映射表 xlsx（可再传一份别名表 json）
产出：标准拆分底稿 xlsx，可直接下载

运行：streamlit run app.py
"""
from __future__ import annotations

import hashlib
import io
import json
import os
import re
import sys

import pandas as pd
import streamlit as st

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from 匹配人员 import S_OK, load_mapping, norm, parse_aliases       # noqa: E402
from 解析发票 import parse_invoice                                   # noqa: E402
from 生成拆分底稿 import build, guess_period                         # noqa: E402
from 输出底稿 import write_workbook                                  # noqa: E402

SHEET_TITLE = "账单拆分底稿生成器"

st.set_page_config(page_title=SHEET_TITLE, page_icon="📑", layout="wide",
                   initial_sidebar_state="expanded")


# ================================================================ 计算层（带缓存）

@st.cache_data(show_spinner=False)
def _parse_pdf(data: bytes, filename: str):
    """解析一张账单 PDF。返回 Invoice。"""
    return parse_invoice(io.BytesIO(data), filename=filename)


@st.cache_data(show_spinner=False)
def _load_mapping(data: bytes):
    """读取人员映射表。返回 Person 列表。"""
    return load_mapping(io.BytesIO(data))


# ================================================================ 视图层小工具

def df_detail(rows) -> pd.DataFrame:
    rec = [{"账单姓名": r.bill_name, "匹配姓名": r.matched_name or "—",
            "项目(Project)": r.project, "部门成本中心": r.dept,
            "结算金额": round(r.amount, 2), "匹配状态": r.status,
            "匹配方式": r.method, "备注": r.note} for r in rows]
    d = pd.DataFrame(rec)
    d.loc[len(d)] = {"账单姓名": "合计", "匹配姓名": "", "项目(Project)": "",
                     "部门成本中心": "", "结算金额": round(sum(r.amount for r in rows), 2),
                     "匹配状态": "", "匹配方式": "", "备注": ""}
    return d


def df_overview(sheets) -> pd.DataFrame:
    rec = []
    for sd in sheets:
        inv = sd.invoice
        total = inv.total if inv.total is not None else 0.0
        diff = round(inv.amount_sum_raw - total, 3)
        rec.append({
            "账单文件": inv.filename,
            "Invoice Number": inv.number,
            "明细累加合计": round(sd.total, 2),
            "Invoice Total(账单底部)": total,
            "差额": diff,
            "核对结果": "✔ 一致" if abs(diff) < 0.01 else "✖ 不一致",
            "本账单出现的Activity类型": inv.activity_summary(),
        })
    return pd.DataFrame(rec)


def df_agg(rows, keyfunc, label):
    d = {}
    for r in rows:
        d[keyfunc(r)] = d.get(keyfunc(r), 0.0) + r.amount
    total = sum(d.values()) or 1.0
    rec = [{label: k, "结算金额合计": round(v, 2), "占比": v / total}
           for k, v in sorted(d.items(), key=lambda kv: (-kv[1], kv[0]))]
    rec.append({label: "合计", "结算金额合计": round(sum(d.values()), 2), "占比": 1.0})
    return pd.DataFrame(rec)


# ================================================================ 侧边栏

st.sidebar.title("📑 " + SHEET_TITLE)
st.sidebar.caption("Italent 账单 PDF × HC 人员映射表 → 标准拆分底稿")

# 文件夹模式：设置环境变量 SPLIT_LOCAL_DIR 后无需上传，直接读本地目录（自建/演示用）
LOCAL_DIR = os.environ.get("SPLIT_LOCAL_DIR", "").strip()


def _read_dir(folder: str):
    """从目录读取账单 PDF / 映射表 / 别名表。"""
    pdfs, xlsx, alias = [], None, b""
    for fn in sorted(os.listdir(folder)):
        p = os.path.join(folder, fn)
        if not os.path.isfile(p):
            continue
        low = fn.lower()
        if low.endswith(".pdf"):
            with open(p, "rb") as fh:
                pdfs.append((fn, fh.read()))
        elif low.endswith((".xlsx", ".xlsm")) and not fn.startswith("~$"):
            with open(p, "rb") as fh:
                data = fh.read()
            if xlsx is None or re.search(r"hc|report|内部", fn, re.I):
                xlsx = (fn, data)
        elif low.endswith(".json"):
            with open(p, "rb") as fh:
                alias = fh.read()
    return pdfs, xlsx, alias


if LOCAL_DIR:
    st.sidebar.info(f"📂 文件夹模式：`{LOCAL_DIR}`")
    period_in = st.sidebar.text_input("期号（留空自动推断）", value="")
    _pdf_items, _map_item, _alias_bytes = _read_dir(LOCAL_DIR)
    mapping_name = _map_item[0] if _map_item else ""
else:
    pdf_files = st.sidebar.file_uploader(
        "① 账单 PDF（可多选）", type=["pdf"], accept_multiple_files=True,
        help="Italent Workforce 出的账单，文件名一般形如 INVOICE#1059.pdf")
    mapping_file = st.sidebar.file_uploader(
        "② 人员映射表（HC Report）", type=["xlsx", "xlsm"],
        help="需要含 Full Name / Department Cost Center 表头")
    alias_file = st.sidebar.file_uploader(
        "③ 别名映射.json（可选）", type=["json"],
        help="上传上次下载的别名表，已确认过的人员不会再提示待确认")
    period_in = st.sidebar.text_input("期号（留空自动推断）", value="", placeholder="如 202608")

    _pdf_items = [(f.name, f.getvalue()) for f in (pdf_files or [])]
    _map_item = (mapping_file.name, mapping_file.getvalue()) if mapping_file else None
    _alias_bytes = alias_file.getvalue() if alias_file else b""
    mapping_name = _map_item[0] if _map_item else ""

st.sidebar.divider()
if st.sidebar.button("🔄 清空缓存并重新计算", width="stretch"):
    st.cache_data.clear()
    st.rerun()
st.sidebar.caption("作者：WorkBuddy　｜　本地也有一份同样的命令行工具")


# 别名表：随上传文件变化而重置（网页里手动确认的结果存在 session_state）
_alias_key = hashlib.md5(_alias_bytes).hexdigest()
if st.session_state.get("alias_key") != _alias_key:
    st.session_state["alias_key"] = _alias_key
    st.session_state["aliases"] = parse_aliases(_alias_bytes) if _alias_bytes else {}


# ================================================================ 首页引导

st.title("📑 " + SHEET_TITLE)
st.caption("上传账单 PDF 和人员映射表，自动按人拆分、匹配项目与部门成本中心，"
           "核对账单总额，输出标准「拆分底稿」xlsx。")

if not _pdf_items or not _map_item:
    st.info("**使用方法**\n\n"
            "1. 左侧 ① 上传当期账单 PDF（可一次多张）\n"
            "2. 左侧 ② 上传 HC 人员映射表\n"
            "3. 需要的话 ③ 上传上次下载的 `别名映射.json`\n\n"
            "传完自动开始计算。")
    c1, c2, c3 = st.columns(3)
    c1.markdown("#### 1️⃣ 拆分\n按账单姓名把明细金额汇总到人，"
                "自动匹配映射表里的项目与部门成本中心。")
    c2.markdown("#### 2️⃣ 核对\n明细累加合计与账单底部 Total 自动对差，"
                "不一致会红字标出。")
    c3.markdown("#### 3️⃣ 汇总\n按部门成本中心 / 项目 / 账单三个维度出金额与占比，"
                "外加项目×部门明细矩阵。")
    st.stop()


# ================================================================ 计算

st.caption(f"人员映射表：`{mapping_name}`　｜　账单：{len(_pdf_items)} 张")
try:
    people = _load_mapping(_map_item[1])
except Exception as e:
    st.error(f"人员映射表读取失败：{e}")
    st.stop()

invoices, failed = [], []
for name, data in _pdf_items:
    try:
        invoices.append(_parse_pdf(data, name))
    except Exception as e:
        failed.append((name, str(e)))

if failed:
    for name, err in failed:
        st.warning(f"**{name}** 解析失败，已跳过：{err}")
if not invoices:
    st.error("没有解析成功的账单 PDF，请确认上传的是 Italent 账单。")
    st.stop()

aliases = st.session_state.get("aliases", {})
sheets = build(invoices, people, aliases)
all_rows = [r for sd in sheets for r in sd.rows]

period = period_in.strip() or guess_period(invoices) or "未定期号"
out_name = f"CM-Italent-工具拆分底稿-{period}.xlsx"

buf = io.BytesIO()
write_workbook(sheets, buf)
xlsx_bytes = buf.getvalue()

# 别名表用「原始账单姓名」更好读：反查一次
_raw = {}
for r in all_rows:
    if norm(r.bill_name) in aliases and r.status == S_OK:
        _raw[r.bill_name] = r.matched_name
alias_pretty = json.dumps(_raw, ensure_ascii=False, indent=2)


# ================================================================ 概览

n_inv = len(invoices)
n_people = len(all_rows)
n_amt = round(sum(r.amount for r in all_rows), 2)
n_bad = len([r for r in all_rows if r.status != S_OK])
n_diff = len([sd for sd in sheets
              if abs(sd.invoice.amount_sum_raw - (sd.invoice.total or 0)) >= 0.01])

m = st.columns(5)
m[0].metric("账单张数", n_inv)
m[1].metric("明细条数", sum(len(i.items) for i in invoices))
m[2].metric("涉及人数", n_people)
m[3].metric("结算金额合计", f"{n_amt:,.2f}")
m[4].metric("待确认人员", n_bad, delta=None if n_bad == 0 else "需处理", delta_color="inverse")

if n_diff:
    st.error(f"⚠ 有 {n_diff} 张账单的明细累加合计与账单底部 Total 不一致，请先核对。")
elif n_bad == 0:
    st.success("✔ 全部账单金额核对一致，全部人员匹配完成。")
else:
    st.success(f"✔ 全部 {n_inv} 张账单金额核对一致。")

st.download_button(
    f"⬇️ 下载拆分底稿（{out_name}）", data=xlsx_bytes, file_name=out_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary", width="stretch")

st.divider()


# ================================================================ 总览表

st.subheader("① 账单核对总览")
ov = df_overview(sheets)
st.dataframe(ov.style.format({"明细累加合计": "{:,.2f}",
                              "Invoice Total(账单底部)": "{:,.3f}",
                              "差额": "{:,.2f}"}),
             width="stretch", hide_index=True)


# ================================================================ 待确认人员

todo = [r for r in all_rows if r.status != S_OK]
if todo:
    st.divider()
    st.subheader("② 待确认人员（选好后点应用，会写进别名表并自动重算）")
    st.caption("选择结果只影响本次页面与下载文件；别忘了在页面底部把 "
               "`别名映射.json` 下载下来，下次上传它就不用再确认了。")

    # 同一账单姓名在多张账单里重复出现时只问一次
    uniq = {}
    for r in todo:
        uniq.setdefault(r.bill_name, r)

    people_opts = [None] + list(range(len(people)))

    def _fmt(i):
        if i is None:
            return "— 保持未确认 / 确实匹配不上 —"
        p = people[i]
        return f"{p.display}　｜　{p.project}　｜　{p.dept}"

    picks = {}
    with st.form("确认表单"):
        for name, r in uniq.items():
            default = 0
            if r.matched_name:
                for i, p in enumerate(people):
                    if p.display == r.matched_name:
                        default = i + 1
                        break
            c1, c2 = st.columns([1, 2])
            with c1:
                st.markdown(f"**{name}**  \n"
                            f"<span style='color:#C00;font-size:0.85em'>{r.status}</span>",
                            unsafe_allow_html=True)
                if r.note:
                    st.caption(r.note)
            with c2:
                picks[name] = st.selectbox(
                    "应为", people_opts, index=default, format_func=_fmt,
                    key=f"pick::{name}", label_visibility="collapsed")
        applied = st.form_submit_button("✅ 应用这些确认结果", type="primary")

    if applied:
        changed = 0
        for name, idx in picks.items():
            if idx is None:
                continue
            aliases[norm(name)] = people[idx].display
            changed += 1
        st.session_state["aliases"] = aliases
        st.toast(f"已写入 {changed} 条别名，正在重新计算…")
        st.rerun()
else:
    st.divider()
    st.success("② 没有待确认人员，全部匹配完成。")


# ================================================================ 每张账单明细

st.divider()
st.subheader("③ 每张账单拆分明细")
for i, sd in enumerate(sheets, 1):
    inv = sd.invoice
    diff = round(inv.amount_sum_raw - (inv.total or 0), 3)
    tag = "✔" if abs(diff) < 0.01 else "✖ 差额 " + f"{diff:,.2f}"
    with st.expander(f"{i:02d}　{inv.filename}　｜　{inv.number}　｜　"
                     f"{len(sd.rows)} 人　｜　合计 {sd.total:,.2f}　｜　{tag}",
                     expanded=(i == 1)):
        st.caption(f"开票日 {inv.invoice_date or '—'}　｜　"
                   f"出单日 {inv.shipped_date or '—'}　｜　"
                   f"明细 {len(inv.items)} 条　｜　{inv.activity_summary()}")
        st.dataframe(df_detail(sd.rows).style.format({"结算金额": "{:,.2f}"}),
                     width="stretch", hide_index=True)


# ================================================================ 汇总

st.divider()
st.subheader("④ 汇总（跨全部账单）")
t1, t2, t3 = st.tabs(["按部门成本中心", "按项目", "项目 × 部门矩阵"])
with t1:
    st.dataframe(df_agg(all_rows, lambda x: x.dept, "部门成本中心")
                 .style.format({"结算金额合计": "{:,.2f}", "占比": "{:.1%}"}),
                 width="stretch", hide_index=True)
with t2:
    st.dataframe(df_agg(all_rows, lambda x: x.project, "项目(Project)")
                 .style.format({"结算金额合计": "{:,.2f}", "占比": "{:.1%}"}),
                 width="stretch", hide_index=True)
with t3:
    mat = df_agg(all_rows, lambda x: f"{x.project}｜{x.dept}", "项目 ｜ 部门成本中心")
    st.dataframe(mat.style.format({"结算金额合计": "{:,.2f}", "占比": "{:.1%}"}),
                 width="stretch", hide_index=True)


# ================================================================ 匹配日志 & 别名下载

st.divider()
with st.expander("⑤ 匹配日志（核对匹配依据）"):
    log = pd.DataFrame([{"账单文件": sd.invoice.filename, "账单姓名": r.bill_name,
                         "匹配姓名": r.matched_name or "—", "状态": r.status,
                         "匹配方式": r.method, "备注": r.note}
                        for sd in sheets for r in sd.rows])
    st.dataframe(log, width="stretch", hide_index=True)

st.download_button("⬇️ 下载 别名映射.json（下次上传它即可跳过重复确认）",
                   data=alias_pretty.encode("utf-8"),
                   file_name="别名映射.json", mime="application/json",
                   width="stretch")
st.caption("底稿输出目录：`输出/CM-Italent-工具拆分底稿-<期号>.xlsx`　｜　"
           "解析规则：Italent 账单明细表 + HC Report 映射表")
