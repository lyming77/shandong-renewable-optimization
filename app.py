# -*- coding: utf-8 -*-
"""
山东电网新能源发电能力优化 —— 场站影响分析网页
用法: streamlit run app.py
"""
import hashlib
import io

import numpy as np
import openpyxl
import pandas as pd
import streamlit as st

from model import analyze

st.set_page_config(page_title="新能源发电能力优化影响分析", layout="wide")
st.title("山东电网新能源发电能力优化 —— 场站影响分析")

st.caption(
    "上传与 `测算数据.xlsx` 同结构的文件（每个场站一个 Sheet，15 分钟点，"
    "列为 年/月/日/时刻/实时电价/出清计划/超短期预测/日前预测/实际出力/可用容量），"
    "页面录入各场站额定容量后一键测算 7 个试运行场景。"
)

# ---------------- 数据上传 ----------------
uploaded = st.file_uploader("上传测算数据 xlsx", type=["xlsx"])

if uploaded is not None:
    raw = uploaded.getvalue()

    # 文件变化时清掉上次结果，避免残留
    fkey = hashlib.md5(raw).hexdigest()
    if st.session_state.get("_filekey") != fkey:
        st.session_state.pop("results", None)
        st.session_state["_filekey"] = fkey

    # 先列出 Sheet 名（独立打开，避免后续 BytesIO 被读空）
    try:
        wb_list = openpyxl.load_workbook(io.BytesIO(raw), read_only=True)
        sheet_names = list(wb_list.sheetnames)
        wb_list.close()
    except Exception as e:
        st.error(f"无法读取文件：{e}")
        st.stop()

    if not sheet_names:
        st.error("文件中没有任何 Sheet。")
        st.stop()

    st.subheader("1. 场站额定容量录入")
    st.caption("按 Sheet 名自动匹配内置场站映射并预填；可修改显示名与额定容量。容量为空或 ≤0 的行将被跳过。")

    # 预填默认值
    rows = []
    for name in sheet_names:
        if name in analyze.STATIONS:
            real_name, rated = analyze.STATIONS[name]
        else:
            real_name, rated = name, None
        rows.append({"Sheet名": name, "场站显示名": real_name, "额定容量MW": rated})

    df_in = pd.DataFrame(rows)
    df_edit = st.data_editor(
        df_in,
        num_rows="fixed",
        disabled=["Sheet名"],
        hide_index=True,
        key="capacity_editor",
    )

    if st.button("开始分析", type="primary"):
        # 组装 capacities
        capacities = {}
        skipped = []
        for _, row in df_edit.iterrows():
            sheet = str(row["Sheet名"]).strip()
            disp = str(row["场站显示名"]).strip() or sheet
            rated = row["额定容量MW"]
            try:
                rated_f = float(rated)
            except (TypeError, ValueError):
                rated_f = np.nan
            if np.isnan(rated_f) or rated_f <= 0:
                skipped.append(sheet)
                continue
            capacities[sheet] = (disp, rated_f)

        if not capacities:
            st.warning("没有录入有效的额定容量，无法分析。")
            st.stop()

        with st.spinner("测算中…"):
            try:
                res = analyze.run_analysis(io.BytesIO(raw), capacities)
            except Exception as e:
                st.error(f"分析失败：{e}")
                st.stop()

        rows_out, event_rows, recov_rows, adj_rows, full_rows = res
        st.session_state["results"] = dict(
            df_sum=pd.DataFrame(rows_out),
            df_ev=pd.DataFrame(event_rows),
            df_rec=pd.DataFrame(recov_rows),
            df_full=pd.DataFrame(full_rows),
            rows_out=rows_out,
            event_rows=event_rows,
            recov_rows=recov_rows,
            adj_rows=adj_rows,
            full_rows=full_rows,
            skipped=skipped,
        )

    # ---- 结果渲染（结果已缓存到 session_state，切换选择器不会丢）----
    if "results" in st.session_state:
        R = st.session_state["results"]
        df_sum = R["df_sum"]
        df_ev = R["df_ev"]
        df_rec = R["df_rec"]
        df_full = R["df_full"]
        rows_out = R["rows_out"]
        event_rows = R["event_rows"]
        recov_rows = R["recov_rows"]
        adj_rows = R["adj_rows"]
        full_rows = R["full_rows"]

        if R["skipped"]:
            st.info(f"以下 Sheet 因容量无效被跳过：{', '.join(R['skipped'])}")

        st.subheader("2. 汇总对比（场站 × 场景）")
        st.dataframe(df_sum, width="stretch", hide_index=True)

        # ---- 收益矩阵 / 压减电量矩阵 ----
        st.subheader("3. 收益与压减电量矩阵")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**净收益影响（元，正=净少赚）**")
            pivot_rev = df_sum.pivot_table(
                index="实际场站", columns="场景", values="净收益影响元", aggfunc="sum"
            )
            st.dataframe(pivot_rev, width="stretch")
        with c2:
            st.markdown("**压减电量（MWh）**")
            pivot_ene = df_sum.pivot_table(
                index="实际场站", columns="场景", values="压减电量MWh", aggfunc="sum"
            )
            st.dataframe(pivot_ene, width="stretch")

        # ---- 分组柱状图 ----
        st.markdown("**各场站 × 场景 净收益影响 / 压减电量**")
        chart_rev = df_sum.pivot_table(index="场景", columns="实际场站", values="净收益影响元", aggfunc="sum")
        chart_ene = df_sum.pivot_table(index="场景", columns="实际场站", values="压减电量MWh", aggfunc="sum")
        cc1, cc2 = st.columns(2)
        with cc1:
            st.bar_chart(chart_rev)
        with cc2:
            st.bar_chart(chart_ene)

        # ---- 触发 / 恢复明细 ----
        st.subheader("4. 触发 / 恢复事件明细")
        tc1, tc2 = st.columns(2)
        with tc1:
            st.markdown("**触发事件**")
            st.dataframe(df_ev, width="stretch", hide_index=True)
        with tc2:
            st.markdown("**恢复事件**")
            st.dataframe(df_rec, width="stretch", hide_index=True)

        # ---- 时序图 ----
        st.subheader("5. 单场站单场景时序")
        if not df_full.empty:
            sel1, sel2 = st.columns(2)
            with sel1:
                stn_sel = st.selectbox("场站", sorted(df_full["实际场站"].unique()), key="ts_stn")
            with sel2:
                sc_sel = st.selectbox("场景", sorted(df_full["场景"].unique()), key="ts_sc")
            sub = df_full[(df_full["实际场站"] == stn_sel) & (df_full["场景"] == sc_sel)].copy()
            sub["时间"] = pd.to_datetime(sub["时间"], errors="coerce")
            sub = sub.dropna(subset=["时间"]).sort_values("时间").reset_index(drop=True)
            if not sub.empty:
                # ---- 该场站该场景汇总指标卡 ----
                srow = df_sum[(df_sum["实际场站"] == stn_sel) & (df_sum["场景"] == sc_sel)]
                if not srow.empty:
                    srow = srow.iloc[0]
                    m1, m2, m3, m4, m5, m6 = st.columns(6)
                    m1.metric("触发次数", int(srow["触发次数"]))
                    m2.metric("恢复次数", int(srow["恢复次数"]))
                    m3.metric("压减周期数", int(srow["压减周期数"]))
                    m4.metric("压减电量 MWh", f"{srow['压减电量MWh']:,.2f}")
                    m5.metric("净收益影响 元", f"{srow['净收益影响元']:,.0f}")
                    m6.metric("压减出力峰值 MW", f"{srow['压减出力峰值MW']:,.1f}")

                # ---- 出力对比（含触发/恢复标记）----
                st.markdown("**出力对比**")
                plot_cols = ["原出清计划MW", "调整后计划MW", "之前实际出力MW", "压减后实际出力MW"]
                st.line_chart(sub.set_index("时间")[plot_cols])

                # 触发/恢复标记（散点，与出力曲线同轴）
                mark_rows = []
                for _, rr in sub.iterrows():
                    if str(rr["触发标记"]):
                        mark_rows.append({"时间": rr["时间"], "标记": "触发", "点位MW": float(rr["调整后计划MW"])})
                    if str(rr["恢复标记"]) == "恢复":
                        mark_rows.append({"时间": rr["时间"], "标记": "恢复", "点位MW": float(rr["调整后计划MW"])})
                if mark_rows:
                    mark_df = pd.DataFrame(mark_rows)
                    st.markdown("**触发/恢复标记**（红=触发，绿=恢复）")
                    st.scatter_chart(mark_df, x="时间", y="点位MW", color="标记")

                # ---- 压减量 + ΔP% ----
                st.markdown("**压减量与调整幅度**")
                p1, p2 = st.columns(2)
                with p1:
                    st.markdown("压减量 (MW)")
                    st.bar_chart(sub.set_index("时间")[["压减量MW"]])
                with p2:
                    st.markdown("调整幅度 ΔP%")
                    st.line_chart(sub.set_index("时间")[["ΔP%"]])
                st.caption("压减量 = max(0, 之前实际出力 − 压减后计划)，即实际出力被真正削减的功率。")

                # ---- 实时电价 ----
                st.markdown("**实时电价（元/MWh）**")
                st.line_chart(sub.set_index("时间")[["实时电价元/MWh"]])
                st.caption("正价时段少发=少赚(损失)；负价时段少发=少亏(避免支出)。")

                # ---- 该场站该场景的触发/恢复事件 ----
                st.markdown("**本组合的触发 / 恢复事件**")
                ev_sub = df_ev[(df_ev["实际场站"] == stn_sel) & (df_ev["场景"] == sc_sel)]
                rc_sub = df_rec[(df_rec["实际场站"] == stn_sel) & (df_rec["场景"] == sc_sel)]
                e1, e2 = st.columns(2)
                with e1:
                    st.dataframe(ev_sub, width="stretch", hide_index=True, height=220)
                with e2:
                    st.dataframe(rc_sub, width="stretch", hide_index=True, height=220)
            else:
                st.info("该组合没有明细数据。")

        # ---- 下载 ----
        st.subheader("6. 下载结果")
        try:
            wb = analyze.build_workbook(rows_out, event_rows, recov_rows, adj_rows, full_rows)
            buf = io.BytesIO()
            wb.save(buf)
            xlsx_bytes = buf.getvalue()
        except Exception as e:
            st.error(f"生成 Excel 失败：{e}")
            xlsx_bytes = None

        csv_text = analyze.build_csv(full_rows)
        dl1, dl2 = st.columns(2)
        with dl1:
            if xlsx_bytes:
                st.download_button(
                    "下载 汇总结果.xlsx", data=xlsx_bytes,
                    file_name="汇总结果.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        with dl2:
            st.download_button(
                "下载 全量明细.csv", data=csv_text.encode("utf-8-sig"),
                file_name="调整后计划_全量明细.csv", mime="text/csv",
            )
else:
    st.info("请上传 xlsx 文件开始。")
