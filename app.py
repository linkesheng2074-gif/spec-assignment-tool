import tempfile
import traceback
from pathlib import Path

import streamlit as st

st.set_page_config(page_title="规格书赋值分析工具", layout="wide")

st.title("规格书赋值分析工具")
st.success("页面已启动。如果能看到这行，说明 Streamlit 页面本身正常。")
st.caption("V11 安全版：页面启动时不导入分析引擎，点击开始分析后才导入 spec_engine.py，避免白屏。")

with st.sidebar:
    st.header("使用说明")
    st.write("1. 上传汇总模板 template.xlsx")
    st.write("2. 上传赋值标准 assign_standard.xlsx")
    st.write("3. 上传目标规格 target_spec.xlsx")
    st.write("4. 上传仿真值 sim_value.xlsx")
    st.write("5. 上传一个或多个 Lot 测试数据")
    st.write("6. 点击开始分析并下载结果")

st.subheader("1. 上传基础文件")
col1, col2 = st.columns(2)

with col1:
    template_file = st.file_uploader("上传汇总模板（template.xlsx）", type=["xlsx"], key="template")
    target_file = st.file_uploader("上传目标规格（target_spec.xlsx）", type=["xlsx"], key="target")

with col2:
    assign_file = st.file_uploader("上传赋值标准（assign_standard.xlsx）", type=["xlsx"], key="assign")
    sim_file = st.file_uploader("上传仿真值（sim_value.xlsx）", type=["xlsx"], key="sim")

st.subheader("2. 上传 Lot 测试数据")
lot_files = st.file_uploader(
    "上传一个或多个 Lot 测试数据 Excel（支持 .xlsx / .xlsm）",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
    key="lots",
)

ready = bool(template_file and assign_file and target_file and sim_file and lot_files)

if not ready:
    st.info("请先上传所有必要文件。")
else:
    total_size_mb = sum(f.size for f in lot_files) / 1024 / 1024
    st.success(f"文件已上传：Lot 文件 {len(lot_files)} 个，合计约 {total_size_mb:.1f} MB。")

if ready and st.button("开始分析", type="primary"):
    progress = st.progress(0, text="准备开始分析...")
    status = st.empty()

    try:
        status.write("1/9 正在导入分析引擎...")
        progress.progress(0.05, text="正在导入分析引擎...")

        try:
            import spec_engine
        except Exception as e:
            st.error("spec_engine.py 导入失败。")
            st.exception(e)
            st.stop()

        status.write("✅ 分析引擎导入完成")

        with tempfile.TemporaryDirectory() as td:
            tmpdir = Path(td)
            lot_dir = tmpdir / "lots"
            output_dir = tmpdir / "output"
            lot_dir.mkdir(parents=True, exist_ok=True)
            output_dir.mkdir(parents=True, exist_ok=True)

            template_path = tmpdir / "template.xlsx"
            assign_path = tmpdir / "assign_standard.xlsx"
            target_path = tmpdir / "target_spec.xlsx"
            sim_path = tmpdir / "sim_value.xlsx"
            output_path = output_dir / "spec_assignment_output.xlsx"

            template_path.write_bytes(template_file.getbuffer())
            assign_path.write_bytes(assign_file.getbuffer())
            target_path.write_bytes(target_file.getbuffer())
            sim_path.write_bytes(sim_file.getbuffer())

            for uploaded_file in lot_files:
                (lot_dir / uploaded_file.name).write_bytes(uploaded_file.getbuffer())

            steps = [
                ("读取汇总模板", spec_engine.read_template, (str(template_path),)),
                ("读取赋值标准", spec_engine.read_assign_standard, (str(assign_path),)),
                ("读取目标规格", spec_engine.read_target_spec, (str(target_path),)),
                ("读取仿真值", spec_engine.read_sim_value, (str(sim_path),)),
                ("读取 Lot 测试数据", spec_engine.read_all_lots, (str(lot_dir),)),
            ]

            results = []
            for idx, (name, func, args) in enumerate(steps, start=2):
                status.write(f"{idx}/9 {name}...")
                progress.progress(idx / 10, text=name)
                results.append(func(*args))
                status.write(f"✅ {name}完成")

            template_df, assign_df, target_df, sim_df, raw_df = results

            status.write("7/9 生成 Summary...")
            progress.progress(0.70, text="生成 Summary...")
            summary_df = spec_engine.build_summary(template_df, raw_df, assign_df, target_df, sim_df)

            status.write("8/9 生成 Raw_Pivot 和检查表...")
            progress.progress(0.82, text="生成 Raw_Pivot 和检查表...")
            raw_pivot_df = spec_engine.build_raw_pivot(raw_df)
            need_assign_df = spec_engine.build_need_check_assign(summary_df)
            need_target_df = spec_engine.build_need_check_target(summary_df)
            need_sim_df = spec_engine.build_need_check_sim(summary_df)

            status.write("9/9 导出 Excel...")
            progress.progress(0.92, text="导出 Excel...")
            spec_engine.export_excel(
                summary_df,
                raw_df,
                raw_pivot_df,
                need_assign_df,
                need_target_df,
                need_sim_df,
                str(output_path),
            )

            progress.progress(1.0, text="分析完成")
            status.write("✅ 分析完成")
            st.success("分析完成")

            tab1, tab2, tab3, tab4 = st.tabs([
                "Summary 预览",
                "Need_Check_Assign",
                "Need_Check_Target",
                "Need_Check_Sim",
            ])

            with tab1:
                st.dataframe(summary_df.head(100), use_container_width=True)
            with tab2:
                st.dataframe(need_assign_df, use_container_width=True)
            with tab3:
                st.dataframe(need_target_df, use_container_width=True)
            with tab4:
                st.dataframe(need_sim_df, use_container_width=True)

            with open(output_path, "rb") as f:
                st.download_button(
                    label="下载分析结果 Excel",
                    data=f,
                    file_name="spec_assignment_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

    except Exception as e:
        progress.empty()
        st.error("分析失败，请检查上传文件格式或脚本逻辑。")
        st.exception(e)
        with st.expander("查看完整错误堆栈", expanded=True):
            st.code(traceback.format_exc())
