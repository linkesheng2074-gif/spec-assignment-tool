import tempfile
from pathlib import Path

import streamlit as st

st.set_page_config(page_title="规格书赋值分析工具", layout="wide")

try:
    from spec_engine import run_analysis
except Exception as e:
    st.error("spec_engine.py 导入失败，请检查代码或依赖。")
    st.exception(e)
    st.stop()

st.title("规格书赋值分析工具")
st.caption("上传模板、赋值标准、目标规格、仿真值和 Lot 测试数据，自动生成规格赋值分析结果。")

with st.sidebar:
    st.header("使用说明")
    st.write("1. 上传 template.xlsx")
    st.write("2. 上传 assign_standard.xlsx")
    st.write("3. 上传 target_spec.xlsx")
    st.write("4. 上传 sim_value.xlsx")
    st.write("5. 上传一个或多个 Lot 测试数据")
    st.write("6. 点击开始分析并下载结果")

st.subheader("1. 上传基础文件")

template_file = st.file_uploader("上传 template.xlsx", type=["xlsx"], key="template")
assign_file = st.file_uploader("上传 assign_standard.xlsx", type=["xlsx"], key="assign")
target_file = st.file_uploader("上传 target_spec.xlsx", type=["xlsx"], key="target")
sim_file = st.file_uploader("上传 sim_value.xlsx", type=["xlsx"], key="sim")

st.subheader("2. 上传 Lot 测试数据")

lot_files = st.file_uploader(
    "上传一个或多个 Lot 测试数据 Excel",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
    key="lots",
)

ready = template_file and assign_file and target_file and sim_file and lot_files

if not ready:
    st.info("请先上传所有必要文件。")

if ready and st.button("开始分析", type="primary"):
    with st.spinner("正在分析，请稍等..."):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir = Path(tmpdir)

                template_path = tmpdir / "template.xlsx"
                assign_path = tmpdir / "assign_standard.xlsx"
                target_path = tmpdir / "target_spec.xlsx"
                sim_path = tmpdir / "sim_value.xlsx"

                lot_dir = tmpdir / "lots"
                output_dir = tmpdir / "output"
                output_path = output_dir / "spec_assignment_output.xlsx"

                lot_dir.mkdir(parents=True, exist_ok=True)
                output_dir.mkdir(parents=True, exist_ok=True)

                template_path.write_bytes(template_file.getbuffer())
                assign_path.write_bytes(assign_file.getbuffer())
                target_path.write_bytes(target_file.getbuffer())
                sim_path.write_bytes(sim_file.getbuffer())

                for uploaded_file in lot_files:
                    lot_path = lot_dir / uploaded_file.name
                    lot_path.write_bytes(uploaded_file.getbuffer())

                result = run_analysis(
                    lot_dir=str(lot_dir),
                    template_file=str(template_path),
                    assign_file=str(assign_path),
                    target_file=str(target_path),
                    sim_file=str(sim_path),
                    output_file=str(output_path),
                )

                st.success("分析完成")

                summary_df = result.get("summary_df")
                need_assign_df = result.get("need_assign_df")
                need_target_df = result.get("need_target_df")
                need_sim_df = result.get("need_sim_df")

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
            st.error("分析失败，请检查上传文件格式或脚本逻辑。")
            st.exception(e)
