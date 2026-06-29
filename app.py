import tempfile
from pathlib import Path

import streamlit as st

from spec_engine import run_analysis


st.set_page_config(
    page_title="规格书赋值分析工具",
    page_icon="📊",
    layout="wide"
)

st.title("📊 规格书赋值分析工具")
st.caption("支持多 Lot 测试数据汇总、规格赋值、目标规格风险判断、仿真值对比。")


with st.sidebar:
    st.header("使用说明")
    st.markdown(
        """
        1. 上传汇总模板 `template.xlsx`  
        2. 上传赋值标准 `assign_standard.xlsx`  
        3. 上传目标规格 `target_spec.xlsx`  
        4. 上传仿真值 `sim_value.xlsx`  
        5. 上传一个或多个 Lot 测试数据  
        6. 点击开始分析  
        7. 下载输出结果 Excel  
        """
    )

    st.warning("真实客户测试数据、规格、仿真值不建议上传到公网平台，正式使用建议部署在公司内网。")


st.subheader("1. 上传基础文件")

col1, col2 = st.columns(2)

with col1:
    template_file = st.file_uploader(
        "上传汇总模板 template.xlsx",
        type=["xlsx"],
        key="template"
    )

    assign_file = st.file_uploader(
        "上传赋值标准 assign_standard.xlsx",
        type=["xlsx"],
        key="assign"
    )

with col2:
    target_file = st.file_uploader(
        "上传目标规格 target_spec.xlsx",
        type=["xlsx"],
        key="target"
    )

    sim_file = st.file_uploader(
        "上传仿真值 sim_value.xlsx",
        type=["xlsx"],
        key="sim"
    )


st.subheader("2. 上传 Lot 测试数据")

lot_files = st.file_uploader(
    "上传一个或多个 Lot 测试数据 Excel",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
    key="lots"
)


st.subheader("3. 开始分析")

ready = (
    template_file is not None
    and assign_file is not None
    and target_file is not None
    and sim_file is not None
    and lot_files is not None
    and len(lot_files) > 0
)

if not ready:
    st.info("请先上传 template / assign_standard / target_spec / sim_value / Lot 测试数据。")


if st.button("开始分析", type="primary", disabled=not ready):
    with st.spinner("正在分析，请稍候..."):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir = Path(tmpdir)

                lot_dir = tmpdir / "lots"
                output_dir = tmpdir / "output"

                lot_dir.mkdir(parents=True, exist_ok=True)
                output_dir.mkdir(parents=True, exist_ok=True)

                template_path = tmpdir / "template.xlsx"
                assign_path = tmpdir / "assign_standard.xlsx"
                target_path = tmpdir / "target_spec.xlsx"
                sim_path = tmpdir / "sim_value.xlsx"
                output_path = output_dir / "spec_assignment_output.xlsx"

                # 保存上传的基础文件
                template_path.write_bytes(template_file.getbuffer())
                assign_path.write_bytes(assign_file.getbuffer())
                target_path.write_bytes(target_file.getbuffer())
                sim_path.write_bytes(sim_file.getbuffer())

                # 保存上传的 Lot 文件
                for lot_file in lot_files:
                    lot_path = lot_dir / lot_file.name
                    lot_path.write_bytes(lot_file.getbuffer())

                # 调用核心分析代码
                result = run_analysis(
                    lot_dir=str(lot_dir),
                    template_file=str(template_path),
                    assign_file=str(assign_path),
                    target_file=str(target_path),
                    sim_file=str(sim_path),
                    output_file=str(output_path)
                )

                st.success("分析完成！")

                summary_df = result["summary_df"]
                need_assign_df = result["need_assign_df"]
                need_target_df = result["need_target_df"]
                need_sim_df = result["need_sim_df"]

                c1, c2, c3, c4 = st.columns(4)

                c1.metric("Summary 参数行数", len(summary_df))
                c2.metric("需补赋值规则", len(need_assign_df))
                c3.metric("需补目标规格", len(need_target_df))
                c4.metric("需补仿真值", len(need_sim_df))

                st.subheader("结果预览")
                st.dataframe(summary_df.head(100), use_container_width=True)

                st.subheader("检查项")

                tab1, tab2, tab3 = st.tabs([
                    "Need_Check_Assign",
                    "Need_Check_Target",
                    "Need_Check_Sim"
                ])

                with tab1:
                    st.dataframe(need_assign_df, use_container_width=True)

                with tab2:
                    st.dataframe(need_target_df, use_container_width=True)

                with tab3:
                    st.dataframe(need_sim_df, use_container_width=True)

                with open(output_path, "rb") as f:
                    output_bytes = f.read()

                st.download_button(
                    label="下载分析结果 Excel",
                    data=output_bytes,
                    file_name="spec_assignment_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error("运行失败，请检查上传文件格式、Sheet 名称或参数命名规则。")
            st.exception(e)
