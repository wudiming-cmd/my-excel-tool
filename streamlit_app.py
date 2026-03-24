import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="表格自动填充工具", layout="wide")

st.title("🚀 表格数据对应填充工具")
st.markdown("---")

# 1. 文件上传区
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 上传【数据源表】")
    st.info("这个表里有你想提取的信息（如：包含所有人的电话、地址）")
    source_file = st.file_uploader("选择源 Excel", type=['xlsx'], key="source")

with col2:
    st.subheader("2. 上传【目标待填表】")
    st.info("这个表是你现在要填写的空表")
    target_file = st.file_uploader("选择目标 Excel", type=['xlsx'], key="target")

# 2. 处理逻辑
if source_file and target_file:
    df_source = pd.read_excel(source_file)
    df_target = pd.read_excel(target_file)
    
    st.success("文件读取成功！")
    st.markdown("---")
    
    # 3. 参数配置区
    st.subheader("3. 配置对齐规则")
    
    c1, c2 = st.columns(2)
    with c1:
        source_key = st.selectbox("源表里的匹配列（如：身份证/姓名）", df_source.columns)
    with c2:
        target_key = st.selectbox("目标表里的匹配列（必须与左侧对应）", df_target.columns)
    
    selected_columns = st.multiselect("你想从【源表】提取哪些列填入【目标表】？", 
                                     [c for c in df_source.columns if c != source_key])

    if st.button("开始合并数据"):
        if not selected_columns:
            st.warning("请至少选择一个要填充的列")
        else:
            # 核心对齐逻辑：左连接
            # 只取源表中的 key 和 用户选中的列
            df_to_merge = df_source[[source_key] + selected_columns]
            
            # 执行合并
            result_df = pd.merge(df_target, df_to_merge, 
                                 left_on=target_key, 
                                 right_on=source_key, 
                                 how='left')
            
            # 如果两个表的 key 列名不一样，合并后会多出一列，这里把它删掉
            if source_key != target_key:
                result_df = result_df.drop(columns=[source_key])

            st.write("### 预览合并后的结果 (前5行)：")
            st.dataframe(result_df.head())

            # 4. 下载区
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False)
            
            st.download_button(
                label="✅ 点击下载合并后的 Excel",
                data=output.getvalue(),
                file_name="合并结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.warning("请先在上方上传两个 Excel 文件。")
