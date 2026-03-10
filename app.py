# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO

class ShipmentDataChecker:
    """发货数据核对Agent"""
    
    def __init__(self):
        self.country_mapping = {
            'us': '美国',
            'ca': '加拿大',
            'uk': '英国',
            'eu': '德国'
        }
        
        self.results = {
            'matched': [],
            'error': [],
            'not_found': [],
            'skipped': []
        }
    
    def normalize_country(self, country_code):
        if pd.isna(country_code):
            return None
        country_code = str(country_code).lower().strip()
        return self.country_mapping.get(country_code, country_code)
    
    def extract_custom_shipment(self, value):
        if pd.isna(value):
            return None
        value_str = str(value).strip()
        match = re.search(r'(\d+)', value_str)
        if match:
            return float(match.group(1))
        try:
            return float(value_str)
        except ValueError:
            return None
    
    def fill_product_name(self, df):
        """
        填充产品名称，处理合并单元格的情况
        如果当前行为空，则向上查找第一个非空值
        """
        product_names = []
        current_product = None
        
        for idx, row in df.iterrows():
            product_val = row.get('产品')
            if pd.notna(product_val) and str(product_val).strip():
                current_product = str(product_val).strip()
                product_names.append(current_product)
            else:
                # 如果当前行为空，使用上一个非空值
                product_names.append(current_product)
        
        return product_names
    
    def build_file1_index(self, df1):
        # 先处理产品名称（解决合并单元格问题）
        filled_products = self.fill_product_name(df1)
        
        index = {}
        for idx, row in df1.iterrows():
            if pd.isna(row.get('ASIN')) or pd.isna(row.get('标签（FNSKU)')):
                continue
                
            asin = str(row['ASIN']).strip() if not pd.isna(row['ASIN']) else ""
            fnsku = str(row['标签（FNSKU)']).strip() if not pd.isna(row['标签（FNSKU)']) else ""
            country = self.normalize_country(row.get('国家'))
            
            # 使用填充后的产品名称
            product_name = filled_products[idx] if idx < len(filled_products) else ""
            
            if not country or not asin or not fnsku:
                continue
                
            key = (asin, fnsku, country)
            custom_shipment = self.extract_custom_shipment(row.get('自定义发货'))
            
            # 如果当前key已存在，保留第一个非空的产品名称
            if key in index:
                existing_product = index[key]['product_name']
                if not existing_product and product_name:
                    index[key]['product_name'] = product_name
            else:
                index[key] = {
                    'row_index': idx,
                    'custom_shipment': custom_shipment,
                    'product_name': product_name,
                }
        return index
    
    def check(self, df1, df2, tolerance=80):
        self.results = {'matched': [], 'error': [], 'not_found': [], 'skipped': []}
        file1_index = self.build_file1_index(df1)
        
        for idx, row in df2.iterrows():
            asin = str(row['ASIN']).strip() if not pd.isna(row['ASIN']) else ""
            fnsku = str(row['FNSKU']).strip() if not pd.isna(row['FNSKU']) else ""
            country = str(row['国家']).strip() if not pd.isna(row['国家']) else ""
            
            # 处理计划发货量，保留为整数
            planned_qty = float(row['计划发货量']) if not pd.isna(row['计划发货量']) else 0
            planned_qty = int(round(planned_qty))
            
            key = (asin, fnsku, country)
            
            if key in file1_index:
                file1_record = file1_index[key]
                custom_qty = file1_record['custom_shipment']
                product_name = file1_record['product_name'] if file1_record['product_name'] else "未知产品"
                
                if custom_qty is None:
                    continue
                
                # 转换为整数
                custom_qty = int(round(custom_qty))
                error = abs(planned_qty - custom_qty)
                
                if error > tolerance:
                    self.results['error'].append({
                        '产品名称': product_name,
                        'ASIN': asin, 
                        'FNSKU': fnsku, 
                        '国家': country,
                        '补货建议数据': custom_qty,      # 改名为补货建议数据
                        'ERP数据': planned_qty,         # 改名为ERP数据
                        '误差': int(error)
                    })
        
        return self.results

# Streamlit页面配置
st.set_page_config(
    page_title="误差超标核对",
    page_icon="⚠️",
    layout="wide"
)

# 自定义CSS样式
st.markdown("""
<style>
    .upload-box {
        border: 2px dashed #ccc;
        padding: 20px;
        border-radius: 10px;
        margin: 10px 0;
    }
    .error-table {
        margin-top: 20px;
    }
    .dataframe {
        font-size: 14px;
    }
</style>
""", unsafe_allow_html=True)

# 创建两列上传区域
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📁 补货建议")
    file1 = st.file_uploader("选择111补货建议.xlsx", type=['xlsx', 'xls'], key="file1")

with col2:
    st.markdown("### 📁 发货计划 (ERP数据)")
    file2 = st.file_uploader("选择发货计划.xlsx", type=['xlsx', 'xls'], key="file2")

# 设置误差容忍度
st.markdown("---")
col_slider1, col_slider2 = st.columns([3, 1])
with col_slider1:
    tolerance = st.slider("误差容忍度", min_value=0, max_value=200, value=80, step=10)
with col_slider2:
    st.metric("当前值", tolerance)

# 核对按钮
if file1 and file2:
    if st.button("🔍 开始核对", type="primary", use_container_width=True):
        with st.spinner("正在核对数据..."):
            try:
                df1 = pd.read_excel(file1, sheet_name='Sheet1')
                df2 = pd.read_excel(file2, sheet_name='sheet1')
                
                # 显示读取状态
                st.caption(f"已读取补货建议: {len(df1)}行, ERP数据: {len(df2)}行")
                
                checker = ShipmentDataChecker()
                results = checker.check(df1, df2, tolerance)
                
                st.markdown("---")
                
                # 只显示误差超标的结果
                if results['error']:
                    st.warning(f"⚠️ 发现 {len(results['error'])} 条误差超标记录")
                    
                    # 创建DataFrame
                    df_error = pd.DataFrame(results['error'])
                    
                    # 重新排列列顺序
                    column_order = ['产品名称', 'ASIN', 'FNSKU', '国家', '补货建议数据', 'ERP数据', '误差']
                    df_error = df_error[column_order]
                    
                    # 显示表格
                    st.dataframe(
                        df_error.style.applymap(
                            lambda x: 'background-color: #ffcccc' if isinstance(x, (int, float)) and x > tolerance else '',
                            subset=['误差']
                        ),
                        use_container_width=True,
                        height=400
                    )
                    
                    # 导出功能
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_error.to_excel(writer, sheet_name='误差超标记录', index=False)
                        
                        # 添加统计信息
                        stats_df = pd.DataFrame([{
                            '总超标数': len(results['error']),
                            '误差容忍度': tolerance
                        }])
                        stats_df.to_excel(writer, sheet_name='统计信息', index=False)
                    
                    output.seek(0)
                    
                    st.download_button(
                        label="📥 下载超标记录",
                        data=output,
                        file_name="误差超标记录.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                else:
                    st.success(f"✅ 没有发现误差超过 {tolerance} 的记录")
                    
            except Exception as e:
                st.error(f"核对出错: {str(e)}")
                import traceback
                with st.expander("查看详细错误"):
                    st.code(traceback.format_exc())

else:
    st.info("👆 请上传两个Excel文件")

# 简短的说明
with st.expander("📖 说明"):
    st.markdown("""
    **文件格式要求：**
    - **补货建议**：需包含`产品`、`ASIN`、`标签（FNSKU)`、`国家`(us/ca/uk/eu)、`自定义发货`列
    - **ERP数据**：需包含`ASIN`、`FNSKU`、`国家`(美国/加拿大/英国/德国)、`计划发货量`列
    
    **输出内容说明：**
    | 列名 | 说明 |
    |------|------|
    | 产品名称 | 从补货建议中获取的产品名称 |
    | ASIN | 商品标准识别码 |
    | FNSKU | FNSKU编码 |
    | 国家 | 销售国家 |
    | **补货建议数据** | 从补货建议文件中读取的自定义发货量 |
    | **ERP数据** | 从发货计划文件中读取的计划发货量 |
    | 误差 | 两者的差值 |
    
    **功能特点：**
    - ✅ 自动处理合并单元格，产品名称会向下填充
    - ✅ 所有数量值显示为整数
    - ✅ 只显示误差超过容忍度的记录
    - ✅ 红色高亮显示超标行
    """)
