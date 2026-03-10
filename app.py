# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import re
import requests
from io import BytesIO
from urllib.parse import urlparse

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
            'error': []
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
        self.results = {'error': []}
        file1_index = self.build_file1_index(df1)
        
        for idx, row in df2.iterrows():
            asin = str(row['ASIN']).strip() if not pd.isna(row['ASIN']) else ""
            fnsku = str(row['FNSKU']).strip() if not pd.isna(row['FNSKU']) else ""
            country = str(row['国家']).strip() if not pd.isna(row['国家']) else ""
            
            planned_qty = float(row['计划发货量']) if not pd.isna(row['计划发货量']) else 0
            planned_qty = int(round(planned_qty))
            
            key = (asin, fnsku, country)
            
            if key in file1_index:
                file1_record = file1_index[key]
                custom_qty = file1_record['custom_shipment']
                product_name = file1_record['product_name'] if file1_record['product_name'] else "未知产品"
                
                if custom_qty is None:
                    continue
                
                custom_qty = int(round(custom_qty))
                error = abs(planned_qty - custom_qty)
                
                if error > tolerance:
                    self.results['error'].append({
                        '产品名称': product_name,
                        'ASIN': asin, 
                        'FNSKU': fnsku, 
                        '国家': country,
                        '自定义发货量': custom_qty,
                        '计划发货量': planned_qty, 
                        '误差': int(error)
                    })
        
        return self.results

def convert_kdocs_url_to_direct(url):
    """
    将金山文档的分享URL转换为可以直接下载Excel的URL
    金山文档的分享URL格式：https://www.kdocs.cn/l/xxx
    我们需要提取文档ID并构造下载链接
    """
    try:
        # 从URL中提取文档ID
        # https://www.kdocs.cn/l/caxFVgMMlMXG  -> 文档ID是 caxFVgMMlMXG
        doc_id = url.split('/')[-1]
        
        # 金山文档的下载接口
        # 这是通过分析金山文档分享机制得到的下载链接格式
        direct_url = f"https://www.kdocs.cn/api/office/file/{doc_id}/download"
        
        return direct_url
    except Exception as e:
        st.error(f"URL解析错误: {e}")
        return None

def load_kdocs_file(url):
    """
    从金山文档URL加载Excel文件
    """
    try:
        # 添加请求头模拟浏览器访问
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Accept-Language': 'zh-CN,zh;q=0.9',
        }
        
        # 转换URL并下载
        download_url = convert_kdocs_url_to_direct(url)
        if not download_url:
            return None
            
        st.info(f"正在从金山文档读取数据...")
        response = requests.get(download_url, headers=headers, timeout=30)
        
        if response.status_code == 200:
            # 从内存中读取Excel
            excel_data = BytesIO(response.content)
            df = pd.read_excel(excel_data, sheet_name='Sheet1')
            return df
        else:
            st.error(f"下载失败，状态码: {response.status_code}")
            return None
            
    except requests.exceptions.RequestException as e:
        st.error(f"网络请求错误: {e}")
        return None
    except Exception as e:
        st.error(f"读取Excel失败: {e}")
        return None

# Streamlit页面配置
st.set_page_config(
    page_title="误差超标核对（在线版）",
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
    .info-box {
        background-color: #e7f3ff;
        padding: 15px;
        border-radius: 5px;
        border-left: 4px solid #2196F3;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# 标题
st.markdown("# ⚠️ 误差超标核对系统（在线版）")

# 显示金山文档链接信息
st.markdown("""
<div class="info-box">
    <strong>📌 附件1（补货建议）来源：</strong> 金山文档在线表格<br>
    <strong>链接：</strong> https://www.kdocs.cn/l/caxFVgMMlMXG
</div>
""", unsafe_allow_html=True)

# 创建两列布局
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📁 附件1：在线补货建议")
    st.markdown("**状态：** 已配置在线文档")
    
    # 手动触发读取按钮
    if st.button("🔄 重新读取在线文档", use_container_width=True):
        with st.spinner("正在读取金山文档..."):
            df1 = load_kdocs_file("https://www.kdocs.cn/l/caxFVgMMlMXG")
            if df1 is not None:
                st.session_state['df1'] = df1
                st.success(f"✅ 读取成功！共 {len(df1)} 行数据")
                # 显示前几行预览
                preview_cols = ['产品', 'ASIN', '标签（FNSKU)', '国家', '自定义发货']
                available_cols = [col for col in preview_cols if col in df1.columns]
                st.dataframe(df1[available_cols].head(3), use_container_width=True)
            else:
                st.error("❌ 读取失败，请检查链接是否有效")

with col2:
    st.markdown("### 📁 附件2：发货计划")
    file2 = st.file_uploader("选择发货计划.xlsx", type=['xlsx', 'xls'], key="file2")
    if file2:
        st.success(f"✅ 已上传: {file2.name}")

# 设置误差容忍度
st.markdown("---")
col_slider1, col_slider2 = st.columns([3, 1])
with col_slider1:
    tolerance = st.slider("误差容忍度", min_value=0, max_value=200, value=80, step=10)
with col_slider2:
    st.metric("当前值", tolerance)

# 核对按钮
if 'df1' in st.session_state and file2:
    if st.button("🔍 开始核对", type="primary", use_container_width=True):
        with st.spinner("正在核对数据..."):
            try:
                df1 = st.session_state['df1']
                df2 = pd.read_excel(file2, sheet_name='sheet1')
                
                checker = ShipmentDataChecker()
                results = checker.check(df1, df2, tolerance)
                
                st.markdown("---")
                
                # 只显示误差超标的结果
                if results['error']:
                    st.warning(f"⚠️ 发现 {len(results['error'])} 条误差超标记录")
                    
                    # 创建DataFrame
                    df_error = pd.DataFrame(results['error'])
                    
                    # 重新排列列顺序
                    column_order = ['产品名称', 'ASIN', 'FNSKU', '国家', '自定义发货量', '计划发货量', '误差']
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

elif 'df1' not in st.session_state:
    st.info("👆 请先点击「重新读取在线文档」按钮加载附件1")

else:
    st.info("👆 请上传附件2（发货计划）")

# 使用说明
with st.expander("📖 说明"):
    st.markdown("""
    **功能特点：**
    - ✅ 自动从金山文档读取附件1（补货建议）
    - ✅ 只需要上传附件2（发货计划）
    - ✅ 自动处理合并单元格，产品名称会向下填充
    - ✅ 所有数量值显示为整数
    - ✅ 只显示误差超过容忍度的记录
    
    **输出列：** 产品名称 | ASIN | FNSKU | 国家 | 自定义发货量 | 计划发货量 | 误差
    
    **如果读取失败，可以尝试：**
    1. 确认链接是否公开可访问
    2. 手动下载文件后使用旧版上传功能
    """)
