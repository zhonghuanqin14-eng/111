# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import re
import requests
from io import BytesIO
import time

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

def load_kdocs_file(url):
    """
    从金山文档URL加载Excel文件 - 修复版
    使用模拟浏览器行为下载文件
    """
    try:
        # 从URL中提取文档ID
        doc_id = url.split('/')[-1]
        
        # 设置更完整的浏览器请求头
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'none',
            'Sec-Fetch-User': '?1',
            'Cache-Control': 'max-age=0',
        }
        
        st.info("🔄 正在访问金山文档...")
        
        # 第一步：访问分享链接，获取重定向后的真实地址
        session = requests.Session()
        response = session.get(url, headers=headers, timeout=30, allow_redirects=True)
        
        if response.status_code != 200:
            st.error(f"访问失败，状态码: {response.status_code}")
            return None
        
        # 第二步：尝试找到下载按钮的链接
        # 金山文档通常有导出功能，我们可以尝试构造导出链接
        export_url = f"https://www.kdocs.cn/api/office/file/{doc_id}/export?type=xlsx"
        
        st.info("🔄 正在导出Excel文件...")
        export_response = session.get(export_url, headers=headers, timeout=30)
        
        if export_response.status_code == 200:
            # 尝试直接读取
            try:
                df = pd.read_excel(BytesIO(export_response.content), engine='openpyxl')
                st.success("✅ Excel文件读取成功！")
                return df
            except Exception as e:
                st.warning(f"直接读取失败，尝试备用方法: {e}")
        
        # 第三步：备用方法 - 尝试获取预览数据
        st.info("🔄 尝试获取预览数据...")
        preview_url = f"https://www.kdocs.cn/api/office/file/{doc_id}/preview"
        preview_response = session.get(preview_url, headers=headers, timeout=30)
        
        if preview_response.status_code == 200:
            try:
                data = preview_response.json()
                st.write("调试信息：", data)  # 这行可以帮助我们了解返回的数据结构
            except:
                pass
        
        # 如果以上都失败，提示用户手动下载
        st.warning("⚠️ 自动读取失败，请使用手动下载方式")
        st.markdown(f"""
        **手动下载步骤：**
        1. 打开链接：{url}
        2. 点击右上角的「下载」按钮
        3. 下载文件后，使用下面的上传功能
        """)
        
        return None
            
    except requests.exceptions.RequestException as e:
        st.error(f"网络请求错误: {e}")
        return None
    except Exception as e:
        st.error(f"处理失败: {e}")
        return None

# Streamlit页面配置
st.set_page_config(
    page_title="误差超标核对",
    page_icon="⚠️",
    layout="wide"
)

# 自定义CSS样式
st.markdown("""
<style>
    .info-box {
        background-color: #e7f3ff;
        padding: 15px;
        border-radius: 5px;
        border-left: 4px solid #2196F3;
        margin: 10px 0;
    }
    .warning-box {
        background-color: #fff3cd;
        padding: 15px;
        border-radius: 5px;
        border-left: 4px solid #ffc107;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# 标题
st.markdown("# ⚠️ 误差超标核对系统")

# 显示金山文档链接信息
st.markdown("""
<div class="info-box">
    <strong>📌 附件1（补货建议）来源：</strong> 金山文档在线表格<br>
    <strong>链接：</strong> https://www.kdocs.cn/l/caxFVgMMlMXG
</div>
""", unsafe_allow_html=True)

# 创建选项卡
tab1, tab2 = st.tabs(["🌐 在线读取（自动）", "📁 本地上传（手动）"])

with tab1:
    st.markdown("### 🌐 从金山文档自动读取")
    
    # 显示链接确认
    st.markdown(f"**当前链接:** `https://www.kdocs.cn/l/caxFVgMMlMXG`")
    
    # 手动触发读取按钮
    if st.button("🔄 读取在线文档", type="primary", use_container_width=True):
        with st.spinner("正在连接金山文档..."):
            df1 = load_kdocs_file("https://www.kdocs.cn/l/caxFVgMMlMXG")
            if df1 is not None:
                st.session_state['df1'] = df1
                st.success(f"✅ 读取成功！共 {len(df1)} 行数据")
                # 显示前几行预览
                preview_cols = ['产品', 'ASIN', '标签（FNSKU)', '国家', '自定义发货']
                available_cols = [col for col in preview_cols if col in df1.columns]
                st.dataframe(df1[available_cols].head(5), use_container_width=True)
            else:
                st.markdown("""
                <div class="warning-box">
                    <strong>💡 提示：</strong> 如果自动读取失败，请切换到「本地上传」选项卡，手动下载后上传。
                </div>
                """, unsafe_allow_html=True)

with tab2:
    st.markdown("### 📁 手动上传附件1")
    st.info("如果在线读取失败，请先手动下载金山文档，然后在这里上传")
    file1 = st.file_uploader("选择补货建议.xlsx", type=['xlsx', 'xls'], key="file1_manual")
    if file1:
        df1 = pd.read_excel(file1, sheet_name='Sheet1')
        st.session_state['df1'] = df1
        st.success(f"✅ 上传成功！共 {len(df1)} 行数据")

# 附件2上传区域（放在外面，两个选项卡共享）
st.markdown("---")
st.markdown("### 📁 附件2：发货计划")
file2 = st.file_uploader("选择发货计划.xlsx", type=['xlsx', 'xls'], key="file2")

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
    st.info("👆 请先加载附件1（可以使用在线读取或手动上传）")
else:
    st.info("👆 请上传附件2（发货计划）")

# 使用说明
with st.expander("📖 使用说明"):
    st.markdown("""
    ### 两种方式加载附件1：
    
    **1. 在线读取（自动）**
    - 点击「读取在线文档」按钮
    - 系统会自动从金山文档获取数据
    - 如果失败，请使用手动上传
    
    **2. 本地上传（手动）**
    - 打开金山文档链接
    - 点击右上角「下载」按钮
    - 将下载的文件上传
    
    **功能特点：**
    - ✅ 自动处理合并单元格，产品名称会向下填充
    - ✅ 所有数量值显示为整数
    - ✅ 只显示误差超过容忍度的记录
    
    **输出列：** 产品名称 | ASIN | FNSKU | 国家 | 自定义发货量 | 计划发货量 | 误差
    """)
