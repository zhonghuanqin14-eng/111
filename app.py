# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
import base64
from datetime import datetime

# 页面配置 - 必须在最前面
st.set_page_config(
    page_title="发货数据核对",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 自定义CSS样式 - 美化界面
st.markdown("""
<style>
    /* 全局样式 */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        font-family: 'Inter', sans-serif;
    }
    
    /* 主容器样式 */
    .main-container {
        background: rgba(255, 255, 255, 0.95);
        border-radius: 20px;
        padding: 30px;
        box-shadow: 0 10px 40px rgba(0, 0, 0, 0.1);
        margin: 20px 0;
        backdrop-filter: blur(10px);
    }
    
    /* 标题样式 */
    .title-container {
        text-align: center;
        margin-bottom: 40px;
        padding: 20px;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 15px;
        color: white;
        box-shadow: 0 10px 30px rgba(102, 126, 234, 0.4);
    }
    
    .title-container h1 {
        font-size: 2.8em;
        font-weight: 700;
        margin: 0;
        letter-spacing: 1px;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
    }
    
    .title-container p {
        font-size: 1.2em;
        margin: 10px 0 0;
        opacity: 0.95;
        font-weight: 300;
    }
    
    /* 上传区域样式 */
    .upload-box {
        background: white;
        border-radius: 15px;
        padding: 25px;
        box-shadow: 0 5px 20px rgba(0, 0, 0, 0.05);
        height: 100%;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
        border: 1px solid rgba(0,0,0,0.05);
    }
    
    .upload-box:hover {
        transform: translateY(-5px);
        box-shadow: 0 15px 30px rgba(0, 0, 0, 0.1);
    }
    
    .upload-header {
        display: flex;
        align-items: center;
        margin-bottom: 20px;
        padding-bottom: 15px;
        border-bottom: 2px solid #f0f2f5;
    }
    
    .upload-icon {
        font-size: 2em;
        margin-right: 15px;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        width: 50px;
        height: 50px;
        border-radius: 12px;
        display: flex;
        align-items: center;
        justify-content: center;
        color: white;
    }
    
    .upload-title {
        font-size: 1.3em;
        font-weight: 600;
        color: #1e293b;
        margin: 0;
    }
    
    .upload-subtitle {
        font-size: 0.9em;
        color: #64748b;
        margin: 5px 0 0;
    }
    
    /* 文件信息卡片 */
    .file-info {
        background: #f8fafc;
        border-radius: 12px;
        padding: 15px;
        margin-top: 15px;
        border: 1px solid #e2e8f0;
    }
    
    .file-info-item {
        display: flex;
        align-items: center;
        margin: 8px 0;
    }
    
    .file-info-label {
        font-weight: 600;
        color: #475569;
        width: 80px;
    }
    
    .file-info-value {
        color: #0f172a;
        flex: 1;
    }
    
    .success-badge {
        background: #10b981;
        color: white;
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.85em;
        font-weight: 500;
        display: inline-block;
    }
    
    /* 参数设置区域 */
    .params-container {
        background: white;
        border-radius: 15px;
        padding: 25px;
        margin: 30px 0;
        box-shadow: 0 5px 20px rgba(0, 0, 0, 0.05);
    }
    
    .params-title {
        font-size: 1.3em;
        font-weight: 600;
        color: #1e293b;
        margin-bottom: 20px;
        display: flex;
        align-items: center;
    }
    
    .params-title span {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        width: 35px;
        height: 35px;
        border-radius: 10px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        margin-right: 12px;
    }
    
    /* 按钮样式 */
    .stButton button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-size: 1.2em;
        font-weight: 600;
        padding: 15px 30px;
        border: none;
        border-radius: 12px;
        cursor: pointer;
        transition: all 0.3s ease;
        box-shadow: 0 10px 20px rgba(102, 126, 234, 0.3);
        width: 100%;
        letter-spacing: 1px;
    }
    
    .stButton button:hover {
        transform: translateY(-3px);
        box-shadow: 0 15px 30px rgba(102, 126, 234, 0.4);
    }
    
    /* 结果统计卡片 */
    .stats-container {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 20px;
        margin: 30px 0;
    }
    
    .stat-card {
        background: white;
        border-radius: 15px;
        padding: 25px;
        text-align: center;
        box-shadow: 0 5px 20px rgba(0, 0, 0, 0.05);
        transition: transform 0.3s ease;
    }
    
    .stat-card:hover {
        transform: translateY(-5px);
    }
    
    .stat-card.warning {
        border-left: 4px solid #f59e0b;
    }
    
    .stat-number {
        font-size: 2.5em;
        font-weight: 700;
        margin: 10px 0;
    }
    
    .stat-label {
        font-size: 1em;
        color: #64748b;
        font-weight: 500;
    }
    
    /* 表格样式 */
    .dataframe-container {
        background: white;
        border-radius: 15px;
        padding: 20px;
        box-shadow: 0 5px 20px rgba(0, 0, 0, 0.05);
        margin: 20px 0;
    }
    
    .dataframe-title {
        font-size: 1.3em;
        font-weight: 600;
        color: #1e293b;
        margin-bottom: 20px;
        display: flex;
        align-items: center;
    }
    
    /* 自定义表格样式 */
    .stDataFrame {
        font-family: 'Inter', sans-serif;
    }
    
    .stDataFrame thead tr th {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-weight: 600;
        padding: 12px;
    }
    
    /* 指标卡片 */
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 12px;
        padding: 20px;
        color: white;
        text-align: center;
    }
    
    /* 进度条样式 */
    .stProgress > div > div {
        background: linear-gradient(90deg, #667eea, #764ba2);
    }
    
    /* 分隔线 */
    .custom-divider {
        background: linear-gradient(90deg, transparent, #667eea, transparent);
        height: 2px;
        margin: 30px 0;
    }
    
    /* 页脚样式 */
    .footer {
        text-align: center;
        padding: 20px;
        color: #64748b;
        font-size: 0.9em;
        margin-top: 40px;
    }
    
    /* 动画效果 */
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(20px); }
        to { opacity: 1; transform: translateY(0); }
    }
    
    .fade-in {
        animation: fadeIn 0.6s ease-out;
    }
</style>
""", unsafe_allow_html=True)

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
        """填充产品名称，处理合并单元格的情况"""
        product_names = []
        current_product = None
        
        for idx, row in df.iterrows():
            product_val = row.get('产品')
            if pd.notna(product_val) and str(product_val).strip():
                current_product = str(product_val).strip()
                product_names.append(current_product)
            else:
                product_names.append(current_product)
        
        return product_names
    
    def build_file1_index(self, df1):
        filled_products = self.fill_product_name(df1)
        
        index = {}
        for idx, row in df1.iterrows():
            if pd.isna(row.get('ASIN')) or pd.isna(row.get('标签（FNSKU)')):
                continue
                
            asin = str(row['ASIN']).strip() if not pd.isna(row['ASIN']) else ""
            fnsku = str(row['标签（FNSKU)']).strip() if not pd.isna(row['标签（FNSKU)']) else ""
            country = self.normalize_country(row.get('国家'))
            product_name = filled_products[idx] if idx < len(filled_products) else ""
            
            if not country or not asin or not fnsku:
                continue
                
            key = (asin, fnsku, country)
            custom_shipment = self.extract_custom_shipment(row.get('自定义发货'))
            
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
            planned_qty = float(row['计划发货量']) if not pd.isna(row['计划发货量']) else 0
            planned_qty = int(round(planned_qty))
            
            key = (asin, fnsku, country)
            
            if key in file1_index:
                file1_record = file1_index[key]
                custom_qty = file1_record['custom_shipment']
                product_name = file1_record['product_name'] if file1_record['product_name'] else "未知产品"
                
                if custom_qty is None:
                    self.results['skipped'].append({
                        '产品名称': product_name,
                        'ASIN': asin, 
                        'FNSKU': fnsku, 
                        '国家': country,
                        '补货建议数据': '无数据',
                        'ERP数据': planned_qty, 
                        '状态': '跳过'
                    })
                    continue
                
                custom_qty = int(round(custom_qty))
                error = abs(planned_qty - custom_qty)
                
                if error > tolerance:
                    self.results['error'].append({
                        '产品名称': product_name,
                        'ASIN': asin, 
                        'FNSKU': fnsku, 
                        '国家': country,
                        '补货建议数据': custom_qty,
                        'ERP数据': planned_qty,
                        '误差': int(error),
                        '差异比例': f"{int(error/planned_qty*100) if planned_qty > 0 else 0}%"
                    })
                else:
                    self.results['matched'].append({
                        '产品名称': product_name,
                        'ASIN': asin, 
                        'FNSKU': fnsku, 
                        '国家': country,
                        '补货建议数据': custom_qty,
                        'ERP数据': planned_qty,
                        '误差': int(error)
                    })
            else:
                self.results['not_found'].append({
                    'ASIN': asin, 
                    'FNSKU': fnsku, 
                    '国家': country, 
                    'ERP数据': planned_qty
                })
        
        return self.results

# 主界面
st.markdown("""
<div class="title-container fade-in">
    <h1>🔍 误差超标核对系统</h1>
    <p>智能核对补货建议与ERP数据，快速定位差异</p>
</div>
""", unsafe_allow_html=True)

# 主要内容容器
st.markdown('<div class="main-container fade-in">', unsafe_allow_html=True)

# 创建两列上传区域
col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="upload-box">
        <div class="upload-header">
            <div class="upload-icon">📋</div>
            <div>
                <div class="upload-title">补货建议文件</div>
                <div class="upload-subtitle">上传111补货建议.xlsx</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    file1 = st.file_uploader("选择文件", type=['xlsx', 'xls'], key="file1", label_visibility="collapsed")
    
    if file1:
        st.markdown(f"""
        <div class="file-info">
            <div class="file-info-item">
                <span class="file-info-label">文件名</span>
                <span class="file-info-value">{file1.name}</span>
            </div>
            <div class="file-info-item">
                <span class="file-info-label">大小</span>
                <span class="file-info-value">{file1.size / 1024:.1f} KB</span>
            </div>
            <div style="text-align: right; margin-top: 10px;">
                <span class="success-badge">✓ 已上传</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="upload-box">
        <div class="upload-header">
            <div class="upload-icon">📊</div>
            <div>
                <div class="upload-title">ERP数据文件</div>
                <div class="upload-subtitle">上传发货计划.xlsx</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    file2 = st.file_uploader("选择文件", type=['xlsx', 'xls'], key="file2", label_visibility="collapsed")
    
    if file2:
        st.markdown(f"""
        <div class="file-info">
            <div class="file-info-item">
                <span class="file-info-label">文件名</span>
                <span class="file-info-value">{file2.name}</span>
            </div>
            <div class="file-info-item">
                <span class="file-info-label">大小</span>
                <span class="file-info-value">{file2.size / 1024:.1f} KB</span>
            </div>
            <div style="text-align: right; margin-top: 10px;">
                <span class="success-badge">✓ 已上传</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

# 参数设置区域
st.markdown("""
<div class="params-container">
    <div class="params-title">
        <span>⚙️</span> 核对参数设置
    </div>
</div>
""", unsafe_allow_html=True)

col_slider1, col_slider2, col_slider3 = st.columns([2, 1, 1])

with col_slider1:
    tolerance = st.slider("误差容忍度", min_value=0, max_value=200, value=80, step=10)

with col_slider2:
    st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
    st.metric("当前值", tolerance, delta=None, delta_color="off")

with col_slider3:
    st.markdown('<div style="margin-top: 28px;"></div>', unsafe_allow_html=True)
    if file1 and file2:
        st.markdown(f"<span class='success-badge'>✓ 就绪</span>", unsafe_allow_html=True)

# 核对按钮
if file1 and file2:
    if st.button("🔍 开始核对", type="primary", use_container_width=True):
        with st.spinner("正在核对数据，请稍候..."):
            try:
                df1 = pd.read_excel(file1, sheet_name='Sheet1')
                df2 = pd.read_excel(file2, sheet_name='sheet1')
                
                # 显示进度
                progress_bar = st.progress(0)
                for i in range(100):
                    if i < 30:
                        progress_bar.progress(i + 1)
                    time.sleep(0.01)
                
                checker = ShipmentDataChecker()
                results = checker.check(df1, df2, tolerance)
                
                progress_bar.progress(100)
                time.sleep(0.5)
                progress_bar.empty()
                
                st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
                
                # 统计卡片
                total_checked = len(results['matched']) + len(results['error'])
                
                st.markdown("""
                <div class="stats-container">
                """, unsafe_allow_html=True)
                
                col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
                
                with col_stat1:
                    st.markdown(f"""
                    <div class="stat-card">
                        <div style="font-size: 2em;">📊</div>
                        <div class="stat-number">{total_checked}</div>
                        <div class="stat-label">总核对数</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col_stat2:
                    st.markdown(f"""
                    <div class="stat-card">
                        <div style="font-size: 2em;">✅</div>
                        <div class="stat-number" style="color: #10b981;">{len(results['matched'])}</div>
                        <div class="stat-label">匹配成功</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col_stat3:
                    st.markdown(f"""
                    <div class="stat-card warning">
                        <div style="font-size: 2em;">⚠️</div>
                        <div class="stat-number" style="color: #f59e0b;">{len(results['error'])}</div>
                        <div class="stat-label">误差超标</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col_stat4:
                    st.markdown(f"""
                    <div class="stat-card">
                        <div style="font-size: 2em;">❓</div>
                        <div class="stat-number" style="color: #64748b;">{len(results['not_found'])}</div>
                        <div class="stat-label">未匹配</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                # 只显示误差超标的结果
                if results['error']:
                    st.markdown("""
                    <div class="dataframe-container">
                        <div class="dataframe-title">
                            <span style="background: #f59e0b; color: white; width: 35px; height: 35px; border-radius: 10px; display: inline-flex; align-items: center; justify-content: center; margin-right: 12px;">⚠️</span>
                            误差超标记录
                        </div>
                    """, unsafe_allow_html=True)
                    
                    # 创建DataFrame
                    df_error = pd.DataFrame(results['error'])
                    
                    # 重新排列列顺序
                    column_order = ['产品名称', 'ASIN', 'FNSKU', '国家', '补货建议数据', 'ERP数据', '误差', '差异比例']
                    df_error = df_error[column_order]
                    
                    # 添加高亮样式
                    def highlight_error(val):
                        if isinstance(val, (int, float)) and val > tolerance:
                            return 'background-color: #ffebee; color: #c62828; font-weight: bold;'
                        return ''
                    
                    # 显示表格
                    st.dataframe(
                        df_error.style.applymap(highlight_error, subset=['误差']),
                        use_container_width=True,
                        height=400
                    )
                    
                    st.markdown('</div>', unsafe_allow_html=True)
                    
                    # 导出功能
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_error.to_excel(writer, sheet_name='误差超标记录', index=False)
                        
                        # 添加统计信息
                        stats_df = pd.DataFrame([{
                            '核对时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                            '总核对数': total_checked,
                            '误差超标数': len(results['error']),
                            '匹配成功数': len(results['matched']),
                            '未匹配数': len(results['not_found']),
                            '误差容忍度': tolerance
                        }])
                        stats_df.to_excel(writer, sheet_name='统计信息', index=False)
                    
                    output.seek(0)
                    
                    st.download_button(
                        label="📥 下载超标记录",
                        data=output,
                        file_name=f"误差超标记录_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                else:
                    st.success("✨ 太棒了！没有发现误差超过 {} 的记录".format(tolerance))
                    
            except Exception as e:
                st.error(f"❌ 核对出错: {str(e)}")
                import traceback
                with st.expander("查看详细错误"):
                    st.code(traceback.format_exc())

else:
    st.info("👆 请上传两个Excel文件开始核对")

st.markdown('</div>', unsafe_allow_html=True)

# 使用说明
with st.expander("📖 使用说明", expanded=False):
    st.markdown("""
    ### 📋 文件格式要求
    
    **补货建议文件：**
    - 需包含以下列：`产品`、`ASIN`、`标签（FNSKU)`、`国家`、`自定义发货`
    - 国家代码：`us`(美国)、`ca`(加拿大)、`uk`(英国)、`eu`(德国)
    
    **ERP数据文件：**
    - 需包含以下列：`ASIN`、`FNSKU`、`国家`、`计划发货量`
    - 国家名称：`美国`、`加拿大`、`英国`、`德国`
    
    ### ✨ 功能特点
    
    - ✅ **智能匹配** - 自动根据ASIN、FNSKU和国家进行匹配
    - ✅ **合并单元格处理** - 自动填充产品名称，解决合并单元格问题
    - ✅ **实时统计** - 显示核对进度和结果统计
    - ✅ **差异比例** - 计算误差占总量的百分比
    - ✅ **一键导出** - 支持导出Excel格式的核对结果
    
    ### 🎯 输出说明
    
    | 列名 | 说明 |
    |------|------|
    | 产品名称 | 从补货建议中获取的产品名称 |
    | ASIN | 商品标准识别码 |
    | FNSKU | FNSKU编码 |
    | 国家 | 销售国家 |
    | **补货建议数据** | 补货建议文件中的自定义发货量 |
    | **ERP数据** | ERP系统中的计划发货量 |
    | 误差 | 两者的差值 |
    | 差异比例 | 误差占ERP数据的百分比 |
    """)

# 页脚
st.markdown("""
<div class="footer">
    <p>© 2024 误差超标核对系统 | 版本 2.0 | 设计用于高效数据核对</p>
    <p style="font-size: 0.8em; margin-top: 10px;">✨ 智能核对 · 精准定位 · 高效导出</p>
</div>
""", unsafe_allow_html=True)

# 添加time模块（用于进度条动画）
import time
