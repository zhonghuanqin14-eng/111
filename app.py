# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime

# 页面配置 - 必须在最前面
st.set_page_config(
    page_title="发货数据核对",
    page_icon="📦",
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
    
    /* 标题卡片 - 居中 */
    .title-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 16px;
        padding: 20px 30px;
        margin-bottom: 30px;
        box-shadow: 0 10px 30px rgba(102, 126, 234, 0.3);
        color: white;
        text-align: center;
    }
    
    .title-text {
        font-size: 2.2em;
        font-weight: 700;
        margin: 0;
        letter-spacing: 1px;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
    }
    
    .title-subtext {
        font-size: 1em;
        margin: 5px 0 0;
        opacity: 0.9;
        font-weight: 300;
    }
    
    /* 上传卡片样式 */
    .upload-card {
        background: white;
        border-radius: 20px;
        padding: 0;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.08);
        transition: all 0.3s ease;
        height: 100%;
        overflow: hidden;
        border: 1px solid rgba(0,0,0,0.05);
    }
    
    .upload-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 20px 40px rgba(0, 0, 0, 0.12);
    }
    
    /* 卡片头部 */
    .card-header {
        padding: 25px 25px 15px 25px;
        border-bottom: 2px solid #f0f2f5;
    }
    
    .card-header-content {
        display: flex;
        align-items: center;
    }
    
    .card-icon {
        width: 60px;
        height: 60px;
        border-radius: 18px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 2.2em;
        margin-right: 18px;
    }
    
    .icon-bg-blue {
        background: linear-gradient(135deg, #2193b0 0%, #6dd5ed 100%);
    }
    
    .icon-bg-purple {
        background: linear-gradient(135deg, #8e2de2 0%, #4a00e0 100%);
    }
    
    .card-title {
        font-size: 1.5em;
        font-weight: 700;
        color: #1e293b;
        margin: 0;
    }
    
    .card-subtitle {
        font-size: 0.95em;
        color: #64748b;
        margin: 5px 0 0;
    }
    
    /* 卡片主体 */
    .card-body {
        padding: 20px 25px 25px 25px;
    }
    
    /* 自定义文件上传器样式 */
    .stFileUploader > div {
        padding: 0 !important;
    }
    
    .stFileUploader > div > div {
        border: 2px dashed #e2e8f0 !important;
        border-radius: 16px !important;
        padding: 25px !important;
        background: #f8fafc !important;
        transition: all 0.3s ease !important;
    }
    
    .stFileUploader > div > div:hover {
        border-color: #667eea !important;
        background: #f1f5f9 !important;
    }
    
    /* 文件信息卡片 */
    .file-info-card {
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
        border-radius: 16px;
        padding: 20px;
        margin-top: 15px;
        border: 1px solid #e2e8f0;
    }
    
    .file-info-row {
        display: flex;
        align-items: center;
        padding: 10px 0;
        border-bottom: 1px solid #e2e8f0;
    }
    
    .file-info-row:last-child {
        border-bottom: none;
    }
    
    .file-info-label {
        font-weight: 600;
        color: #475569;
        width: 80px;
        font-size: 0.95em;
    }
    
    .file-info-value {
        color: #0f172a;
        flex: 1;
        font-weight: 500;
    }
    
    .success-chip {
        background: #10b981;
        color: white;
        padding: 6px 16px;
        border-radius: 30px;
        font-size: 0.9em;
        font-weight: 500;
        display: inline-flex;
        align-items: center;
        gap: 6px;
    }
    
    .success-chip::before {
        content: "✓";
        font-weight: 700;
        font-size: 1.1em;
    }
    
    /* 按钮样式 */
    .stButton button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-size: 1.2em;
        font-weight: 600;
        padding: 18px 30px;
        border: none;
        border-radius: 16px;
        cursor: pointer;
        transition: all 0.3s ease;
        box-shadow: 0 10px 25px rgba(102, 126, 234, 0.3);
        width: 100%;
        letter-spacing: 1px;
        margin-top: 20px;
    }
    
    .stButton button:hover {
        transform: translateY(-3px);
        box-shadow: 0 15px 35px rgba(102, 126, 234, 0.4);
    }
    
    /* 容忍度提示 */
    .tolerance-badge {
        background: linear-gradient(135deg, #f6f9fc 0%, #edf2f7 100%);
        border-radius: 12px;
        padding: 15px 25px;
        margin: 20px 0;
        border: 1px solid #e2e8f0;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 10px;
    }
    
    .tolerance-text {
        font-size: 1.1em;
        color: #4a5568;
    }
    
    .tolerance-value {
        background: #667eea;
        color: white;
        padding: 5px 20px;
        border-radius: 30px;
        font-weight: 600;
        font-size: 1.2em;
    }
    
    /* 统计卡片 */
    .stats-grid {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 20px;
        margin: 30px 0;
    }
    
    .stat-card {
        background: white;
        border-radius: 20px;
        padding: 25px;
        text-align: center;
        box-shadow: 0 8px 20px rgba(0, 0, 0, 0.04);
        transition: transform 0.3s ease;
        border: 1px solid rgba(0,0,0,0.03);
    }
    
    .stat-card:hover {
        transform: translateY(-5px);
    }
    
    .stat-icon {
        font-size: 2.2em;
        margin-bottom: 10px;
    }
    
    .stat-number {
        font-size: 2.5em;
        font-weight: 700;
        margin: 10px 0;
        line-height: 1.2;
    }
    
    .stat-label {
        font-size: 1em;
        color: #64748b;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    /* 表格容器 */
    .table-container {
        background: white;
        border-radius: 20px;
        padding: 25px;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.05);
        margin: 25px 0;
    }
    
    .table-header {
        display: flex;
        align-items: center;
        margin-bottom: 25px;
    }
    
    .table-header-icon {
        background: #f59e0b;
        color: white;
        width: 40px;
        height: 40px;
        border-radius: 12px;
        display: flex;
        align-items: center;
        justify-content: center;
        margin-right: 15px;
        font-size: 1.2em;
    }
    
    .table-header-title {
        font-size: 1.3em;
        font-weight: 600;
        color: #1e293b;
    }
    
    .table-header-subtitle {
        font-size: 0.9em;
        color: #64748b;
        margin-left: 10px;
    }
    
    /* 表格样式 */
    .stDataFrame {
        font-family: 'Inter', sans-serif;
        border-radius: 12px;
        overflow: hidden;
    }
    
    .stDataFrame thead tr th {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-weight: 600;
        padding: 15px 12px;
        font-size: 0.95em;
    }
    
    /* 进度条 */
    .stProgress > div > div {
        background: linear-gradient(90deg, #667eea, #764ba2);
        border-radius: 10px;
        height: 10px;
    }
    
    /* 分隔线 */
    .divider {
        background: linear-gradient(90deg, transparent, #cbd5e1, transparent);
        height: 2px;
        margin: 30px 0;
    }
    
    /* 页脚 */
    .footer {
        text-align: center;
        padding: 30px 20px 10px;
        color: #64748b;
        font-size: 0.9em;
    }
    
    /* 动画 */
    @keyframes slideUp {
        from { opacity: 0; transform: translateY(20px); }
        to { opacity: 1; transform: translateY(0); }
    }
    
    .animate-in {
        animation: slideUp 0.5s ease-out;
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
                        '误差': int(error)
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

# 顶部标题卡片 - 居中
st.markdown("""
<div class="title-card animate-in">
    <div class="title-text">📦 发货数据核对</div>
    <div class="title-subtext">补货建议 & ERP数据 · 智能核对</div>
</div>
""", unsafe_allow_html=True)

# 主要内容容器
st.markdown('<div class="main-container animate-in">', unsafe_allow_html=True)

# 创建两列上传卡片
col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="upload-card">
        <div class="card-header">
            <div class="card-header-content">
                <div class="card-icon icon-bg-blue">📋</div>
                <div>
                    <div class="card-title">补货建议文件</div>
                    <div class="card-subtitle">上传 补货建议.xlsx</div>
                </div>
            </div>
        </div>
        <div class="card-body">
    """, unsafe_allow_html=True)
    
    # 文件上传器直接放在卡片体内
    file1 = st.file_uploader(" ", type=['xlsx', 'xls'], key="file1", label_visibility="collapsed")
    
    if file1:
        st.markdown(f"""
        <div class="file-info-card">
            <div class="file-info-row">
                <span class="file-info-label">文件名</span>
                <span class="file-info-value">{file1.name}</span>
            </div>
            <div class="file-info-row">
                <span class="file-info-label">大小</span>
                <span class="file-info-value">{file1.size / 1024:.1f} KB</span>
            </div>
            <div style="text-align: right; margin-top: 15px;">
                <span class="success-chip">已上传</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown('</div></div>', unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="upload-card">
        <div class="card-header">
            <div class="card-header-content">
                <div class="card-icon icon-bg-purple">📊</div>
                <div>
                    <div class="card-title">ERP数据文件</div>
                    <div class="card-subtitle">上传 发货计划.xlsx</div>
                </div>
            </div>
        </div>
        <div class="card-body">
    """, unsafe_allow_html=True)
    
    # 文件上传器直接放在卡片体内
    file2 = st.file_uploader(" ", type=['xlsx', 'xls'], key="file2", label_visibility="collapsed")
    
    if file2:
        st.markdown(f"""
        <div class="file-info-card">
            <div class="file-info-row">
                <span class="file-info-label">文件名</span>
                <span class="file-info-value">{file2.name}</span>
            </div>
            <div class="file-info-row">
                <span class="file-info-label">大小</span>
                <span class="file-info-value">{file2.size / 1024:.1f} KB</span>
            </div>
            <div style="text-align: right; margin-top: 15px;">
                <span class="success-chip">已上传</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown('</div></div>', unsafe_allow_html=True)

# 容忍度固定提示
st.markdown("""
<div class="tolerance-badge">
    <span class="tolerance-text">🔍 误差范围</span>
    <span class="tolerance-value">80</span>
</div>
""", unsafe_allow_html=True)

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
                    import time
                    time.sleep(0.01)
                
                # 固定容忍度为80
                checker = ShipmentDataChecker()
                results = checker.check(df1, df2, tolerance=80)
                
                progress_bar.progress(100)
                time.sleep(0.5)
                progress_bar.empty()
                
                st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
                
                # 统计卡片
                total_checked = len(results['matched']) + len(results['error'])
                
                col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
                
                with col_stat1:
                    st.markdown(f"""
                    <div class="stat-card">
                        <div class="stat-icon">📊</div>
                        <div class="stat-number">{total_checked}</div>
                        <div class="stat-label">总核对数</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col_stat2:
                    st.markdown(f"""
                    <div class="stat-card">
                        <div class="stat-icon">✅</div>
                        <div class="stat-number" style="color: #10b981;">{len(results['matched'])}</div>
                        <div class="stat-label">匹配成功</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col_stat3:
                    st.markdown(f"""
                    <div class="stat-card">
                        <div class="stat-icon">⚠️</div>
                        <div class="stat-number" style="color: #f59e0b;">{len(results['error'])}</div>
                        <div class="stat-label">误差超标</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col_stat4:
                    st.markdown(f"""
                    <div class="stat-card">
                        <div class="stat-icon">❓</div>
                        <div class="stat-number" style="color: #64748b;">{len(results['not_found'])}</div>
                        <div class="stat-label">未匹配</div>
                    </div>
                    """, unsafe_allow_html=True)
                
                # 只显示误差超标的结果
                if results['error']:
                    st.markdown("""
                    <div class="table-container">
                        <div class="table-header">
                            <div class="table-header-icon">⚠️</div>
                            <div class="table-header-title">误差超标记录</div>
                            <div class="table-header-subtitle">需重点关注</div>
                        </div>
                    """, unsafe_allow_html=True)
                    
                    # 创建DataFrame
                    df_error = pd.DataFrame(results['error'])
                    
                    # 重新排列列顺序
                    column_order = ['产品名称', 'ASIN', 'FNSKU', '国家', '补货建议数据', 'ERP数据', '误差']
                    df_error = df_error[column_order]
                    
                    # 添加高亮样式
                    def highlight_error(val):
                        if isinstance(val, (int, float)) and val > 80:
                            return 'background-color: #fee2e2; color: #b91c1c; font-weight: 600;'
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
                            '误差容忍度': 80
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
                    st.success("✨ 太棒了！没有发现误差超过 80 的记录")
                    
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
    - ✅ **固定容忍度** - 误差超过80即标记为超标
    - ✅ **一键导出** - 支持导出Excel格式的核对结果
    """)

# 页脚
st.markdown("""
<div class="footer">
    <p>© 2024 发货数据核对 | 版本 2.0</p>
    <p style="font-size: 0.85em; margin-top: 10px; color: #94a3b8;">
        补货建议 vs ERP数据 · 智能核对 · 快速定位差异
    </p>
</div>
""", unsafe_allow_html=True)
