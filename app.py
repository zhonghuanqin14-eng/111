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
    
    def build_file1_index(self, df1):
        index = {}
        for idx, row in df1.iterrows():
            if pd.isna(row.get('ASIN')) or pd.isna(row.get('标签（FNSKU)')):
                continue
            asin = str(row['ASIN']).strip() if not pd.isna(row['ASIN']) else ""
            fnsku = str(row['标签（FNSKU)']).strip() if not pd.isna(row['标签（FNSKU)']) else ""
            country = self.normalize_country(row.get('国家'))
            if not country or not asin or not fnsku:
                continue
            key = (asin, fnsku, country)
            custom_shipment = self.extract_custom_shipment(row.get('自定义发货'))
            index[key] = {
                'row_index': idx,
                'custom_shipment': custom_shipment,
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
            
            key = (asin, fnsku, country)
            
            if key in file1_index:
                file1_record = file1_index[key]
                custom_qty = file1_record['custom_shipment']
                
                if custom_qty is None:
                    self.results['skipped'].append({
                        'ASIN': asin, 'FNSKU': fnsku, '国家': country,
                        '计划发货量': planned_qty, '原因': '附件1中该行无自定义发货数据'
                    })
                    continue
                
                error = abs(planned_qty - custom_qty)
                record = {
                    'ASIN': asin, 'FNSKU': fnsku, '国家': country,
                    '计划发货量': planned_qty, '自定义发货量': custom_qty,
                    '误差': round(error, 2)
                }
                
                if error <= tolerance:
                    self.results['matched'].append(record)
                else:
                    self.results['error'].append(record)
            else:
                self.results['not_found'].append({
                    'ASIN': asin, 'FNSKU': fnsku, '国家': country, '计划发货量': planned_qty
                })
        
        return self.results

# Streamlit页面配置
st.set_page_config(
    page_title="发货数据核对系统",
    page_icon="📊",
    layout="wide"
)

# 自定义CSS样式
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 30px;
    }
    .stat-card {
        background: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

# 标题
st.markdown("""
<div class="main-header">
    <h1>📊 发货数据核对系统</h1>
    <p>上传两个Excel文件，自动核对ASIN、FNSKU、国家和发货量</p>
</div>
""", unsafe_allow_html=True)

# 创建两列布局
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📁 上传附件1：补货建议")
    file1 = st.file_uploader("选择附件1（Excel文件）", type=['xlsx', 'xls'], key="file1")
    
    if file1:
        st.success(f"✅ 已上传: {file1.name}")
        try:
            df1_preview = pd.read_excel(file1, sheet_name='Sheet1')
            st.write("预览（前5行）:")
            st.dataframe(df1_preview.head(), use_container_width=True)
        except Exception as e:
            st.error(f"读取文件失败: {e}")

with col2:
    st.markdown("### 📁 上传附件2：发货计划")
    file2 = st.file_uploader("选择附件2（Excel文件）", type=['xlsx', 'xls'], key="file2")
    
    if file2:
        st.success(f"✅ 已上传: {file2.name}")
        try:
            df2_preview = pd.read_excel(file2, sheet_name='sheet1')
            st.write("预览（前5行）:")
            st.dataframe(df2_preview.head(), use_container_width=True)
        except Exception as e:
            st.error(f"读取文件失败: {e}")

# 设置误差容忍度
st.markdown("### ⚙️ 核对参数设置")
tolerance = st.slider("误差容忍度", min_value=0, max_value=200, value=80, step=10)

# 核对按钮
if file1 and file2:
    if st.button("🚀 开始核对", type="primary", use_container_width=True):
        with st.spinner("正在核对数据，请稍候..."):
            try:
                df1 = pd.read_excel(file1, sheet_name='Sheet1')
                df2 = pd.read_excel(file2, sheet_name='sheet1')
                
                checker = ShipmentDataChecker()
                results = checker.check(df1, df2, tolerance)
                
                # 显示统计卡片
                st.markdown("### 📊 核对结果统计")
                col_a, col_b, col_c, col_d = st.columns(4)
                
                with col_a:
                    st.markdown(f"""
                    <div class="stat-card">
                        <h3 style="color: #28a745;">✅ 匹配成功</h3>
                        <h2>{len(results['matched'])}</h2>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col_b:
                    st.markdown(f"""
                    <div class="stat-card">
                        <h3 style="color: #dc3545;">⚠️ 误差超标</h3>
                        <h2>{len(results['error'])}</h2>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col_c:
                    st.markdown(f"""
                    <div class="stat-card">
                        <h3 style="color: #ffc107;">❌ 未找到</h3>
                        <h2>{len(results['not_found'])}</h2>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col_d:
                    st.markdown(f"""
                    <div class="stat-card">
                        <h3 style="color: #17a2b8;">⏭️ 跳过</h3>
                        <h2>{len(results['skipped'])}</h2>
                    </div>
                    """, unsafe_allow_html=True)
                
                # 显示详细结果
                tab1, tab2, tab3, tab4 = st.tabs(["✅ 匹配成功", "⚠️ 误差超标", "❌ 未找到记录", "⏭️ 跳过记录"])
                
                with tab1:
                    if results['matched']:
                        st.dataframe(pd.DataFrame(results['matched']), use_container_width=True)
                    else:
                        st.info("暂无匹配成功的记录")
                
                with tab2:
                    if results['error']:
                        st.warning(f"以下 {len(results['error'])} 条记录误差超过 {tolerance}：")
                        st.dataframe(pd.DataFrame(results['error']), use_container_width=True)
                    else:
                        st.success("✅ 所有记录误差均在容忍范围内")
                
                with tab3:
                    if results['not_found']:
                        st.error(f"以下 {len(results['not_found'])} 条记录在附件1中找不到对应数据：")
                        st.dataframe(pd.DataFrame(results['not_found']), use_container_width=True)
                    else:
                        st.success("✅ 所有附件2的记录都在附件1中找到匹配")
                
                with tab4:
                    if results['skipped']:
                        st.info(f"以下 {len(results['skipped'])} 条记录在附件1中没有自定义发货数据：")
                        st.dataframe(pd.DataFrame(results['skipped']), use_container_width=True)
                    else:
                        st.success("✅ 没有跳过任何记录")
                
                # 导出功能
                st.markdown("### 📥 导出结果")
                export_dfs = {
                    '匹配成功': pd.DataFrame(results['matched']),
                    '误差超标': pd.DataFrame(results['error']),
                    '未找到记录': pd.DataFrame(results['not_found']),
                    '跳过记录': pd.DataFrame(results['skipped'])
                }
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for sheet_name, df in export_dfs.items():
                        if not df.empty:
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                output.seek(0)
                
                st.download_button(
                    label="📥 下载核对结果Excel",
                    data=output,
                    file_name="发货数据核对结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"核对过程中发生错误: {str(e)}")

# 页脚
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666;">
    <p>📊 发货数据核对系统 v1.0</p>
</div>
""", unsafe_allow_html=True)
