import streamlit as st
import pandas as pd
import os
import tempfile
from io import BytesIO
from collections import defaultdict
import logging
import traceback
from datetime import datetime

# 导入现有的处理模块
from settings import get_material_length
import convertWindow
import convertDoor

# 设置页面配置
st.set_page_config(
    page_title="DECA切割工具",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS样式
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

/* 全局样式 */
.stApp {
    background: #f8fafc;
    font-family: 'Inter', sans-serif;
}

/* 主容器 */
.main .block-container {
    padding-top: 1rem;
    padding-bottom: 1rem;
    background: #ffffff;
    border-radius: 8px;
    box-shadow: 0 2px 12px rgba(0, 0, 0, 0.06);
    margin: 0.5rem;
    border: 1px solid #e2e8f0;
}

/* 主标题 */
.main-header {
    font-size: 2rem;
    color: #1e293b;
    text-align: center;
    margin-bottom: 1.2rem;
    font-weight: 600;
}

/* 卡片样式 */
.upload-section, .result-section {
    background: #ffffff;
    padding: 0.8rem;
    border-radius: 6px;
    margin-bottom: 1rem;
    box-shadow: 0 1px 6px rgba(0, 0, 0, 0.03);
    border: 1px solid #e2e8f0;
}

/* 侧边栏样式 */
.css-1d391kg {
    background: #667eea;
    padding-top: 1rem;
}

.css-1d391kg .css-1v0mbdj {
    color: white;
}

.css-1d391kg .block-container {
    padding-top: 1rem;
    padding-bottom: 1rem;
}

/* 按钮样式 */
.stButton > button {
    background: #667eea;
    color: white;
    border: none;
    border-radius: 8px;
    padding: 0.75rem 2rem;
    font-weight: 600;
    font-size: 1rem;
}

/* 下载按钮 */
.stDownloadButton > button {
    background: #4facfe;
    color: white;
    border: none;
    border-radius: 8px;
    padding: 0.75rem 1.5rem;
    font-weight: 500;
}

/* 指标卡片 */
.metric-card {
    background: #ffffff;
    padding: 1rem;
    border-radius: 6px;
    text-align: center;
    box-shadow: 0 1px 6px rgba(0, 0, 0, 0.04);
    border: 1px solid #e2e8f0;
}

/* 文件上传区域 */
.stFileUploader {
    background: #f8fafc;
    border-radius: 6px;
    padding: 0.8rem;
    border: 2px dashed #cbd5e1;
}

/* 选择框样式 */
.stSelectbox > div > div {
    background: #ffffff;
    border-radius: 8px;
    border: 1px solid #e2e8f0;
}

/* 数据表格 */
.stDataFrame {
    border-radius: 8px;
    overflow: hidden;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
}

/* 成功和错误消息 */
.success-message {
    color: #059669;
    font-weight: 600;
    background: #ecfdf5;
    padding: 1rem;
    border-radius: 8px;
    border-left: 4px solid #059669;
}

.error-message {
    color: #dc2626;
    font-weight: 600;
    background: #fef2f2;
    padding: 1rem;
    border-radius: 8px;
    border-left: 4px solid #dc2626;
}

/* 进度条 */
.stProgress > div > div {
    background: #667eea;
    border-radius: 6px;
}

/* 页脚 */
.footer {
    background: #667eea;
    color: white;
    text-align: center;
    padding: 1.2rem;
    border-radius: 6px;
    margin-top: 1.5rem;
    font-weight: 500;
}



/* 响应式设计 */
@media (max-width: 768px) {
    .main-header {
        font-size: 2rem;
    }
    
    .upload-section, .result-section {
        padding: 1.5rem;
        margin: 1rem 0;
    }
}
</style>
""", unsafe_allow_html=True)

def setup_logger():
    """设置日志记录器"""
    logger = logging.getLogger('cutting_process')
    logger.setLevel(logging.INFO)
    
    if not logger.handlers:
        handler = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    
    return logger

def find_best_combination(lengths, target_length, cut_loss=4, min_remaining=10):
    lengths = sorted(lengths, reverse=True)
    best_combination = []
    best_remaining = target_length
    current_combination = []
    current_length = 0

    def backtrack(index):
        nonlocal best_combination, best_remaining, current_combination, current_length

        # 如果当前组合比最佳组合更好，更新最佳组合
        if len(current_combination) > 0 and target_length - current_length < best_remaining:
            best_combination = current_combination.copy()
            best_remaining = target_length - current_length

        # 基本情况：已经考虑了所有长度或剩余长度小于最小要求
        if index == len(lengths) or target_length - current_length < min_remaining:
            return

        # 尝试添加当前长度
        if current_length + lengths[index] + cut_loss <= target_length:
            current_combination.append(lengths[index])
            current_length += lengths[index] + cut_loss
            backtrack(index + 1)
            current_combination.pop()
            current_length -= lengths[index] + cut_loss

        # 尝试不添加当前长度
        backtrack(index + 1)

    backtrack(0)
    return best_combination

def process_cutting_data(df):
    """处理切割数据的核心函数"""
    logger = setup_logger()
    
    try:
        logger.info("开始处理数据")
        
        # 检查必要的列是否存在
        required_columns = ['Material Name', 'Qty', 'Length', 'Order No', 'Bin No']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"数据中缺少必要的列: {', '.join(missing_columns)}")
        
        # 创建一个临时DataFrame进行排序和计算
        temp_df = df.copy()
        temp_df['original_index'] = temp_df.index
        temp_df = temp_df.sort_values(['Material Name', 'Qty', 'Length', 'Order No', 'Bin No'], 
                                      ascending=[True, True, False, True, True])
        
        cutting_info = defaultdict(list)
        material_total_lengths = defaultdict(float)
        material_max_cutting_id = defaultdict(int)
        processed_rows = set()

        for (material, qty), material_group in temp_df.groupby(['Material Name', 'Qty']):
            material_length = get_material_length(material)
            logger.info(f"处理材料 {material}，数量 {qty}，标准长度：{material_length}")
            
            all_lengths = material_group['Length'].tolist()
            cutting_id = material_max_cutting_id[material] + 1
            
            while all_lengths:
                best_combination = find_best_combination(all_lengths, material_length - 6, 4, 10)
                remaining = material_length - sum(best_combination) - (len(best_combination) - 1) * 4 - 6
                logger.debug(f"切割 ID {cutting_id} 的最佳组合: {best_combination}，剩余长度: {remaining}")
                
                for pieces_id, length in enumerate(best_combination, 1):
                    unprocessed_rows = material_group[
                        (material_group['Length'] == length) & 
                        (~material_group.index.isin(processed_rows))
                    ]
                    
                    if not unprocessed_rows.empty:
                        row = unprocessed_rows.iloc[0]
                        key = (row['Material Name'], row['Qty'], row['Length'], row['Order No'], row['Bin No'])
                        cutting_info[(material, qty)].append((key, cutting_id, pieces_id, row['original_index']))
                        
                        processed_rows.add(row.name)
                        all_lengths.remove(length)
                        material_total_lengths[(material, qty)] += length
                        
                        logger.debug(f"添加切割信息: 材料={material}, 长度={length}, 切割ID={cutting_id}, 件数ID={pieces_id}")
                    else:
                        logger.warning(f"未找到长度为 {length} 的未处理行，跳过")
                
                cutting_id += 1
            
            material_max_cutting_id[material] = cutting_id - 1
            logger.info(f"材料 {material}，数量 {qty} 处理完成，总长度: {material_total_lengths[(material, qty)]:.2f}")

        logger.info("完成切割信息计算")

        # 创建结果 DataFrame，保持原始顺序
        result_df = df.copy()
        result_df['Cutting ID'] = 0
        result_df['Pieces ID'] = 0

        # 填充 Cutting ID 和 Pieces ID
        for (material, qty), info_list in cutting_info.items():
            for (key, cutting_id, pieces_id, original_index) in info_list:
                result_df.loc[original_index, 'Cutting ID'] = cutting_id
                result_df.loc[original_index, 'Pieces ID'] = pieces_id

        logger.info("完成 Cutting ID 和 Pieces ID 填充")
        
        return True, "数据处理成功", result_df
    
    except Exception as e:
        logger.error(f"处理数据时出错: {str(e)}", exc_info=True)
        return False, f"处理数据时出错: {str(e)}", None

def process_uploaded_file(uploaded_file, process_type):
    """处理上传的文件"""
    tmp_file_path = None
    try:
        # 创建临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name
        
        # 根据处理类型调用相应的转换函数
        if process_type == "Windows":
            df, _ = convertWindow.process_file(tmp_file_path)
        elif process_type == "Door":
            df, _ = convertDoor.process_file(tmp_file_path)
        else:
            return False, "未知的处理类型", None
        
        # 处理切割数据
        success, message, result_df = process_cutting_data(df)
        return success, message, result_df
        
    except Exception as e:
        error_message = f"处理文件时出错: {str(e)}"
        st.error(error_message)
        return False, error_message, None
    
    finally:
        # 确保临时文件被删除
        if tmp_file_path and os.path.exists(tmp_file_path):
            try:
                os.unlink(tmp_file_path)
            except Exception as cleanup_error:
                st.warning(f"清理临时文件时出错: {str(cleanup_error)}")

def convert_df_to_excel(df):
    """将DataFrame转换为Excel格式的字节流"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='CutFrame')
    return output.getvalue()

def main():
    """主函数"""
    # 页面标题
    st.markdown('''
    <div>
        <h1 class="main-header">🔧 DECA 智能切割工具</h1>
        <p style="text-align: center; font-size: 1.2rem; color: #64748b; margin-bottom: 3rem; font-weight: 500;">
            专业的门窗材料切割优化解决方案
        </p>
    </div>
    ''', unsafe_allow_html=True)
    
    # 侧边栏
    with st.sidebar:
        st.markdown("**🎯 操作中心**")
        st.caption("简单三步，完成切割优化")
        
        st.markdown("**📋 操作步骤**")
        st.markdown("""
        - 选择处理类型
        - 上传Excel文件  
        - 一键智能处理
        - 下载优化结果
        """)
        
        st.markdown("**⚙️ 系统状态**")
        st.markdown("""
        ✅ 配置已加载  
        ✅ 算法已优化  
        ✅ 运行正常
        """)
    
    # 主要内容区域
    col1, col2 = st.columns([1.2, 0.8], gap="large")
    
    with col1:
        st.markdown('''
         <div class="upload-section">
             <div style="text-align: center; margin-bottom: 1rem;">
                 <h3 style="color: #1e293b; font-weight: 600; margin-bottom: 0.3rem; font-size: 1.2rem;">📁 智能文件处理</h3>
                 <p style="color: #64748b; font-size: 0.85rem;">上传Excel文件，开始切割优化</p>
             </div>
        </div>
        ''', unsafe_allow_html=True)
        
        # 处理类型选择
        st.markdown('''
        <div style="margin-bottom: 1.5rem;">
            <label style="font-weight: 600; color: #374151; font-size: 1rem; margin-bottom: 0.5rem; display: block;">🎯 选择处理类型</label>
        </div>
        ''', unsafe_allow_html=True)
        
        process_type = st.selectbox(
            "选择处理类型",
            ["Windows", "Door"],
            help="选择要处理的文件类型",
            label_visibility="collapsed"
        )
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # 文件上传
        st.markdown('''
        <div style="margin-bottom: 1.5rem;">
            <label style="font-weight: 600; color: #374151; font-size: 1rem; margin-bottom: 0.5rem; display: block;">📎 上传Excel文件</label>
        </div>
        ''', unsafe_allow_html=True)
        
        uploaded_file = st.file_uploader(
            "选择Excel文件",
            type=['xlsx', 'xls', 'xlsm'],
            help="支持的文件格式：.xlsx, .xls, .xlsm",
            label_visibility="collapsed"
        )
    
    with col2:
        st.markdown('''
        <div class="upload-section" style="background: #f8fafc;">
            <div style="text-align: center; margin-bottom: 1rem;">
                <h3 style="color: #1e293b; font-weight: 600; margin-bottom: 0.3rem; font-size: 1.2rem;">📊 处理状态</h3>
                <p style="color: #64748b; font-size: 0.85rem;">实时监控处理进度</p>
            </div>
        </div>
        ''', unsafe_allow_html=True)
        
        if uploaded_file is not None:
            # 文件信息卡片
            st.markdown(f'''
            <div style="background: linear-gradient(145deg, #ecfdf5 0%, #d1fae5 100%); padding: 0.8rem; border-radius: 8px; margin-bottom: 0.6rem; border-left: 3px solid #059669;">
                <div style="color: #059669; font-weight: 600; margin-bottom: 0.3rem; font-size: 0.9rem;">✅ 文件已上传</div>
                <div style="color: #374151; font-size: 0.8rem; line-height: 1.4;">
                    <strong>文件:</strong> {uploaded_file.name}<br>
                    <strong>类型:</strong> {process_type}<br>
                    <strong>大小:</strong> {uploaded_file.size / 1024:.1f} KB
                </div>
            </div>
            ''', unsafe_allow_html=True)
            
            # 处理按钮
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🚀 开始智能处理", type="primary", width='stretch'):
                with st.spinner("🔄 正在进行智能切割优化，请稍候..."):
                    success, message, result_df = process_uploaded_file(uploaded_file, process_type)
                    
                    if success and result_df is not None:
                        st.session_state['result_df'] = result_df
                        st.session_state['processed_filename'] = uploaded_file.name
                        st.balloons()
                        st.success(f"🎉 {message}")
                    else:
                        st.error(f"❌ {message}")
        else:
            st.markdown('''
            <div style="background: linear-gradient(145deg, #fef3cd 0%, #fde68a 100%); padding: 0.8rem; border-radius: 8px; border-left: 3px solid #f59e0b; text-align: center;">
                <div style="color: #92400e; font-weight: 600; margin-bottom: 0.3rem; font-size: 0.9rem;">⚠️ 等待文件上传</div>
                <div style="color: #78350f; font-size: 0.8rem;">请在左侧选择并上传Excel文件</div>
            </div>
            ''', unsafe_allow_html=True)
    
    # 结果显示和下载区域
    if 'result_df' in st.session_state:
        st.markdown("<br><br>", unsafe_allow_html=True)
        
        st.markdown('''
        <div class="result-section">
            <div style="text-align: center; margin-bottom: 3rem;">
                <h2 style="color: #1e293b; font-weight: 600; margin-bottom: 0.5rem;">📈 处理结果</h2>
                <p style="color: #64748b; font-size: 1rem;">智能切割优化完成，查看详细结果</p>
            </div>
        </div>
        ''', unsafe_allow_html=True)
        
        result_df = st.session_state['result_df']
        processed_filename = st.session_state['processed_filename']
        
        # 显示统计信息 - 使用自定义卡片
        col1, col2, col3, col4 = st.columns(4, gap="medium")
        
        metrics = [
            ("📊 总行数", len(result_df), "#3b82f6"),
            ("🏗️ 材料种类", result_df['Material Name'].nunique(), "#10b981"),
            ("✂️ 切割组数", result_df['Cutting ID'].max(), "#f59e0b"),
            ("📏 总长度", f"{result_df['Length'].sum():.1f}mm", "#ef4444")
        ]
        
        for i, (col, (title, value, color)) in enumerate(zip([col1, col2, col3, col4], metrics)):
            with col:
                st.markdown(f'''
                <div class="metric-card" style="border-left: 4px solid {color};">
                    <div style="color: {color}; font-size: 0.9rem; font-weight: 600; margin-bottom: 0.5rem;">{title}</div>
                    <div style="color: #1e293b; font-size: 1.8rem; font-weight: 700;">{value}</div>
                </div>
                ''', unsafe_allow_html=True)
        
        st.markdown("<br><br>", unsafe_allow_html=True)
        
        # 显示数据预览
        st.markdown('''
        <div style="margin-bottom: 1.5rem;">
            <h3 style="color: #1e293b; font-weight: 600; margin-bottom: 1rem;">📋 数据预览</h3>
            <p style="color: #64748b; font-size: 0.9rem;">显示前10行处理结果</p>
        </div>
        ''', unsafe_allow_html=True)
        
        st.dataframe(
            result_df.head(10), 
            width='stretch',
            height=400
        )
        
        st.markdown("<br><br>", unsafe_allow_html=True)
        
        # 下载区域
        st.markdown('''
        <div style="text-align: center; margin-bottom: 2rem;">
            <h3 style="color: #1e293b; font-weight: 600; margin-bottom: 0.5rem;">📥 下载处理结果</h3>
            <p style="color: #64748b; font-size: 0.9rem;">选择您需要的文件格式进行下载</p>
        </div>
        ''', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns([1, 2, 1])
        
        with col2:
            download_col1, download_col2 = st.columns(2, gap="medium")
            
            with download_col1:
                # Excel下载
                excel_data = convert_df_to_excel(result_df)
                excel_filename = f"{os.path.splitext(processed_filename)[0]}_CutFrame.xlsx"
                st.download_button(
                    label="📊 下载Excel文件",
                    data=excel_data,
                    file_name=excel_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width='stretch'
                )
            
            with download_col2:
                # CSV下载
                csv_data = result_df.to_csv(index=False, encoding='utf-8-sig')
                csv_filename = f"{os.path.splitext(processed_filename)[0]}_CutFrame.csv"
                st.download_button(
                    label="📄 下载CSV文件",
                    data=csv_data,
                    file_name=csv_filename,
                    mime="text/csv",
                    width='stretch'
                )
    
    # 页脚
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown('''
    <div class="footer">
        <div style="display: flex; justify-content: center; align-items: center; gap: 2rem; flex-wrap: wrap;">
            <div style="text-align: center;">
                <div style="font-size: 1.2rem; font-weight: 600; margin-bottom: 0.5rem;">🔧 DECA 智能切割工具</div>
                <div style="font-size: 0.9rem; opacity: 0.8;">专业 · 高效 · 智能</div>
            </div>
            <div style="height: 40px; width: 1px; background: rgba(255,255,255,0.3);"></div>
            <div style="text-align: center;">
                <div style="font-size: 0.9rem; opacity: 0.9;">版本 2.0 | Web版</div>
                <div style="font-size: 0.8rem; opacity: 0.7;">© 2024 DECA Technology</div>
            </div>
        </div>
    </div>
    ''', unsafe_allow_html=True)

if __name__ == "__main__":
    main()