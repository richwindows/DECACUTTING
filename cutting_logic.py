"""
Core cutting logic extracted from app.py for use in API endpoints
"""
import pandas as pd
import numpy as np
from collections import defaultdict
import logging
from settings import get_material_length


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
                
                # 如果没有找到有效组合，处理剩余的单个长度
                if not best_combination:
                    if all_lengths:
                        # 取第一个剩余长度单独处理
                        single_length = all_lengths[0]
                        logger.warning(f"无法找到最佳组合，单独处理长度: {single_length}")
                        
                        # 查找对应的行
                        unprocessed_rows = material_group[
                            (material_group['Length'] == single_length) & 
                            (~material_group.index.isin(processed_rows))
                        ]
                        
                        if not unprocessed_rows.empty:
                            row = unprocessed_rows.iloc[0]
                            key = (row['Material Name'], row['Qty'], row['Length'], row['Order No'], row['Bin No'])
                            cutting_info[(material, qty)].append((key, cutting_id, 1, row['original_index']))
                            
                            processed_rows.add(row.name)
                            all_lengths.remove(single_length)
                            material_total_lengths[(material, qty)] += single_length
                            
                            logger.debug(f"添加单独切割信息: 材料={material}, 长度={single_length}, 切割ID={cutting_id}")
                        else:
                            # 如果找不到对应行，直接移除以避免无限循环
                            all_lengths.remove(single_length)
                            logger.warning(f"未找到长度为 {single_length} 的未处理行，直接移除")
                    
                    cutting_id += 1
                    continue
                
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