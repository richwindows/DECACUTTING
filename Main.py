import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk, scrolledtext
from settings import get_material_length, get_last_directory, update_last_directory, get_settings, update_settings
from collections import defaultdict
import logging
import traceback
import subprocess
from PIL import Image, ImageTk
import importlib
import sys
import convertWindow
import convertDoor
import gc
import tkinter.font as tkFont
import csv
import shutil

def setup_logger():
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)
    
    # 创建一个文件处理器，指定UTF-8编码
    log_file = 'cutting_process.log'
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    
    # 创建一个控制台处理器，指定UTF-8编码
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    
    # 设置控制台输出编码为UTF-8
    import sys
    if hasattr(sys.stdout, 'reconfigure'):
        try:
            sys.stdout.reconfigure(encoding='utf-8')
            sys.stderr.reconfigure(encoding='utf-8')
        except:
            pass
    
    # 创建一个格式器
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # 将处理器添加到日志记录器
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
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

def open_settings(parent):
    settings_window = tk.Toplevel(parent)
    settings_window.title("材料长度设置")

    # 设置窗口大小
    window_width = 400
    window_height = 500
    # 获取屏幕尺寸
    screen_width = settings_window.winfo_screenwidth()
    screen_height = settings_window.winfo_screenheight()

    # 计算窗口位置，使其位于屏幕中央
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2

    # 设置窗口大小和位置
    settings_window.geometry(f'{window_width}x{window_height}+{x}+{y}')

    # 创建一个主框架来容纳所有内容
    main_frame = ttk.Frame(settings_window)
    main_frame.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)

    # 创建一个框架来容纳滚动条和材料列表
    list_frame = ttk.Frame(main_frame)
    list_frame.pack(expand=True, fill=tk.BOTH)
    frame = ttk.Frame(main_frame)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # 创建一个画布
    canvas = tk.Canvas(frame)
    scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    # 创建材料输入框
    entries = []
    row = 0
    ttk.Label(scrollable_frame, text="材料名称").grid(row=row, column=0, padx=5, pady=5)
    ttk.Label(scrollable_frame, text="长度").grid(row=row, column=1, padx=5, pady=5)
    
    current_settings = get_settings()
    for material in sorted(current_settings.keys()):
        if material not in ['default', 'last_open_directory', 'last_save_directory']:
            row += 1
            ttk.Label(scrollable_frame, text=material).grid(row=row, column=0, padx=5, pady=5)
            entry = ttk.Entry(scrollable_frame, width=10)
            entry.insert(0, str(current_settings[material]))
            entry.grid(row=row, column=1, padx=5, pady=5)
            entries.append((material, entry))

    # 添加新材料的输入框
    row += 1
    new_material_entry = ttk.Entry(scrollable_frame, width=20)
    new_material_entry.grid(row=row, column=0, padx=5, pady=5)
    new_length_entry = ttk.Entry(scrollable_frame, width=10)
    new_length_entry.grid(row=row, column=1, padx=5, pady=5)

    def add_new_material():
        material = new_material_entry.get().strip()
        length = new_length_entry.get().strip()
        if material and length:
            try:
                length = float(length)
                new_settings = {material: length}
                update_settings(new_settings)
                messagebox.showinfo("添加成功", f"已添加材料 {material}，长度为 {length}")
                settings_window.destroy()
                open_settings(parent)
            except ValueError:
                messagebox.showerror("输入错误", "长度必须是一个数字")
        else:
            messagebox.showerror("输入错误", "材料名称和长度都不能为空")

    ttk.Button(scrollable_frame, text="添加新材料", command=add_new_material).grid(row=row+1, column=0, columnspan=2, pady=10)

    def save():
        new_settings = {material: float(entry.get()) for material, entry in entries}
        new_settings['default'] = current_settings['default']
        new_settings['last_open_directory'] = current_settings['last_open_directory']
        new_settings['last_save_directory'] = current_settings['last_save_directory']
        update_settings(new_settings)
        settings_window.destroy()

    ttk.Button(settings_window, text="保存", command=save).pack(pady=10)

    # 放置画布和滚动条
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

def process_cutting_data(df, csv_file):
    
    df.to_csv(csv_file, index=False, encoding='utf-8-sig')
    
    logger = setup_logger()
    try:
        logger.info(f"开始处理数据")
        
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

        # 输出每种材料的总长度到日志
        # logger.info("\n材料总长度信息:")
        # for (material, qty), total_length in material_total_lengths.items():
        #     logger.info(f"{material} (数量 {qty}): {total_length:.2f}")

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

        # 保存处理后的数据
        file_extension = os.path.splitext(csv_file)[1].lower()
        if file_extension in ['.xlsx', '.xls', '.xlsm']:
            result_df.to_excel(csv_file, index=False, engine='openpyxl')
        else:
            result_df.to_csv(csv_file, index=False)
        
        logger.info(f"数据处理成功，已保存到: {csv_file}")
        return True, "数据处理成功"
    
    except Exception as e:
        logger.error(f"处理数据时出错: {str(e)}", exc_info=True)
        return False, f"处理数据时出错: {str(e)}"

def process_single_file(input_file, type_value):
    file_extension = os.path.splitext(input_file)[1].lower()
    
    if file_extension not in ['.xlsx', '.xls', '.xlsm']:
        return False, "不支持的文件类型"

    try:
        if type_value == "Windows":
            df,csv_file = convertWindow.process_file(input_file)
            print("转换窗户和排序完成！")
        elif type_value == "Door":
            df,csv_file = convertDoor.process_file(input_file)
            print("转换门和排序完成！")
        else:
            return False, "未知的处理类型"

        success, message = process_cutting_data(df, csv_file)
        return success, message

    except Exception as e:
        error_message = f"处理文件时出错: {str(e)}"
        print(error_message)
        traceback.print_exc()
        return False, error_message

    finally:
        # 清理内存
        if 'df' in locals():
            del df
        gc.collect()

class ModernUI:
    def __init__(self, master):
        self.master = master
        self.master.title("材料切割处理程序")
        
        # 设置窗口大小
        window_width = 400
        window_height = 300  # 增加高度以容纳logo
        
        # 获取屏幕尺寸
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        
        # 计算窗口位置
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        # 设置窗口大小和位置
        self.master.geometry(f'{window_width}x{window_height}+{x}+{y}')
        self.master.configure(bg="#ffffff")

        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("TButton", padding=10, relief="flat", background="#4a90e2", foreground="white", font=("Helvetica", 12))
        self.style.map("TButton", background=[("active", "#3a7abd")])

        self.create_widgets()

    def create_widgets(self):
        main_frame = tk.Frame(self.master, bg="#ffffff")
        main_frame.pack(expand=True, fill=tk.BOTH)

        # 添加logo
        logo_path = os.path.join(os.path.dirname(__file__), "table-saw.png")
        if os.path.exists(logo_path):
            try:
                logo_image = Image.open(logo_path)
                logo_image = logo_image.resize((100, 100), Image.LANCZOS)  # 使用 LANCZOS 替代 ANTIALIAS
                self.logo_photo = ImageTk.PhotoImage(logo_image)
                logo_label = tk.Label(main_frame, image=self.logo_photo, bg="#ffffff")
                logo_label.pack(pady=(20, 20))
            except Exception as e:
                print(f"加载图标时出错: {e}")
        else:
            print(f"图标文件不存在: {logo_path}")

        button_frame = tk.Frame(main_frame, bg="#ffffff")
        button_frame.pack()

        start_button = ttk.Button(button_frame, text="开始处理", command=self.show_type_selection, width=15)
        start_button.pack(pady=10)

        settings_button = ttk.Button(button_frame, text="设置材料长度", command=self.open_settings, width=15)
        settings_button.pack(pady=10)

    def show_type_selection(self):
        for widget in self.master.winfo_children():
            widget.destroy()

        selection_frame = tk.Frame(self.master, bg="#ffffff")
        selection_frame.place(relx=0.5, rely=0.5, anchor="center")

        button_frame = tk.Frame(selection_frame, bg="#ffffff")
        button_frame.pack()
        
        windows_button = ttk.Button(button_frame, text="窗", command=lambda: self.select_files("Windows"), width=10)
        windows_button.pack(side="left", padx=10, pady=10)

        door_button = ttk.Button(button_frame, text="门", command=lambda: self.select_files("Door"), width=10)
        door_button.pack(side="left", padx=10, pady=10)

        

    def select_files(self, type_value):
        initial_dir = get_last_directory('open')
        input_files = filedialog.askopenfilenames(
            title="选择文件",
            initialdir=initial_dir,
            filetypes=[("All files", "*.*"), ("CSV files", "*.csv"), ("Excel files", "*.xlsx;*.xls;*.xlsm")]
        )
        if not input_files:
            return  # 如果用户没有选择文件，直接返回

        update_last_directory('open', os.path.dirname(input_files[0]))

        # 选择保存文件夹
        initial_save_dir = get_last_directory('save')
        output_folder = filedialog.askdirectory(title="选择保存文件夹", initialdir=initial_save_dir)
        if not output_folder:
            output_folder = initial_save_dir  # 如果用户取消选择，使用默认文件夹
        else:
            update_last_directory('save', output_folder)  # 如果选择了新文件夹，更新保存目录

        results = []
        for input_file in input_files:
            file_name = os.path.basename(input_file)
            output_file = os.path.join(output_folder, f"{os.path.splitext(file_name)[0]}_CutFrame.csv")
            
            if os.path.exists(output_file):
                overwrite = messagebox.askyesno("文件已存在", f"{output_file} 已存在。是否覆盖？")
                if not overwrite:
                    continue

            try:
                success, message = process_single_file(input_file, type_value)
                if success:
                    # 找到处理后的文件并移动到新的输出文件夹
                    processed_file = input_file.replace(os.path.splitext(input_file)[1], "_CutFrame.csv")
                    if os.path.exists(processed_file):
                        shutil.move(processed_file, output_file)
                    else:
                        message += " 但未找到处理后的文件。"
                        success = False
                results.append((input_file, success, message))
            except Exception as e:
                print(f"处理文件 {input_file} 时发生错误:")
                traceback.print_exc()
                results.append((input_file, False, str(e)))

            # 每处理完一个文件就更新结果显示
            self.update_results(results)
    
    def update_results(self, results):
        if not hasattr(self, 'result_window'):
            self.create_result_window()
        
        self.tree.delete(*self.tree.get_children())
        for input_file, success, message in results:
            status = "成功" if success else "失败"
            self.tree.insert("", "end", values=(input_file, status, message))
    
    def create_result_window(self):
        self.result_window = tk.Toplevel(self.master)
        self.result_window.title("处理结果")
        
        # 设置结果窗口大小和位置
        result_width = 500
        result_height = 300
        screen_width = self.result_window.winfo_screenwidth()
        screen_height = self.result_window.winfo_screenheight()
        x = (screen_width - result_width) // 2
        y = (screen_height - result_height) // 2
        self.result_window.geometry(f'{result_width}x{result_height}+{x}+{y}')
        self.result_window.configure(bg="#ffffff")

        self.tree = ttk.Treeview(self.result_window, columns=("文件", "状态", "消息"), show="headings", style="Custom.Treeview")
        self.tree.heading("文件", text="文件")
        self.tree.heading("状态", text="状态")
        self.tree.heading("消息", text="消息")
        self.tree.pack(fill="both", expand=True, padx=20, pady=20)

    def open_settings(self):
        open_settings(self.master)

def main():
    root = tk.Tk()
    app = ModernUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
