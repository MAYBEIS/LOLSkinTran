import os
import pandas as pd
from pathlib import Path
import shutil

def read_excel_mappings(file_path):
    """读取Excel文件中的翻译映射，区分表1和表2"""
    xls = pd.ExcelFile(file_path)
    folder_mappings = {}  # 表1用于文件夹
    file_mappings = {}    # 表2用于文件
    
    # 获取所有表名
    sheet_names = xls.sheet_names
    print(f"发现 {len(sheet_names)} 个工作表")
    
    # 检查表1和表2是否存在
    if '表1' not in sheet_names:
        print("警告: Excel文件中未找到'表1'工作表，将不会处理文件夹重命名")
    if '表2' not in sheet_names:
        print("警告: Excel文件中未找到'表2'工作表，将不会处理文件重命名")
    
    # 读取表1 - 文件夹映射
    if 'Sheet1' in sheet_names:
        df = xls.parse('Sheet1')
        if not df.empty:
            first_col = df.columns[0]
            last_col = None
            
            for col in reversed(df.columns):
                if not df[col].isna().all():
                    last_col = col
                    break
                    
            if last_col:
                print(f"从'表1'加载文件夹映射: {first_col} -> {last_col}")
                for idx, row in df.iterrows():
                    key = row[first_col]
                    value = row[last_col]
                    if pd.notna(key) and pd.notna(value):
                        key_str = str(key).strip()
                        value_str = str(value).strip()
                        folder_mappings[key_str] = value_str
                        if idx < 10:  # 打印前10条用于调试
                            print(f"  文件夹映射: '{key_str}' -> '{value_str}'")
                print(f"从'表1'加载了 {len(folder_mappings)} 条文件夹映射")
            else:
                print("'表1'中未找到有效的映射列")
    
    # 读取表2 - 文件映射
    if 'Sheet1' in sheet_names:
        df = xls.parse('Sheet2')
        if not df.empty:
            first_col = df.columns[0]
            last_col = None
            
            for col in reversed(df.columns):
                if not df[col].isna().all():
                    last_col = col
                    break
                    
            if last_col:
                print(f"从'表2'加载文件映射: {first_col} -> {last_col}")
                for idx, row in df.iterrows():
                    key = row[first_col]
                    value = row[last_col]
                    if pd.notna(key) and pd.notna(value):
                        key_str = str(key).strip()
                        value_str = str(value).strip()
                        file_mappings[key_str] = value_str
                        if idx < 10:  # 打印前10条用于调试
                            print(f"  文件映射: '{key_str}' -> '{value_str}'")
                print(f"从'表2'加载了 {len(file_mappings)} 条文件映射")
            else:
                print("'表2'中未找到有效的映射列")
    
    return folder_mappings, file_mappings

def rename_folders_and_files(root_dir, folder_mappings, file_mappings):
    """使用不同规则重命名文件夹和文件"""
    items_to_rename = []
    
    # 第一遍扫描：收集所有需要重命名的项目
    print("开始扫描需要重命名的项目...")
    for dirpath, dirnames, filenames in os.walk(root_dir, topdown=False):
        # 处理目录（使用表1规则）
        for dirname in dirnames:
            if dirname in folder_mappings:
                relative_path = os.path.relpath(os.path.join(dirpath, dirname), root_dir)
                items_to_rename.append((relative_path, 'dir', dirname))
                print(f"  找到目录需要重命名: {relative_path}")
        
        # 处理文件（使用表2规则）
        for filename in filenames:
            name_part, ext = os.path.splitext(filename)
            if name_part in file_mappings:
                relative_path = os.path.relpath(os.path.join(dirpath, filename), root_dir)
                items_to_rename.append((relative_path, 'file', name_part))
                print(f"  找到文件需要重命名: {relative_path}")
    
    print(f"共发现 {len(items_to_rename)} 个需要重命名的项目")
    
    # 按路径深度排序，从最深层开始处理
    items_to_rename.sort(key=lambda x: x[0].count(os.sep), reverse=True)
    
    # 第二遍：执行重命名
    print("开始执行重命名操作...")
    for relative_path, item_type, original_name in items_to_rename:
        old_path = os.path.join(root_dir, relative_path)
        
        if item_type == 'dir':
            # 目录重命名（使用表1规则）
            dirname = os.path.basename(relative_path)
            parent_dir = os.path.dirname(relative_path)
            new_name = folder_mappings[dirname]
        else:
            # 文件重命名（使用表2规则）
            filename = os.path.basename(relative_path)
            parent_dir = os.path.dirname(relative_path)
            name_part, ext = os.path.splitext(filename)
            new_name = file_mappings[name_part] + ext
        
        # 构建新路径
        new_relative_path = os.path.join(parent_dir, new_name)
        new_path = os.path.join(root_dir, new_relative_path)
        
        # 处理重名冲突
        if os.path.exists(new_path):
            base, ext = os.path.splitext(new_name)
            counter = 1
            while os.path.exists(new_path):
                new_name = f"{base}_{counter}{ext}"
                new_relative_path = os.path.join(parent_dir, new_name)
                new_path = os.path.join(root_dir, new_relative_path)
                counter += 1
            print(f"  检测到重名冲突，新名称调整为: {new_name}")
        
        # 执行重命名
        try:
            os.rename(old_path, new_path)
            print(f"✓ 已重命名: {old_path} -> {new_path}")
        except Exception as e:
            print(f"✗ 重命名失败 {old_path}: {str(e)}")

def main():
    # 配置文件路径
    excel_path = "skin_names.xlsx"  # Excel文件路径
    root_directory = "skins-1"      # 要处理的根目录
    
    # 读取映射关系
    print(f"正在从 {excel_path} 读取映射关系...")
    folder_mappings, file_mappings = read_excel_mappings(excel_path)
    
    if not folder_mappings and not file_mappings:
        print("未找到有效的映射关系，请检查Excel文件格式。")
        return
    
    print(f"总共加载了 {len(folder_mappings)} 条文件夹映射和 {len(file_mappings)} 条文件映射")
    
    # 确认操作
    confirm = input(f"将要处理目录 '{root_directory}'，是否继续? (y/n): ")
    if confirm.lower() != 'y':
        print("操作已取消。")
        return
    
    # 执行重命名
    rename_folders_and_files(root_directory, folder_mappings, file_mappings)
    print("重命名操作完成!")

if __name__ == "__main__":
    main()    