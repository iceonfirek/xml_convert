#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pandas as pd
import sys

def merge_cells_in_xlsx(csv_file, xlsx_file):
    """
    读取CSV文件并将相同名称和IP地址的单元格合并到Excel文件中
    """
    try:
        # 读取CSV文件
        df = pd.read_csv(csv_file, encoding='utf-8-sig')
        
        # 创建Excel writer对象
        writer = pd.ExcelWriter(xlsx_file, engine='openpyxl')
        
        # 将数据写入Excel
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        # 获取工作表
        worksheet = writer.sheets['Sheet1']
        
        # 跟踪需要合并的单元格范围
        merge_ranges = {}
        current_key = None
        start_row = None
        
        # 从第2行开始遍历(跳过标题行)
        for row in range(2, len(df) + 2):
            name = worksheet.cell(row=row, column=1).value
            ip = worksheet.cell(row=row, column=2).value
            key = f"{name}_{ip}"
            
            if current_key is None:
                current_key = key
                start_row = row
            elif key != current_key:
                # 如果key变化,合并之前的单元格
                if row - start_row > 1:
                    for col in range(1, 7):  # 合并前6列
                        merge_ranges[(start_row, col)] = (row-1, col)
                current_key = key
                start_row = row
        
        # 处理最后一组
        if start_row and row - start_row > 0:
            for col in range(1, 7):
                merge_ranges[(start_row, col)] = (row, col)
        
        # 执行单元格合并
        for start, end in merge_ranges.items():
            worksheet.merge_cells(
                start_row=start[0],
                start_column=start[1],
                end_row=end[0],
                end_column=end[1]
            )
            
        # 保存文件
        writer.close()  # 使用close()方法保存并关闭writer对象
        return True, "成功将CSV转换为Excel并合并单元格"
        
    except Exception as e:
        return False, f"处理失败: {str(e)}"

def main():
    if len(sys.argv) != 3:
        print("用法: python csv2xlsx.py <输入CSV文件> <输出XLSX文件>")
        sys.exit(1)
        
    csv_file = sys.argv[1]
    xlsx_file = sys.argv[2]
    
    success, message = merge_cells_in_xlsx(csv_file, xlsx_file)
    if success:
        print(f"成功: {message}")
    else:
        print(f"错误: {message}")
        sys.exit(1)

if __name__ == "__main__":
    main()
