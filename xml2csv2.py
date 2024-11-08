#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
import re
import csv
import xml.etree.ElementTree as ET

def clean_xml_content(xml_path):
    """
    清理XML文件中的无效字符引用
    
    Args:
        xml_path: XML文件路径
    Returns:
        cleaned_content: 清理后的XML内容
    """
    try:
        with open(xml_path, 'r', encoding='utf-8') as f:
            content = f.read()
            
        # 替换无效的字符引用
        # 移除所有 &#x 开头的十六进制字符引用
        content = re.sub(r'&#x[0-9a-fA-F]+;', '', content)
        # 移除所有 &# 开头的十进制字符引用
        content = re.sub(r'&#\d+;', '', content)
        
        return content
    except UnicodeDecodeError:
        # 如果UTF-8解码失败，尝试其他编码
        try:
            with open(xml_path, 'r', encoding='latin1') as f:
                content = f.read()
            content = re.sub(r'&#x[0-9a-fA-F]+;', '', content)
            content = re.sub(r'&#\d+;', '', content)
            return content
        except Exception as e:
            raise Exception(f"无法读取文件编码: {str(e)}")

def validate_xml_structure(xml_content):
    """
    检查XML内容结构是否满足转换要求
    
    Args:
        xml_content: XML文件内容
    Returns:
        (bool, str): (是否有效, 错误信息)
    """
    try:
        root = ET.fromstring(xml_content)
        
        devices = root.find('DeviceCollection')
        if devices is None:
            return False, "找不到 DeviceCollection 元素"
            
        if len(list(devices)) == 0:
            return False, "DeviceCollection 中没有设备数据"
            
        return True, ""
        
    except ET.ParseError as e:
        return False, f"XML格式错误: {str(e)}"
    except Exception as e:
        return False, f"验证XML结构失败: {str(e)}"

def xml_to_csv(xml_path, csv_path):
    """
    将XML文件转换为CSV格式，合并相同设备的基本信息
    """
    try:
        # 定义CSV表头
        headers = [
            '#', '名称', '设备类型', 'IP 地址', '子网掩码', 'MAC 地址', '角色', 
            '供应商名称', '订单号', '固件版本', '硬件版本',
            '#', '名称', 'IP 地址', '子网掩码', 'MAC 地址',
            '#', '端口 ID', '端口说明', '伙伴端口 ID', '伙伴设备名称', '功率预算 [dB]',
            '#', '模块名称', '供应商', '订货号', '序列号', '固件版本', '硬件版本', ''
        ]
        
        tree = ET.parse(xml_path)
        root = tree.getroot()
        
        devices = root.find('DeviceCollection')
        if devices is None:
            return False, "找不到DeviceCollection元素"
            
        rows = []
        device_count = 1
        
        # 用于存储已处理的设备信息
        processed_devices = {}
        
        for device in devices.findall('Device'):
            # 设备唯一标识
            device_key = (
                device.find('NameOfStation').text if device.find('NameOfStation') is not None else '',
                device.find('IpAddress').text if device.find('IpAddress') is not None else '',
                device.find('MAC').text if device.find('MAC') is not None else ''
            )
            
            # 获取基本设备信息
            base_info = {
                '#': str(device_count),
                '名称': device_key[0],
                '设备类型': device.find('DeviceType').text if device.find('DeviceType') is not None else '',
                'IP 地址': device_key[1],
                '子网掩码': device.find('NetworkMask').text if device.find('NetworkMask') is not None else '',
                'MAC 地址': device_key[2],
                '角色': device.find('Role').text if device.find('Role') is not None else '',
                '供应商名称': device.find('ManufacturerName').text if device.find('ManufacturerName') is not None else '',
                '订单号': device.find('.//OrderID').text if device.find('.//OrderID') is not None else '',
                '固件版本': device.find('.//SoftwareRevision').text if device.find('.//SoftwareRevision') is not None else '',
                '硬件版本': device.find('.//HardwareRevision').text if device.find('.//HardwareRevision') is not None else ''
            }
            
            # 复制设备基本信息到第二组
            device_info_2 = {
                '#': '1',
                '名称': device_key[0],
                'IP 地址': device_key[1],
                '子网掩码': base_info['子网掩码'],
                'MAC 地址': device_key[2]
            }
            
            # 获取端口信息
            interface = device.find('.//PnInterface')
            if interface is not None:
                port_list = interface.find('PortList')
                if port_list is not None:
                    port_count = 1
                    first_row = True
                    for port in port_list.findall('Port'):
                        row = {header: '' for header in headers}
                        
                        # 只在第一行显示设备基本信息
                        if first_row and device_key not in processed_devices:
                            row.update(base_info)
                            row.update(device_info_2)
                            processed_devices[device_key] = True
                            first_row = False
                        
                        # 填充端口信息
                        row['#'] = str(port_count)
                        row['端口 ID'] = port.find('PortID').text if port.find('PortID') is not None else ''
                        row['端口说明'] = port.find('PortDesc').text if port.find('PortDesc') is not None else ''
                        row['伙伴端口 ID'] = port.find('RemotePortID').text if port.find('RemotePortID') is not None else ''
                        row['伙伴设备名称'] = port.find('RemoteNameOfStation').text if port.find('RemoteNameOfStation') is not None else ''
                        row['功率预算 [dB]'] = port.find('PowerBudget').text if port.find('PowerBudget') is not None else ''
                        
                        # 添加模块信息
                        modules = device.findall('.//Module')
                        if modules and port_count <= len(modules):
                            module = modules[port_count - 1]
                            row['#'] = str(port_count)
                            row['模块名称'] = module.find('OrderID').text if module.find('OrderID') is not None else ''
                            row['供应商'] = base_info['供应商名称']
                            row['订货号'] = module.find('OrderID').text if module.find('OrderID') is not None else ''
                            row['序列号'] = module.find('SerialNumber').text if module.find('SerialNumber') is not None else ''
                            row['固件版本'] = module.find('SoftwareRevision').text if module.find('SoftwareRevision') is not None else ''
                            row['硬件版本'] = module.find('HardwareRevision').text if module.find('HardwareRevision') is not None else ''
                        
                        rows.append(row)
                        port_count += 1
                    
                    # 添加空行
                    rows.append({header: '' for header in headers})
                
            device_count += 1
        
        # 写入CSV文件
        with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=headers)
            writer.writeheader()
            writer.writerows(rows)
            
        return True, f"成功转换 {len(rows)} 条记录"
        
    except ET.ParseError as e:
        return False, f"XML解析错误: {str(e)}"
    except Exception as e:
        return False, f"转换失败: {str(e)}"

def process_directory(input_dir, output_dir):
    """
    批量处理指定目录下的所有XML文件，只保留第一级目录结构
    如果目标文件已存在则跳过处理
    
    Args:
        input_dir: XML文件所在目录
        output_dir: CSV文件输出目录
    """
    # 确保输出目录存在
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 统计处理结果
    total_files = 0
    success_count = 0
    skipped_count = 0
    failed_files = []
    
    # 遍历输入目录
    for root, dirs, files in os.walk(input_dir):
        # 过滤出XML文件
        xml_files = [f for f in files if f.lower().endswith('.xml')]
        
        if not xml_files:
            continue
            
        # 获取相对于输入目录的路径
        rel_path = os.path.relpath(root, input_dir)
        path_parts = rel_path.split(os.sep)
        
        # 只取第一级目录
        if len(path_parts) > 1:
            output_subdir = os.path.join(output_dir, path_parts[0])
        else:
            output_subdir = output_dir
            
        # 确保输出子目录存在
        if not os.path.exists(output_subdir):
            os.makedirs(output_subdir)
            
        for file in xml_files:
            total_files += 1
            xml_path = os.path.join(root, file)
            
            # 生成输出文件名
            # 如果当前目录只有一个XML文件，使用最后一级目录名作为文件名
            if len(xml_files) == 1:
                csv_filename = os.path.basename(root) + '.csv'
            else:
                csv_filename = os.path.splitext(file)[0] + '.csv'
            
            csv_path = os.path.join(output_subdir, csv_filename)
            
            print(f"[{total_files}] 处理文件：")
            print(f"源文件：{xml_path}")
            print(f"目标文件：{csv_path}")
            
            # 检查目标文件是否已存在
            if os.path.exists(csv_path):
                print("✓ 跳过：目标文件已存在")
                skipped_count += 1
                print("=" * 60)
                continue
            
            try:
                # 清理XML内容
                cleaned_content = clean_xml_content(xml_path)
                
                # 验证清理后的XML结构
                valid, message = validate_xml_structure(cleaned_content)
                if not valid:
                    print(f"✗ 失败：{message}")
                    failed_files.append((xml_path, message))
                    continue
                
                # 创建临时文件存储清理后的内容
                import tempfile
                with tempfile.NamedTemporaryFile(mode='w', suffix='.xml', delete=False, encoding='utf-8') as temp_file:
                    temp_file.write(cleaned_content)
                    temp_path = temp_file.name
                
                try:
                    # 使用临时文件进行转换
                    success, message = xml_to_csv(temp_path, csv_path)
                    if success:
                        success_count += 1
                        print(f"✓ 成功：{message}")
                    else:
                        failed_files.append((xml_path, message))
                        print(f"✗ 失败：{message}")
                finally:
                    # 删除临时文件
                    os.unlink(temp_path)
                    
            except Exception as e:
                failed_files.append((xml_path, str(e)))
                print(f"✗ 失败：处理文件时发生错误 - {str(e)}")
            
            print("=" * 60)
    
    # 打印处理总结
    print("\n处理完成：")
    print(f"总文件数：{total_files}")
    print(f"成功：{success_count}")
    print(f"跳过：{skipped_count}")
    print(f"失败：{len(failed_files)}")
    
    # 如果有失败的文件，打印详细信息
    if failed_files:
        print("\n失败文件列表：")
        for file_path, error in failed_files:
            print(f"- {file_path}")
            print(f"  错误：{error}")

def main():
    if len(sys.argv) != 3:
        print("用法: python xml2csv.py <输入目录> <输出目录>")
        sys.exit(1)
    
    input_dir = sys.argv[1]
    output_dir = sys.argv[2]
    
    # 检查输入目录是否存在
    if not os.path.exists(input_dir):
        print(f"错误: 输入目录 '{input_dir}' 不存在")
        sys.exit(1)
    
    # 开始处理
    try:
        process_directory(input_dir, output_dir)
    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()