import sys  # 添加此行以导入sys模块
import pandas as pd
import xml.etree.ElementTree as ET
from openpyxl import Workbook
import os
import re

def xml_to_xlsx(xml_file, xlsx_file):
    """
    从XML文件提取设备信息并保存为XLSX格式
    """
    try:
        # 首先尝试直接读取并清理内容
        with open(xml_file, 'rb') as f:  # 使用二进制模式读取
            xml_content = f.read()
            
        # 尝试不同的编码方式
        encodings = ['utf-8', 'utf-8-sig', 'utf-16', 'gb2312', 'gbk', 'iso-8859-1']
        xml_text = None
        
        for encoding in encodings:
            try:
                xml_text = xml_content.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
                
        if xml_text is None:
            # 如果所有编码都失败，使用忽略错误的方式
            xml_text = xml_content.decode('utf-8', errors='ignore')
        
        # 清理XML内容
        xml_text = xml_text.replace('&#x0;', '')
        xml_text = xml_text.replace('&#0;', '')
        xml_text = ''.join(char for char in xml_text if char.isprintable() or char in '\n\r\t')
        
        # 移除任何可能的BOM标记
        if xml_text.startswith('\ufeff'):
            xml_text = xml_text[1:]
            
        # 尝试解析清理后的XML
        try:
            root = ET.fromstring(xml_text.encode('utf-8'))
        except ET.ParseError as e:
            # 如果解析失败，尝试更激进的清理
            xml_text = re.sub(r'&#x[0-9a-fA-F]+;', '', xml_text)  # 移除所有十六进制字符引用
            xml_text = re.sub(r'&#\d+;', '', xml_text)  # 移除所有十进制字符引用
            root = ET.fromstring(xml_text.encode('utf-8'))
        
        devices = []
        interfaces = []
        ports = []

        # 提取设备和接口信息
        for device in root.findall('.//Device'):
            # 获取ImRecord信息
            im_record = device.find('ImRecord')
            device_info = {
                'NameOfStation': device.find('NameOfStation').text if device.find('NameOfStation') is not None else '',
                'IpAddress': device.find('IpAddress').text if device.find('IpAddress') is not None else '',
                'DeviceType': device.find('DeviceType').text if device.find('DeviceType') is not None else '',
                'MAC': device.find('MAC').text if device.find('MAC') is not None else '',
                'ManufacturerID': device.find('ManufacturerID').text if device.find('ManufacturerID') is not None else '',
                'ManufacturerName': device.find('ManufacturerName').text if device.find('ManufacturerName') is not None else '',
                'Role': device.find('Role').text if device.find('Role') is not None else '',
                'RunState': device.find('RunState').text if device.find('RunState') is not None else '',
                'DeviceID': device.find('DeviceID').text if device.find('DeviceID') is not None else '',
                'GatewayIp': device.find('GatewayIp').text if device.find('GatewayIp') is not None else '',
                'NetworkMask': device.find('NetworkMask').text if device.find('NetworkMask') is not None else '',
                # ImRecord信息
                'OrderID': im_record.find('OrderID').text if im_record is not None and im_record.find('OrderID') is not None else '',
                'SerialNumber': im_record.find('SerialNumber').text if im_record is not None and im_record.find('SerialNumber') is not None else '',
                'HardwareRevision': im_record.find('HardwareRevision').text if im_record is not None and im_record.find('HardwareRevision') is not None else '',
                'SoftwareRevision': im_record.find('SoftwareRevision').text if im_record is not None and im_record.find('SoftwareRevision') is not None else '',
                'RevisionCounter': im_record.find('RevisionCounter').text if im_record is not None and im_record.find('RevisionCounter') is not None else '',
                'ProfileID': im_record.find('ProfileID').text if im_record is not None and im_record.find('ProfileID') is not None else '',
                'ProfileDetails': im_record.find('ProfileDetails').text if im_record is not None and im_record.find('ProfileDetails') is not None else '',
                'IMVersion': im_record.find('IMVersion').text if im_record is not None and im_record.find('IMVersion') is not None else '',
                'IMSupported': im_record.find('IMSupported').text if im_record is not None and im_record.find('IMSupported') is not None else '',
            }

            # 获取Modules信息
            modules = device.find('Modules')
            if modules is not None:
                for i, module in enumerate(modules.findall('Module')):  # 添加 enumerate 来获取索引
                    module_info = {
                        'ModuleIdentNumber': module.find('ModuleIdentNumber').text if module.find('ModuleIdentNumber') is not None else '',
                        'ModuleName': module.find('ModuleName').text if module.find('ModuleName') is not None else '',
                        'ModuleOrderNumber': module.find('OrderNumber').text if module.find('OrderNumber') is not None else '',
                    }
                    # 将模块信息添加到设备信息中
                    device_info.update({
                        f'Module_{i+1}_IdentNumber': module_info['ModuleIdentNumber'],
                        f'Module_{i+1}_Name': module_info['ModuleName'],
                        f'Module_{i+1}_OrderNumber': module_info['ModuleOrderNumber'],
                    })

            devices.append(device_info)

            # 提取端口信息
            for interface in device.findall('.//PnInterface'):
                port_list = interface.find('PortList')
                if port_list is not None:
                    for port in port_list.findall('Port'):
                        port_info = {
                            'DeviceName': device_info['NameOfStation'],
                            'PortID': port.find('PortID').text if port.find('PortID') is not None else '',
                            'PortDesc': port.find('PortDesc').text if port.find('PortDesc') is not None else '',
                            'OperStatus': port.find('OperStatus').text if port.find('OperStatus') is not None else '',
                            'RemotePortID': port.find('RemotePortID').text if port.find('RemotePortID') is not None else '',
                            'RemoteNameOfStation': port.find('RemoteNameOfStation').text if port.find('RemoteNameOfStation') is not None else '',
                            'RemoteMAC': port.find('RemoteMAC').text if port.find('RemoteMAC') is not None else '',
                            'NetworkLoadIn': port.find('NetworkLoadIn').text if port.find('NetworkLoadIn') is not None else '',
                            'NetworkLoadOut': port.find('NetworkLoadOut').text if port.find('NetworkLoadOut') is not None else '',
                            'IsWireless': port.find('IsWireless').text if port.find('IsWireless') is not None else '',
                            'PowerBudget': port.find('PowerBudget').text if port.find('PowerBudget') is not None else '',
                            'RxPortErrorsFrames': port.find('RxPortErrorsFrames').text if port.find('RxPortErrorsFrames') is not None else '',
                            'RemChassisIdSubtype': port.find('RemChassisIdSubtype').text if port.find('RemChassisIdSubtype') is not None else '',
                            'SwitchGroup': port.find('SwitchGroup').text if port.find('SwitchGroup') is not None else '',
                            'CableDelay': port.find('CableDelay').text if port.find('CableDelay') is not None else '',
                            'MauType': port.find('MauType').text if port.find('MauType') is not None else '',
                        }
                        ports.append(port_info)

        # 创建 Excel 工作簿
        wb = Workbook()
        ws = wb.active
        ws.title = "Combined"

        # 写入表头
        device_headers = ['NameOfStation', 'IpAddress', 'DeviceType', 'MAC', 'ManufacturerID', 
                         'ManufacturerName', 'Role', 'RunState', 'DeviceID', 'GatewayIp', 'NetworkMask',
                         'OrderID', 'SerialNumber', 'HardwareRevision', 'SoftwareRevision', 
                         'RevisionCounter', 'ProfileID', 'ProfileDetails', 'IMVersion', 'IMSupported']
        
        # 添加模块表头
        module_headers = []
        max_modules = 3
        
        # 添加端口表头定义
        port_headers = ['PortID', 'PortDesc', 'OperStatus', 'RemotePortID', 'RemoteNameOfStation',
                       'RemoteMAC', 'CableDelay', 'MauType']

        all_headers = device_headers + module_headers + port_headers
        ws.append(all_headers)

        # 为每个设备写入数据
        current_row = 2
        for device in devices:
            device_ports = [port for port in ports if port['DeviceName'] == device['NameOfStation']]
            
            if device_ports:
                for port in device_ports:
                    row_data = []
                    for header in device_headers:
                        row_data.append(device[header])
                    for header in module_headers:
                        row_data.append(device[header])
                    for header in port_headers:
                        row_data.append(port.get(header, ''))
                    ws.append(row_data)
                
                if len(device_ports) > 1:
                    for col in range(1, len(device_headers) + 1):
                        ws.merge_cells(
                            start_row=current_row,
                            start_column=col,
                            end_row=current_row + len(device_ports) - 1,
                            end_column=col
                        )
                current_row += len(device_ports)
            else:
                row_data = []
                for header in device_headers:
                    row_data.append(device[header])
                row_data.extend([''] * len(module_headers))
                row_data.extend([''] * len(port_headers))
                ws.append(row_data)
                current_row += 1

        # 设置列宽
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column].width = adjusted_width

        wb.save(xlsx_file)
        return True, "处理成功"
        
    except Exception as e:
        return False, f"处理失败: {str(e)}"

def main():
    if len(sys.argv) != 3:
        print("用法: python xml2xlsx.py <XML文件源目录> <Excel文件目标目录>")
        sys.exit(1)
        
    xml_dir = sys.argv[1]
    excel_dir = sys.argv[2]
    
    # 确保源目录存在
    if not os.path.isdir(xml_dir):
        print(f"错误: 源目录 '{xml_dir}' 不存在")
        sys.exit(1)
    
    # 确保目标目录存在，如果不存在则创建
    os.makedirs(excel_dir, exist_ok=True)
    
    # 递归遍历源目录下的所有文件
    xml_files = []
    print(f"\n开始扫描目录: {xml_dir}")
    for root, dirs, files in os.walk(xml_dir):
        # 显示当前正在扫描的目录
        current_dir = os.path.relpath(root, xml_dir)
        if current_dir == '.':
            print(f"扫描根目录...")
        else:
            print(f"扫描目录: {current_dir}")
            
        # 收集当前目录中的XML文件
        xml_count = 0
        for file in files:
            if file.lower().endswith('.xml'):
                xml_files.append(os.path.join(root, file))
                xml_count += 1
        
        if xml_count > 0:
            print(f"  |- 找到 {xml_count} 个XML文件")
    
    if not xml_files:
        print(f"\n警告: 在目录树中未找到任何XML文件")
        sys.exit(0)
    
    # 处理所有找到的XML文件
    total_files = len(xml_files)
    success_count = 0
    failed_count = 0
    
    print(f"\n总共找到 {total_files} 个XML文件，开始处理...")
    print("=" * 50)
    
    # 用于跟踪已处理的文件名
    processed_files = set()
    
    for i, xml_file in enumerate(xml_files, 1):
        # 获取相对于源目录的路径
        rel_path = os.path.relpath(xml_file, xml_dir)
        # 获取文件名和父目录名
        file_name = os.path.basename(xml_file)
        parent_dir = os.path.basename(os.path.dirname(xml_file))
        
        # 如果文件名包含"copy"，跳过处理
        if "copy" in file_name.lower():
            print(f"\n[{i}/{total_files}] 处理文件:")
            print(f"源文件: {rel_path}")
            print("⚠ 跳过: 复制文件")
            continue
        
        # 确定输出文件名
        if file_name.startswith(('20', '19')) or any(c.isdigit() for c in file_name[:2]):
            # 如果文件名是日期格式或以数字开头，使用父目录名
            output_name = f"{parent_dir}.xlsx"
        else:
            # 否则使用原文件名（去掉.xml后缀）
            output_name = os.path.splitext(file_name)[0] + '.xlsx'
        
        # 获取第一级目录
        path_parts = rel_path.split(os.sep)
        if len(path_parts) > 1:
            # 如果文件在子目录中，使用第一级目录
            first_level_dir = path_parts[0]
            xlsx_file = os.path.join(excel_dir, first_level_dir, output_name)
        else:
            # 如果文件在根目录，直接放在目标目录
            xlsx_file = os.path.join(excel_dir, output_name)
            
        # 检查目标文件是否已存在
        if os.path.exists(xlsx_file):
            print(f"\n[{i}/{total_files}] 处理文件:")
            print(f"源文件: {rel_path}")
            print(f"目标文件: {os.path.relpath(xlsx_file, excel_dir)}")
            print("⚠ 跳过: 目标文件已存在")
            continue
            
        # 检查是否是重复文件（仅对非日期格式文件）
        if not (file_name.startswith(('20', '19')) or any(c.isdigit() for c in file_name[:2])):
            if xlsx_file in processed_files:
                print(f"\n[{i}/{total_files}] 处理文件:")
                print(f"源文件: {rel_path}")
                print(f"目标文件: {os.path.relpath(xlsx_file, excel_dir)}")
                print("⚠ 跳过: 文件已存在")
                continue
        
        # 记录已处理的文件
        processed_files.add(xlsx_file)
        
        # 确保目标文件的目录存在
        os.makedirs(os.path.dirname(xlsx_file), exist_ok=True)
        
        print(f"\n[{i}/{total_files}] 处理文件:")
        print(f"源文件: {rel_path}")
        print(f"目标文件: {os.path.relpath(xlsx_file, excel_dir)}")
        
        try:
            success, message = xml_to_xlsx(xml_file, xlsx_file)
            if success:
                print(f"✓ 成功: {message}")
                success_count += 1
            else:
                print(f"✗ 失败: {message}")
                failed_count += 1
        except Exception as e:
            print(f"✗ 错误: {str(e)}")
            failed_count += 1
    
    # 打印最终统计信息
    print("\n" + "=" * 50)
    print("处理完成:")
    print(f"总文件数: {total_files}")
    print(f"成功: {success_count}")
    print(f"失败: {failed_count}")

if __name__ == "__main__":
    main()
