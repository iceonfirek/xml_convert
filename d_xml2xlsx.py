import sys  # 添加此行以导入sys模块
import pandas as pd
import xml.etree.ElementTree as ET
from openpyxl import Workbook

def xml_to_xlsx(xml_file, xlsx_file):
    """
    从XML文件提取设备信息并保存为XLSX格式,合并相同名称和IP的单元格
    """
    try:
        # 解析XML文件
        tree = ET.parse(xml_file)
        root = tree.getroot()
        
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
        print(f"错误: {str(e)}")
        return False, f"处理失败: {str(e)}"

def main():
    if len(sys.argv) != 3:
        print("用法: python xml2xlsx.py <输入XML文件> <输出Excel文件>")
        sys.exit(1)
        
    xml_file = sys.argv[1]
    xlsx_file = sys.argv[2]
    
    success, message = xml_to_xlsx(xml_file, xlsx_file)
    if success:
        print(f"成功: {message}")
    else:
        print(f"错误: {message}")
        sys.exit(1)

if __name__ == "__main__":
    main()
