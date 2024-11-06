#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import csv
import xml.etree.ElementTree as ET

def validate_xml_structure(xml_file):
    """
    检查XML文件结构是否满足转换要求
    返回: (bool, str) - (是否有效, 错误信息)
    """
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        
        devices = root.find('DeviceCollection')
        if devices is None:
            return False, "找不到 DeviceCollection 元素"
            
        if len(list(devices)) == 0:
            return False, "DeviceCollection 中没有设备数据"
            
        return True, ""
        
    except ET.ParseError as e:
        return False, f"XML格式错误: {str(e)}"
    except Exception as e:
        return False, f"读取XML文件失败: {str(e)}"

def extract_port_info(port):
    """从端口元素中提取信息"""
    return {
        'Port_ID': port.findtext('PortID', ''),
        'Port_Desc': port.findtext('PortDesc', ''),
        'Remote_Port_ID': port.findtext('RemotePortID', ''),
        'Remote_Station': port.findtext('RemoteNameOfStation', ''),
        'Remote_MAC': port.findtext('RemoteMAC', ''),
        'Port_Status': port.findtext('OperStatus', '')
    }

def extract_device_info(device):
    """从设备元素中提取信息，包括端口信息"""
    base_info = {
        'NameOfStation': device.findtext('NameOfStation', ''),
        'IpAddress': device.findtext('IpAddress', ''),
        'DeviceType': device.findtext('DeviceType', ''),
        'MAC': device.findtext('MAC', ''),
        'ManufacturerName': device.findtext('ManufacturerName', ''),
        'RunState': device.findtext('RunState', ''),
        'Port_ID': '',
        'Port_Desc': '',
        'Remote_Port_ID': '',
        'Remote_Station': '',
        'Remote_MAC': '',
        'Port_Status': ''
    }
    
    device_records = []
    
    # 查找所有端口
    interfaces = device.find('Interfaces')
    if interfaces is not None:
        for pn_interface in interfaces.findall('PnInterface'):
            port_list = pn_interface.find('PortList')
            if port_list is not None:
                for port in port_list.findall('Port'):
                    record = base_info.copy()
                    port_info = extract_port_info(port)
                    record.update(port_info)
                    device_records.append(record)
    
    # 如果没有找到端口信息，至少返回设备基本信息
    if not device_records:
        device_records.append(base_info)
    
    return device_records

def xml_to_csv(xml_file, csv_file):
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        
        devices = root.find('DeviceCollection')
        if devices is None:
            raise Exception("找不到设备集合")
            
        # 提取所有设备和端口信息
        all_records = []
        for device in devices.findall('Device'):
            device_records = extract_device_info(device)
            all_records.extend(device_records)

        # 按设备名称和IP地址排序
        all_records.sort(key=lambda x: (x['NameOfStation'], x['IpAddress'], x['Port_ID']))

        # 写入CSV文件
        if all_records:
            fieldnames = [
                'NameOfStation', 'IpAddress', 'DeviceType', 'MAC', 
                'ManufacturerName', 'RunState',
                'Port_ID', 'Port_Desc', 
                'Remote_Port_ID', 'Remote_Station', 'Remote_MAC',
                'Port_Status'
            ]
            
            with open(csv_file, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(all_records)
            return True, f"成功将 {len(all_records)} 条记录写入 CSV 文件"
        else:
            return False, "没有找到任何设备数据"
            
    except Exception as e:
        return False, f"处理失败: {str(e)}"

def main():
    if len(sys.argv) != 3:
        print("用法: python xml2csv.py <输入XML文件> <输出CSV文件>")
        sys.exit(1)
        
    xml_file = sys.argv[1]
    csv_file = sys.argv[2]
    
    # 验证XML结构
    valid, message = validate_xml_structure(xml_file)
    if not valid:
        print(f"错误: {message}")
        sys.exit(1)
        
    # 转换文件
    success, message = xml_to_csv(xml_file, csv_file)
    if success:
        print(f"成功: {message}")
    else:
        print(f"错误: {message}")
        sys.exit(1)

if __name__ == "__main__":
    main()
