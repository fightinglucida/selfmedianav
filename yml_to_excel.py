#!/usr/bin/env python
# -*- coding: utf-8 -*-

import yaml
import pandas as pd
import os
import sys

def yml_to_excel(yml_file_path, excel_output_path=None):
    """
    将navsites.yml文件转换为Excel表格
    
    参数:
    yml_file_path: yml文件的路径
    excel_output_path: 输出Excel文件的路径，如果为None，则默认为yml文件同目录下的navsites.xlsx
    
    返回:
    excel_output_path: 生成的Excel文件路径
    """
    # 如果未指定输出路径，则使用默认路径
    if excel_output_path is None:
        dir_path = os.path.dirname(yml_file_path)
        excel_output_path = os.path.join(dir_path, 'navsites.xlsx')
    
    # 读取YAML文件
    try:
        with open(yml_file_path, 'r', encoding='utf-8') as file:
            data = yaml.safe_load(file)
    except Exception as e:
        print(f"读取YAML文件时出错: {e}")
        return None
    
    # 准备存储所有数据的列表
    all_data = []
    
    # 遍历YAML数据结构，提取所需信息
    for taxonomy_item in data:
        taxonomy = taxonomy_item.get('taxonomy', '')
        icon = taxonomy_item.get('icon', '')
        
        for list_item in taxonomy_item.get('list', []):
            term = list_item.get('term', '')
            
            for link in list_item.get('links', []):
                # 提取每个链接的信息
                row_data = {
                    'taxonomy': taxonomy,
                    'icon': icon,
                    'term': term,
                    'title': link.get('title', ''),
                    'logo': link.get('logo', ''),
                    'url': link.get('url', ''),
                    'post_url': link.get('post_url', ''),
                    'description': link.get('description', ''),
                    'qrcode': link.get('qrcode', '')
                }
                all_data.append(row_data)
    
    # 创建DataFrame
    df = pd.DataFrame(all_data)
    
    # 定义列的顺序
    columns_order = ['taxonomy', 'icon', 'term', 'title', 'logo', 'url', 'post_url', 'description', 'qrcode']
    df = df[columns_order]
    
    # 保存为Excel文件
    try:
        df.to_excel(excel_output_path, index=False, engine='openpyxl')
        print(f"Excel文件已成功生成: {excel_output_path}")
        return excel_output_path
    except Exception as e:
        print(f"生成Excel文件时出错: {e}")
        return None

def excel_to_yml(excel_file_path, yml_output_path=None):
    """
    将Excel表格转换回navsites.yml文件格式
    
    参数:
    excel_file_path: Excel文件的路径
    yml_output_path: 输出yml文件的路径，如果为None，则默认为Excel文件同目录下的navsites_new.yml
    
    返回:
    yml_output_path: 生成的yml文件路径
    """
    # 如果未指定输出路径，则使用默认路径
    if yml_output_path is None:
        dir_path = os.path.dirname(excel_file_path)
        yml_output_path = os.path.join(dir_path, 'navsites_new.yml')
    
    # 读取Excel文件
    try:
        df = pd.read_excel(excel_file_path)
    except Exception as e:
        print(f"读取Excel文件时出错: {e}")
        return None
    
    # 准备YAML数据结构
    yaml_data = []
    
    # 按taxonomy分组
    for taxonomy, taxonomy_group in df.groupby('taxonomy'):
        taxonomy_item = {
            'taxonomy': taxonomy,
            'icon': taxonomy_group['icon'].iloc[0],  # 假设同一taxonomy的icon相同
            'list': []
        }
        
        # 按term分组
        for term, term_group in taxonomy_group.groupby('term'):
            term_item = {
                'term': term,
                'links': []
            }
            
            # 添加每个链接
            for _, row in term_group.iterrows():
                # 创建链接项，按照原始YAML的顺序添加字段
                link_item = {}
                
                # 按照固定顺序添加字段
                if not pd.isna(row['title']) and row['title']:
                    link_item['title'] = row['title']
                
                if not pd.isna(row['logo']) and row['logo']:
                    link_item['logo'] = row['logo']
                
                # 保持与原始YAML相同的字段顺序
                if not pd.isna(row['url']) and row['url']:
                    link_item['url'] = row['url']
                
                if not pd.isna(row['post_url']) and row['post_url']:
                    link_item['post_url'] = row['post_url']
                
                if not pd.isna(row['description']) and row['description']:
                    link_item['description'] = row['description']
                
                if not pd.isna(row['qrcode']) and row['qrcode']:
                    link_item['qrcode'] = row['qrcode']
                
                term_item['links'].append(link_item)
            
            taxonomy_item['list'].append(term_item)
        
        yaml_data.append(taxonomy_item)
    
    # 自定义YAML格式化，确保与原始格式一致
    class IndentDumper(yaml.Dumper):
        def increase_indent(self, flow=False, indentless=False):
            return super(IndentDumper, self).increase_indent(flow, False)
    
    # 保存为YAML文件
    try:
        with open(yml_output_path, 'w', encoding='utf-8') as file:
            file.write('---\n')  # YAML文件开头
            yaml.dump(yaml_data, file, Dumper=IndentDumper, default_flow_style=False, 
                      allow_unicode=True, sort_keys=False, indent=2)
        print(f"YAML文件已成功生成: {yml_output_path}")
        return yml_output_path
    except Exception as e:
        print(f"生成YAML文件时出错: {e}")
        return None

if __name__ == "__main__":
    
    
    # 设置默认文件路径
    yml_file = r"f:\linshutech\资源导航站\自媒体工具导航\data\navsites.yml"
    excel_file = r"f:\linshutech\资源导航站\自媒体工具导航\navsites.xlsx"
    new_yml_file = r"f:\linshutech\资源导航站\自媒体工具导航\data\navsites_new.yml"
    
    # 检查命令行参数
    if len(sys.argv) > 1:
        mode = sys.argv[1].lower()
        
        # 如果提供了文件路径参数
        if len(sys.argv) > 2:
            if mode == "to_excel":
                yml_file = sys.argv[2]
                if len(sys.argv) > 3:
                    excel_file = sys.argv[3]
            elif mode == "to_yml":
                excel_file = sys.argv[2]
                if len(sys.argv) > 3:
                    new_yml_file = sys.argv[3]
        
        # 执行相应的转换
        if mode == "to_excel":
            print(f"将 {yml_file} 转换为 {excel_file}...")
            yml_to_excel(yml_file, excel_file)
        elif mode == "to_yml":
            print(f"将 {excel_file} 转换为 {new_yml_file}...")
            excel_to_yml(excel_file, new_yml_file)
        else:
            print("未知的转换模式。使用方法：")
            print("  python yml_to_excel.py to_excel [yml文件路径] [excel输出路径]")
            print("  python yml_to_excel.py to_yml [excel文件路径] [yml输出路径]")
    else:
        # 默认执行两种转换
        # print(f"将 {yml_file} 转换为 {excel_file}...")
        # yml_to_excel(yml_file, excel_file)
        
        print(f"将 {excel_file} 转换为 {new_yml_file}...")
        excel_to_yml(excel_file, new_yml_file)
