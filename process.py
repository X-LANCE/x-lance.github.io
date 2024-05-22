import re

import pandas as pd
from pypinyin import lazy_pinyin

excel_file_path = 'D:/Downloads/实验室信息收集_答卷数据_2024_05_22_20_12_01.xlsx'
df = pd.read_excel(excel_file_path)
df = df[["Q1. 姓名","Q2. 请上传文件","Q3. 学工号"]]

def _load_variables_by_name(prompt, variables):
        placeholders = re.findall(r'!{(\w+)}!', prompt)
        cnt = 0
        for placeholder in placeholders:
            prompt = prompt.replace(f'!{{{placeholder}}}!', variables[cnt])
            cnt += 1
        prompt = prompt.strip()
        return prompt

def construct_containers(df, prompt_path, language):
    with open(prompt_path, 'r', encoding='utf-8') as f:
        prompt = f.read()
    containers = []
    for index, row in df.iterrows():
        variables_dict = {
             "fig_name": "",
             "name": row["Q1. 姓名"],
        }
        variables = [row[column] for column in df.columns]
        prompt = _load_variables_by_name(prompt, variables)
        containers.append(prompt)
    return containers

def main():
     for i in df