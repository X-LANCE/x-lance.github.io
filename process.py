import os
import re
import requests

import pandas as pd
from pypinyin import pinyin, Style

excel_file_path = 'D:/Downloads/实验室信息收集_答卷数据_2024_05_23_18_50_43.xlsx'
name_school_id_path = "C:/Users/W/Documents/姓名_学号.xlsx"
name_id_path = "C:/Users/W/Documents/X-LANCE全体名单-20240314-实验室网站信息.csv"
name_school_id_df = pd.read_excel(name_school_id_path)
name_id_df = pd.read_csv(name_id_path)
storage_path = "assets/img/members/student"
image_list = {}


def chinese_name_to_pinyin(chinese_name):
    if len(chinese_name) < 2:
        return "Invalid Name"

    surname = chinese_name[0]
    given_name = chinese_name[1:]

    surname_pinyin = ''.join(p[0].capitalize() for p in pinyin(surname, style=Style.NORMAL))

    given_name_pinyin = ''.join(p[0] for p in pinyin(given_name, style=Style.NORMAL)).capitalize()

    return f"{given_name_pinyin} {surname_pinyin}"


def _load_variables_by_name(prompt, variables):
        placeholders = re.findall(r'!{(\w+)}!', prompt)
        cnt = 0
        for placeholder in placeholders:
            prompt = prompt.replace(f'!{{{placeholder}}}!', variables[placeholder])
            cnt += 1
        prompt = prompt.strip()
        return prompt


def construct_containers(df, prompt_path, language):
    with open(prompt_path, 'r', encoding='utf-8') as f:
        prompt = f.read()
    containers = []
    for index, row in df.iterrows():
        print(image_list)
        if language == 'zh':
            variable_dict = {
                "fig_name": "../../../" + image_list[row["姓名"]] if row["姓名"] in image_list else "../../../assets/img/octocat.png",
                "name": row["姓名"],
                "id": name_id_df[name_id_df["姓名"] == row["姓名"]]["编号"].values[0] if row["姓名"] in name_id_df["姓名"].values else "",
            }
        elif language == 'en':
            variable_dict = {
                "fig_name": "../../" + image_list[row["姓名"]] if row["姓名"] in image_list else "../../assets/img/octocat.png",
                "name": chinese_name_to_pinyin(row["姓名"]),
                "id": name_id_df[name_id_df["姓名"] == row["姓名"]]["编号"].values[0] if row["姓名"] in name_id_df["姓名"].values else "",
            }
        prompt_processed = _load_variables_by_name(prompt, variable_dict)
        containers.append(prompt_processed)
    return containers


def process_image():
    for root, dirs, files in os.walk(storage_path):
        for file in files:
            image_list[file.split('.')[0]] = os.path.join(root,file)
            if '.' in file.split('_')[0]:
                continue
            file_path = os.path.join(root, file.split('_')[0] + "." + file.split('.')[-1])
            os.rename(os.path.join(root, file), file_path)
            image_list[file.split('_')[0]] = file_path


def construct_page(phd_containers, master_containers, undergraduate_containers, language):
    if language == 'zh':
        page_path = "student_zh_template.txt"
    elif language == 'en':
        page_path = "student_en_template.txt"
    with open(page_path, 'r', encoding='utf-8') as f:
        prompt = f.read()
    variable_dict = {
        "phd_containers": "\n".join(phd_containers),
        "master_containers": "\n".join(master_containers),
        "undergraduate_containers": "\n".join(undergraduate_containers),
        "postdoc_containers": "",
    }
    prompt = _load_variables_by_name(prompt, variable_dict)

    return prompt



def main():
    process_image()
    phd_df = name_school_id_df[name_school_id_df["学号"] < 100000000000]
    phd_containers_zh = construct_containers(phd_df, 'container_template.txt', 'zh')
    phd_containers_en = construct_containers(phd_df, 'container_template.txt', 'en')
    master_df = name_school_id_df[(name_school_id_df["学号"] < 500000000000) & (name_school_id_df["学号"] >= 100000000000)]
    master_containers_zh = construct_containers(master_df, 'container_template.txt', 'zh')
    master_containers_en = construct_containers(master_df, 'container_template.txt', 'en')
    undergraduate_df = name_school_id_df[name_school_id_df["学号"] >= 500000000000]
    undergraduate_containers_zh = construct_containers(undergraduate_df, 'container_template.txt', 'zh')
    undergraduate_containers_en = construct_containers(undergraduate_df, 'container_template.txt', 'en')
    page_zh = construct_page(phd_containers_zh, master_containers_zh, undergraduate_containers_zh, 'zh')
    page_en = construct_page(phd_containers_en, master_containers_en, undergraduate_containers_en, 'en')
    with open('student_zh.md', 'w', encoding='utf-8') as f:
        f.write(page_zh)
    with open('student_en.md', 'w', encoding='utf-8') as f:
        f.write(page_en)
    
     


if __name__ == '__main__':
    main()