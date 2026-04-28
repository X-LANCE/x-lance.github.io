import pandas as pd
import requests
import os
import re
import time
from pypinyin import lazy_pinyin, Style
from tqdm import tqdm

default_pic = "/assets/img/octocat.png"


def format_webpage(url):
    """格式化webpage，去掉https://、http://、www.等前缀"""
    if pd.isnull(url) or url == '':
        return None
    url = str(url).strip()
    # 去掉协议前缀
    url = re.sub(r'^https?://', '', url)
    # 去掉www.前缀
    url = re.sub(r'^www\.', '', url)
    # 去掉末尾的斜杠
    url = url.rstrip('/')
    return url


def format_name_with_link(name, webpage):
    """为名字添加超链接，在名字后增加🏠emoji"""
    if webpage and not pd.isnull(webpage):
        # 确保链接有协议前缀
        link = webpage if webpage.startswith('http') else f'https://{webpage}'
        # 在名字后增加可点击的🏠emoji
        return f'<a href="{link}" style="color: inherit; text-decoration: none;">{name}🏠</a>'
    return name


eng_alu_format = """<div class="member">
        <img src="{pic}" alt="{name}">
        <div style="margin-top: 15px"><b>{name_display}</b><br><b>{xlanceid}</b></div>
    </div>"""
chi_alu_format = """<div class="member">
        <img src="{pic}" alt="{name}">
        <div style="margin-top: 15px"><b>{name_display}</b><br><b>{xlanceid}</b></div>
    </div>"""
eng_stu_format = """<div class="member">
        <img src="{pic}" alt="{name}">
        <div style="margin-top: 15px"><b>{name_display}</b><br><b>{xlanceid}</b></div>
    </div>"""
chi_stu_format = """<div class="member">
        <img src="{pic}" alt="{name}">
        <div style="margin-top: 15px"><b>{name_display}</b><br><b>{xlanceid}</b></div>
    </div>"""


# pic_file = pd.read_excel('./picfile-20250417-20240418.xlsx')
# qn_dict = pic_file.values  # questionnaire dict


def format_filename(name):
    return name.replace(' ', '_')


COMPOUND_SURNAMES = ['欧阳', '司马', '诸葛', '西门', '上官', '令狐', '宇文', '慕容', '端木', '东方', '独孤', '尉迟', '司徒', '申屠', '夏侯', '南宫', '澹台', '皇甫', '长孙', '轩辕']


def chi_to_eng(name):
    is_chi_name = False
    for c in name:
        if '\u4e00' <= c <= '\u9fa5':
            is_chi_name = True
    if not is_chi_name:
        return name
    
    surname = None
    for cs in COMPOUND_SURNAMES:
        if name.startswith(cs):
            surname = cs
            given_name = name[len(cs):]
            break
    
    # 处理单姓情况
    if not surname:
        if len(name) < 1:
            return ""  # 处理空输入
        surname = name[0]
        given_name = name[1:]
    
    # 转换拼音
    def convert_part(chars):
        return ''.join(lazy_pinyin(chars, style=Style.NORMAL)).capitalize()
    
    # 处理姓氏和名字
    surname_py = convert_part(surname)
    given_name_py = convert_part(given_name)
    
    return f"{given_name_py} {surname_py}"


eng_alumni_md = """---
page_id: alumni
layout: page
permalink: /members/alumni/
title: 🧑‍🎓Alumni
description: Alumni of X-LANCE
nav: false
---

<style>
.mycontainer {
  width: 100%;
  display: flex;
  flex-wrap: wrap;
  justify-content: center; /* 水平居中所有项 */
  gap: 30px; /* 每项之间的间距 */
  padding: 20px 0;
}

.member {
  text-align: center;
  width: 150px;
}

.member img {
  width: 150px;
  border-radius: 50%;
}
</style>

<div class="mycontainer">"""

chi_alumni_md = """---
page_id: alumni
layout: page
permalink: /members/alumni/
title: 🧑‍🎓校友
description: X-LANCE毕业校友
nav: false
---

<style>
.mycontainer {
  width: 100%;
  display: flex;
  flex-wrap: wrap;
  justify-content: center; /* 水平居中所有项 */
  gap: 30px; /* 每项之间的间距 */
  padding: 20px 0;
}

.member {
  text-align: center;
  width: 150px;
}

.member img {
  width: 150px;
  border-radius: 50%;
}
</style>


<div class="mycontainer">"""

eng_student_md_P = """---
page_id: student
layout: page
permalink: /members/student/
title: 🧑‍💻Students
description: Students of X-LANCE
nav: false
---

<style>
.mycontainer {
  width: 100%;
  display: flex;
  flex-wrap: wrap;
  justify-content: center; /* 水平居中所有项 */
  gap: 30px; /* 每项之间的间距 */
  padding: 20px 0;
}

.member {
  text-align: center;
  width: 150px;
}

.member img {
  width: 150px;
  border-radius: 50%;
}
</style>

<h2 style="text-align: center"> 🌟Postdocs🌟 </h2>
<div class="mycontainer">
<div class="member">
        <img src="/assets/img/members/student/缪庆亮.jpg" alt="Qingliang Miao">
        <div style="margin-top: 15px"><b>Qingliang Miao</b><br><b></b></div>
    </div>
<div class="member">
        <img src="/assets/img/octocat.png" alt="Kunyang Peng">
        <div style="margin-top: 15px"><b>Kunyang Peng</b><br><b></b></div>
    </div>
<div class="member">
        <img src="/assets/img/octocat.png" alt="Beiyi Liu">
        <div style="margin-top: 15px"><b>Beiyi Liu</b><br><b></b></div>
    </div>
</div>

<h2 style="text-align: center"> 🌟PhD Candidates🌟 </h2>
<div class="mycontainer">"""

chi_student_md_P = """---
page_id: student
layout: page
permalink: /members/student/
title: 🧑‍💻学生
description: X-LANCE在读学生
nav: false
---

<style>
.mycontainer {
  width: 100%;
  display: flex;
  flex-wrap: wrap;
  justify-content: center; /* 水平居中所有项 */
  gap: 30px; /* 每项之间的间距 */
  padding: 20px 0;
}

.member {
  text-align: center;
  width: 150px;
}

.member img {
  width: 150px;
  border-radius: 50%;
}
</style>

<h2 style="text-align: center"> 🌟博士后🌟 </h2>
<div class="mycontainer">
<div class="member">
        <img src="/assets/img/members/student/缪庆亮.jpg" alt="缪庆亮">
        <div style="margin-top: 15px"><b>缪庆亮</b><br><b></b></div>
    </div>
<div class="member">
        <img src="/assets/img/octocat.png" alt="彭坤杨">
        <div style="margin-top: 15px"><b>彭坤杨</b><br><b></b></div>
    </div>
<div class="member">
        <img src="/assets/img/octocat.png" alt="刘贝易">
        <div style="margin-top: 15px"><b>刘贝易</b><br><b></b></div>
    </div>
</div>

<h2 style="text-align: center"> 🌟博士研究生🌟 </h2>
<div class="mycontainer">"""

eng_student_md_M = """
<h2 style="text-align: center"> 🌟Master Candidates🌟 </h2>
<div class="mycontainer">"""

chi_student_md_M = """
<h2 style="text-align: center"> 🌟硕士研究生🌟 </h2>
<div class="mycontainer">"""

eng_student_md_U = """
<h2 style="text-align: center"> 🌟Undergraduates🌟 </h2>
<div class="mycontainer">"""

chi_student_md_U = """
<h2 style="text-align: center"> 🌟本科生🌟 </h2>
<div class="mycontainer">"""

file = pd.read_excel('./final.xlsx')
data = file.values
names = []
eng_names = []
ids = []
degrees = []
pics = []
states = []
webpages = []
for person in data:
    names.append(person[0])
    eng_names.append(person[1])
    ids.append(person[2])
    degrees.append(person[3])
    pics.append(person[4])
    states.append(person[5])
    webpages.append(person[6] if len(person) > 6 else None)


def format_single(person):
    name = person[0]
    eng_name = person[1]
    xlanceid = person[2]
    degree = person[3]
    
    if (isinstance(xlanceid, int) or isinstance(xlanceid, float)) and xlanceid > 0:
        xlanceid = int(xlanceid)
        xlanceid = f"{xlanceid:03d}-{degree}"
    else:
        xlanceid = f""
    
    pic = person[4]
    webpage = person[6] if len(person) > 6 else None
    
    # 生成带链接的名字显示（如果有webpage）
    eng_name_display = format_name_with_link(eng_name, webpage)
    chi_name_display = format_name_with_link(name, webpage)
    
    typ = 1  # 学生
    if '离开' in person[5]:
        typ = 0  # 校友
    else:
        if 'M' in degree:
            typ = 2
        if 'P' in degree:
            typ = 3
    
    if typ == 0:
        return typ, eng_stu_format.format(pic=pic, name=eng_name, name_display=eng_name_display, xlanceid=xlanceid), chi_stu_format.format(pic=pic, name=name, name_display=chi_name_display, xlanceid=xlanceid)
    else:
        return typ, eng_alu_format.format(pic=pic, name=eng_name, name_display=eng_name_display, xlanceid=xlanceid), chi_alu_format.format(pic=pic, name=name, name_display=chi_name_display, xlanceid=xlanceid)


def get_degree_state(curstate):
    state = '在读'
    if '毕业' in curstate:
        state = '离开'
    if '本科' in curstate:
        return 'U', state
    if '硕士' in curstate:
        return 'M', state
    if '博士' in curstate:
        return 'P', state


def upd_degree(d1, d2):
    if d1 == 'U':
        return d1 + d2
    if d1 == 'P':
        return d2 + d1
    # d1 = 'M'
    if d2 == 'U':
        return d2 + d1
    return d1 + d2


def down_pic(url, path, filename, OVER_WRITE_PICS=False):
    """
    下载照片，支持后缀逻辑：
    - 初次出现的照片不需要替换
    - 修改时间小于24小时的照片直接覆盖
    - 否则使用 _2, _3 等后缀
    """
    base_name, ext = os.path.splitext(filename)
    full_path = '..' + path + '/' + filename
    
    if not os.path.exists(full_path):
        # 文件不存在，直接下载（初次出现）
        try:
            print(f"downloading pic of {filename} (first time)")
            response = requests.get(url, stream=True)
            response.raise_for_status()
            pic = path + '/' + filename
            with open(full_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            return pic
        except Exception as e:
            print(f"下载失败：{str(e)}")
            return None
    elif OVER_WRITE_PICS:
        # 文件存在且需要覆盖
        try:
            # 检查文件修改时间
            file_mtime = os.path.getmtime(full_path)
            current_time = time.time()
            hours_since_modified = (current_time - file_mtime) / 3600
            
            if hours_since_modified < 24:
                # 修改时间小于24小时，直接覆盖
                print(f"downloading pic of {filename} (overwriting, last modified {hours_since_modified:.1f}h ago)")
                response = requests.get(url, stream=True)
                response.raise_for_status()
                pic = path + '/' + filename
                with open(full_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                return pic
            else:
                # 修改时间超过24小时，使用后缀
                suffix_num = 2
                while True:
                    new_filename = f"{base_name}_{suffix_num}{ext}"
                    new_full_path = '..' + path + '/' + new_filename
                    if not os.path.exists(new_full_path):
                        break
                    # 检查现有后缀文件是否在24小时内修改过
                    existing_mtime = os.path.getmtime(new_full_path)
                    if (current_time - existing_mtime) / 3600 < 24:
                        # 直接覆盖这个后缀文件
                        print(f"downloading pic of {new_filename} (overwriting suffix file)")
                        response = requests.get(url, stream=True)
                        response.raise_for_status()
                        pic = path + '/' + new_filename
                        with open(new_full_path, 'wb') as f:
                            for chunk in response.iter_content(chunk_size=8192):
                                f.write(chunk)
                        return pic
                    suffix_num += 1
                
                # 使用新后缀下载
                print(f"downloading pic of {new_filename} (new suffix)")
                response = requests.get(url, stream=True)
                response.raise_for_status()
                pic = path + '/' + new_filename
                with open(new_full_path, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                return pic
                
        except Exception as e:
            print(f"下载失败：{str(e)}")
            return None
    else:
        print('..' + path + '/' + filename + ' exists')
        return path + '/' + filename


def upd_xlsx():
    # 请先手动将先前已经导入的行删掉
    pic_file = pd.read_excel('./old/实验室网站个人信息维护【2026年2月更新】_答卷数据_2026_04_28_16_34_35.xlsx')
    qn_dict = pic_file.values  # questionnaire dict
    for person in qn_dict:
        name = person[7]
        eng_name = person[8]
        if pd.isnull(eng_name):
            eng_name = chi_to_eng(name)
        xlanceid = person[9]
        xlanceid_degree = person[10]
        degree, state = get_degree_state(person[11])
        pic = person[15]
        webpage_raw = person[16]
        # 格式化 webpage，去掉 https 等前缀
        webpage = format_webpage(webpage_raw)
        
        if name in names:
            po = names.index(name)
            if eng_name != eng_names[po]:
                print(f"eng_name: {eng_names[po]} != {eng_name}")
                eng_names[po] = eng_name
            if xlanceid != 0:
                ids[po] = xlanceid
            if not pd.isnull(pic):
                pics[po] = down_pic(pic, '/assets/img/members/student', format_filename(name) + '.jpg', OVER_WRITE_PICS=True)
            if degree not in degrees[po]:
                degrees[po] = upd_degree(degree, degrees[po])
            if degrees[po] != xlanceid_degree:
                print(f"xlanceid_degree: {xlanceid_degree} != degree: {degrees[po]}")
                degrees[po] = xlanceid_degree
            states[po] = state
            # 更新 webpage（如果问卷中提供了新的）
            if webpage:
                webpages[po] = webpage
        else:
            names.append(name)
            eng_names.append(eng_name)
            if xlanceid == 0:
                ids.append(None)
            else:
                ids.append(xlanceid)
            degrees.append(degree)
            if degrees[-1] != xlanceid_degree:
                print(f"xlanceid_degree: {xlanceid_degree} != degree: {degree}")
                degrees[-1] = xlanceid_degree
            states.append(state)
            # 添加 webpage
            webpages.append(webpage)
            if pd.isnull(pic):
                pics.append(default_pic)
            else:
                pics.append(down_pic(pic, '/assets/img/members/student', format_filename(name) + '.jpg', OVER_WRITE_PICS=True))
        
    # print(len(names))
    # print(len(ids))
    # print(len(degrees))
    # print(len(pics))
    # print(len(states))
    full_dict = {
        'name': names,
        'eng_name': eng_names,
        'xlanceid': ids,
        'degree': degrees,
        'pic': pics,
        'state': states,
        'webpage': webpages
    }
    
    writer = pd.ExcelWriter('final.xlsx')
    # sheetNames = full_dict.keys()  # 获取所有sheet的名称
    sheetNames = ["Sheet1"]
    # sheets是要写入的excel工作簿名称列表
    data = pd.DataFrame(full_dict)
    for sheetName in sheetNames:
        data.to_excel(writer, sheet_name=sheetName, index=False)
    # 保存writer中的数据至excel
    writer.close()


def generate_md():
    global eng_alumni_md
    global chi_alumni_md
    global eng_student_md_P
    global chi_student_md_P
    global eng_student_md_M
    global chi_student_md_M
    global eng_student_md_U
    global chi_student_md_U
    if os.path.exists('final_new.xlsx'):
        file = pd.read_excel('./final_new.xlsx')
    else:
        file = pd.read_excel('./final.xlsx')
    data = file.values
    current_postdoc_names = {'缪庆亮', '彭坤杨', '刘贝易'}
    for person in tqdm(data):
        state = '' if pd.isnull(person[5]) else str(person[5])
        if person[0] in current_postdoc_names and '离开' not in state:
            continue
        typ, eng, chi = format_single(person)  # type, english_description, chinese_description
        if typ == 0:
            eng_alumni_md += '\n' + eng
            chi_alumni_md += '\n' + chi
        else:
            if typ == 1:
                eng_student_md_U += '\n' + eng
                chi_student_md_U += '\n' + chi
            if typ == 2:
                eng_student_md_M += '\n' + eng
                chi_student_md_M += '\n' + chi
            if typ == 3:
                eng_student_md_P += '\n' + eng
                chi_student_md_P += '\n' + chi
    
    eng_student_md_P += '\n</div>'
    chi_student_md_P += '\n</div>'
    eng_student_md_M += '\n</div>'
    chi_student_md_M += '\n</div>'
    eng_student_md_U += '\n</div>'
    chi_student_md_U += '\n</div>'
    eng_student_md = eng_student_md_P + eng_student_md_M + eng_student_md_U
    chi_student_md = chi_student_md_P + chi_student_md_M + chi_student_md_U
    
    with open('../_pages/en/alumni.md', 'w', encoding='utf-8') as f:
        f.write(eng_alumni_md)
    with open('../_pages/zh/alumni.md', 'w', encoding='utf-8') as f:
        f.write(chi_alumni_md)
    with open('../_pages/en/student.md', 'w', encoding='utf-8') as f:
        f.write(eng_student_md)
    with open('../_pages/zh/student.md', 'w', encoding='utf-8') as f:
        f.write(chi_student_md)


if __name__ == '__main__':
    # upd_xlsx()
    generate_md()
