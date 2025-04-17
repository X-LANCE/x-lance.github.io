import pandas as pd
import requests
import os
from pypinyin import lazy_pinyin, Style
from tqdm import tqdm

default_pic = "../../assets/img/octocat.png"
eng_alu_format = """<div>
        <figure align="center">
        <a href=""><img style="border-radius: 50%; width:150px" src="{pic}" alt=""></a>
        <figcaption><b>{name}</b><br><b>{xlanceid}-{degree}</b></figcaption>
        </figure>
    </div>"""
chi_alu_format = """<div>
        <figure align="center">
        <a href=""><img style="border-radius: 50%; width:150px" src="{pic}" alt=""></a>
        <figcaption><b>{name}</b><br><b>{xlanceid}-{degree}</b></figcaption>
        </figure>
    </div>"""
eng_stu_format = """<div>
        <figure align="center">
        <a href=""><img style="border-radius: 50%; width:150px" src="{pic}" alt=""></a>
        <figcaption><b>{name}</b><br><b>{xlanceid}-{degree}</b></figcaption>
        </figure>
    </div>"""
chi_stu_format = """<div>
        <figure align="center">
        <a href=""><img style="border-radius: 50%; width:150px" src="{pic}" alt=""></a>
        <figcaption><b>{name}</b><br><b>{xlanceid}-{degree}</b></figcaption>
        </figure>
    </div>"""
pic_file = pd.read_excel('./picfile.xlsx')
pic_data = pic_file.values


def get_degree(text):
    if '本科' in text:
        return 'U'
    if '硕士' in text:
        return 'M'
    if '博士' in text:
        return 'P'

def upd_degree(name, origin):
    for person in pic_data:
        if person[7] == name:
            current_degree= get_degree(person[10])
            if current_degree not in origin:
                origin += current_degree
    return origin

def default_(name, path):
    for picname in os.listdir(path):
        if picname.split('.')[0] == name:
            return os.path.join(path, picname)
    return default_pic

def down_img(name, path, filename):
    url = default_pic
    for person in pic_data:
        if person[7] == name:
            print(f"found {name}")
            url = person[13]
            if pd.isnull(url):
                return default_(name, path)
    if url == default_pic:
        return default_(name, path)
    try:
        print(f"downloading pic of {name}")
        response = requests.get(url, stream=True)
        response.raise_for_status()  # 检查请求是否成功
        pic = path+"/"+filename
        with open(pic, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        # print(f"downloaded successfully to {pic}")
        return '../' + pic
        
    except Exception as e:
        print(f"下载失败：{str(e)}")
        return None

def format_filename(name):
    return name.replace(' ', '_')

COMPOUND_SURNAMES = [
    '欧阳', '司马', '诸葛', '西门', '上官',
    '令狐', '宇文', '慕容', '端木', '东方',
    '独孤', '尉迟', '司徒', '申屠', '夏侯',
    '南宫', '澹台', '皇甫', '长孙', '轩辕'
]
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

def format_single(person):
    name = person[0]
    # sjtuid = person[8]
    xlanceid = person[2]
    degree = upd_degree(name, person[3])
    typ = 1 # 学生
    if '离开' in person[4]:
        typ = 0 # 校友
    else:
        if 'M' in degree:
            typ = 2
        if 'P' in degree:
            typ = 3
    # degree = get_degree(person[10])
    pic = down_img(name, '../assets/img/members/student', format_filename(name) + '.jpg')
    if typ == 0:
        return typ, eng_stu_format.format(pic = pic, name = chi_to_eng(name), xlanceid = xlanceid, degree = degree), chi_stu_format.format(pic = pic, name = name, xlanceid = xlanceid, degree = degree)
    else:
        return typ, eng_alu_format.format(pic = pic, name = chi_to_eng(name), xlanceid = xlanceid, degree = degree), chi_alu_format.format(pic = pic, name = name, xlanceid = xlanceid, degree = degree)



eng_alumni_md = """---
page_id: alumni
layout: profiles
permalink: /members/alumni/
title: alumni
description: Alumni of X-LANCE
nav: false

profiles:
  # if you want to include more than one profile, just replicate the following block
  # and create one content file for each profile inside _pages/
  - align: right
    image: members/faculty/qym_square.jpg
    content: members/faculty/qianyanmin.md
    image_circular: true # crops the image to make it circular
    more_info: >
      <h3>Professor Yanmin Qian</h3>
      <p>SEIEE 3-501<br>qian-ym@sjtu.edu.cn</p>
---

<style>
.mycontainer {
  width:100%;
  height: auto;
  display: flex; /* 使用flex布局 */
  flex-wrap: wrap; /* 设置子元素自动换行 */
  overflow:auto;
}
.mycontainer div {
  margin: 0 10px;
  float:left;
}
</style>

[//]: # (<h2> 博士后 </h2>)


<div class="mycontainer">"""
chi_alumni_md = """---
page_id: alumni
layout: profiles
permalink: /members/alumni/
title: 校友
description: X-LANCE校友
nav: false

profiles:
  # if you want to include more than one profile, just replicate the following block
  # and create one content file for each profile inside _pages/
  - align: right
    image: members/faculty/qym_square.jpg
    content: members/faculty/qianyanmin.md
    image_circular: true # crops the image to make it circular
    more_info: >
      <h3>钱彦旻 教授</h3>
      <p>电院3号楼501<br>qian-ym@sjtu.edu.cn</p>
---

<style>
.mycontainer {
  width:100%;
  height: auto;
  display: flex; /* 使用flex布局 */
  flex-wrap: wrap; /* 设置子元素自动换行 */
  overflow:auto;
}
.mycontainer div {
  margin: 0 10px;
  float:left;
}
</style>

[//]: # (<h2> 博士后 </h2>)


<div class="mycontainer">"""
eng_student_md_P = """---
page_id: student
layout: page
permalink: /members/student/
title: Students
description: Students of X-LANCE
nav: false
---

<style>
.mycontainer {
  width:100%;
  height: auto;
  display: flex; /* 使用flex布局 */
  flex-wrap: wrap; /* 设置子元素自动换行 */
  overflow:auto;
}
.mycontainer div {
  margin: 0 10px;
  float:left;
}
</style>


[//]: # (<h2> Postdocs </h2>)
<h2> PhD Candidates </h2>
<div class="mycontainer">"""
chi_student_md_P = """---
page_id: student
layout: page
permalink: /members/student/
title: Students
description: Students of X-LANCE
nav: false
---

<style>
.mycontainer {
  width:100%;
  height: auto;
  display: flex; /* 使用flex布局 */
  flex-wrap: wrap; /* 设置子元素自动换行 */
  overflow:auto;
}
.mycontainer div {
  margin: 0 10px;
  float:left;
}
</style>

[//]: # (<h2> 博士后 </h2>)

<h2> 博士研究生 </h2>
<div class="mycontainer">"""
eng_student_md_M = """
<h2> Master Candidates </h2>
<div class="mycontainer">"""
chi_student_md_M = """
<h2> 硕士研究生 </h2>
<div class="mycontainer">"""
eng_student_md_U = """
<h2> Undergraduates </h2>
<div class="mycontainer">"""
chi_student_md_U = """
<h2> 本科生 </h2>
<div class="mycontainer">"""

file = pd.read_excel('./full.xlsx')
data = file.values

for person in tqdm(data):
    if '在职' in person[4]:
        continue
    if pd.isnull(person[3]):
        continue
    typ, eng, chi = format_single(person) # type, english_description, chinese_description
    # print(eng)
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