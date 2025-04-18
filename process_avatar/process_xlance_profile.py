import pandas as pd
import requests
import os
from pypinyin import lazy_pinyin, Style
from tqdm import tqdm

default_pic = "../../assets/img/octocat.png"
eng_alu_format = """<div>
        <figure align="center">
        <a href=""><img style="border-radius: 50%; width:150px" src="{pic}" alt=""></a>
        <figcaption><b>{name}</b><br><b>{xlanceid}</b></figcaption>
        </figure>
    </div>"""
chi_alu_format = """<div>
        <figure align="center">
        <a href=""><img style="border-radius: 50%; width:150px" src="{pic}" alt=""></a>
        <figcaption><b>{name}</b><br><b>{xlanceid}</b></figcaption>
        </figure>
    </div>"""
eng_stu_format = """<div>
        <figure align="center">
        <a href=""><img style="border-radius: 50%; width:150px" src="{pic}" alt=""></a>
        <figcaption><b>{name}</b><br><b>{xlanceid}</b></figcaption>
        </figure>
    </div>"""
chi_stu_format = """<div>
        <figure align="center">
        <a href=""><img style="border-radius: 50%; width:150px" src="{pic}" alt=""></a>
        <figcaption><b>{name}</b><br><b>{xlanceid}</b></figcaption>
        </figure>
    </div>"""
pic_file = pd.read_excel('./picfile.xlsx')
qn_dict = pic_file.values  # questionnaire dict


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

file = pd.read_excel('./final.xlsx')
data = file.values
names = []
ids = []
degrees = []
pics = []
states = []
for person in data:
    names.append(person[0])
    ids.append(person[1])
    degrees.append(person[2])
    pics.append(person[3])
    states.append(person[4])


def format_single(person):
    name = person[1]
    xlanceid = person[2]
    degree = person[3]
    
    if xlanceid >0:
        xlanceid =int(xlanceid)
        xlanceid=f"{xlanceid:03d}-{degree}"
    else:
        xlanceid=f""
    
    typ = 1  # 学生
    if '离开' in person[5]:
        typ = 0  # 校友
    else:
        if 'M' in degree:
            typ = 2
        if 'P' in degree:
            typ = 3
            
    pic = person[4]
    if typ == 0:
        return typ, eng_stu_format.format(pic=pic, name=chi_to_eng(name), xlanceid=xlanceid), chi_stu_format.format(pic=pic, name=name, xlanceid=xlanceid)
    else:
        return typ, eng_alu_format.format(pic=pic, name=chi_to_eng(name), xlanceid=xlanceid), chi_alu_format.format(pic=pic, name=name, xlanceid=xlanceid)


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


def down_pic(url, path, filename):
    try:
        print(f"downloading pic of {filename}")
        response = requests.get(url, stream=True)
        response.raise_for_status()  # 检查请求是否成功
        pic = path + '/' + filename
        with open(pic, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        # print(f"downloaded successfully to {pic}")
        return '../' + pic
    
    except Exception as e:
        print(f"下载失败：{str(e)}")
        return None


def upd_xlsx():
    for person in qn_dict:
        name = person[7]
        xlanceid = person[9]
        degree, state = get_degree_state(person[10])
        pic = person[13]
        if name in names:
            po = names.index(name)
            if xlanceid != 0:
                ids[po] = xlanceid
            if not pd.isnull(pic):
                pics[po] = down_pic(pic, '../assets/img/members/student', format_filename(name) + '.jpg')
            if degree not in degrees[po]:
                degrees[po] = upd_degree(degree, degrees[po])
            states[po] = state
        else:
            names.append(name)
            if xlanceid == 0:
                ids.append(None)
            else:
                ids.append(xlanceid)
            degrees.append(degree)
            states.append(state)
            if pd.isnull(pic):
                pics.append(default_pic)
            else:
                pics.append(down_pic(pic, '../assets/img/members/student', format_filename(name) + '.jpg'))
    # print(len(names))
    # print(len(ids))
    # print(len(degrees))
    # print(len(pics))
    # print(len(states))
    full_dict = {
        'name': names,
        'xlanceid': ids,
        'degree': degrees,
        'pic': pics,
        'state': states
    }
    
    writer = pd.ExcelWriter('./final_new.xlsx')
    sheetNames = full_dict.keys()  # 获取所有sheet的名称
    # sheets是要写入的excel工作簿名称列表
    data = pd.DataFrame(full_dict)
    for sheetName in sheetNames:
        data.to_excel(writer, sheet_name=sheetName)
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
    file = pd.read_excel('./final_new.xlsx')
    data = file.values
    for person in tqdm(data):
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
    
    with open('../_pages/en/alumni.md', 'w',encoding= 'utf-8') as f:
        f.write(eng_alumni_md)
    with open('../_pages/zh/alumni.md', 'w',encoding= 'utf-8') as f:
        f.write(chi_alumni_md)
    with open('../_pages/en/student.md', 'w',encoding= 'utf-8') as f:
        f.write(eng_student_md)
    with open('../_pages/zh/student.md', 'w',encoding= 'utf-8') as f:
        f.write(chi_student_md)


if __name__ == '__main__':
    # upd_xlsx()
    generate_md()