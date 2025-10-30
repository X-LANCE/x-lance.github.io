import pandas as pd
import requests
import os
from pypinyin import lazy_pinyin, Style
from tqdm import tqdm

default_pic = "../../assets/img/octocat.png"
eng_alu_format = """<div class="member">
        <a href=""><img src="{pic}" alt="{name}"></a>
        <div style="margin-top: 15px"><b>{name}</b><br><b>{xlanceid}</b></div>
    </div>"""
chi_alu_format = """<div class="member">
        <a href=""><img src="{pic}" alt="{name}"></a>
        <div style="margin-top: 15px"><b>{name}</b><br><b>{xlanceid}</b></div>
    </div>"""
eng_stu_format = """<div class="member">
        <a href=""><img src="{pic}" alt="{name}"></a>
        <div style="margin-top: 15px"><b>{name}</b><br><b>{xlanceid}</b></div>
    </div>"""
chi_stu_format = """<div class="member">
        <a href=""><img src="{pic}" alt="{name}"></a>
        <div style="margin-top: 15px"><b>{name}</b><br><b>{xlanceid}</b></div>
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
        <a href=""><img src="/assets/img/members/student/樊帅.jpg" alt="Shuai Fan"></a>
        <div style="margin-top: 15px"><b>Shuai Fan</b><br><b>185-F</b></div>
    </div>
<div class="member">
        <a href=""><img src="/assets/img/members/student/缪庆亮.jpg" alt="Qingliang Miao"></a>
        <div style="margin-top: 15px"><b>Qingliang Miao</b><br><b></b></div>
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
        <a href=""><img src="/assets/img/members/student/樊帅.jpg" alt="樊帅"></a>
        <div style="margin-top: 15px"><b>樊帅</b><br><b>185-F</b></div>
    </div>
<div class="member">
        <a href=""><img src="/assets/img/members/student/缪庆亮.jpg" alt="缪庆亮"></a>
        <div style="margin-top: 15px"><b>缪庆亮</b><br><b>185-F</b></div>
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
for person in data:
    names.append(person[0])
    eng_names.append(person[1])
    ids.append(person[2])
    degrees.append(person[3])
    pics.append(person[4])
    states.append(person[5])


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
    
    typ = 1  # 学生
    if '离开' in person[5]:
        typ = 0  # 校友
    else:
        if 'M' in degree:
            typ = 2
        if 'P' in degree:
            typ = 3
    
    if typ == 0:
        return typ, eng_stu_format.format(pic=pic, name=eng_name, xlanceid=xlanceid), chi_stu_format.format(pic=pic, name=name, xlanceid=xlanceid)
    else:
        return typ, eng_alu_format.format(pic=pic, name=eng_name, xlanceid=xlanceid), chi_alu_format.format(pic=pic, name=name, xlanceid=xlanceid)


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


def down_pic(url, path, filename,OVER_WRITE_PICS=False):
    if not os.path.exists('..'+path) or OVER_WRITE_PICS:
        try:
            print(f"downloading pic of {filename}")
            response = requests.get(url, stream=True)
            response.raise_for_status()  # 检查请求是否成功
            pic = path + '/' + filename
            with open('..'+pic, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            # print(f"downloaded successfully to {pic}")
            return pic
        
        except Exception as e:
            print(f"下载失败：{str(e)}")
            return None
    else:
        print('..'+path + '/' + filename+' exists')
        return path + '/' + filename


def upd_xlsx():
    # 请先手动将先前已经导入的行删掉
    pic_file = pd.read_excel('./old/picfile-20251010.xlsx')
    OVER_WRITE_PICS=False
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
        if name in names:
            po = names.index(name)
            if eng_name != eng_names[po]:
                print(f"eng_name: {eng_names[po]} != {eng_name}")
                eng_names[po] = eng_name
            if xlanceid != 0:
                ids[po] = xlanceid
            if not pd.isnull(pic):
                pics[po] = down_pic(pic, '/assets/img/members/student', format_filename(name) + '.jpg',OVER_WRITE_PICS=True)
            if degree not in degrees[po]:
                degrees[po] = upd_degree(degree, degrees[po])
            if degrees[po] != xlanceid_degree:
                print(f"xlanceid_degree: {xlanceid_degree} != degree: {degrees[po]}")
                degrees[po] = xlanceid_degree
            states[po] = state
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
            if pd.isnull(pic):
                pics.append(default_pic)
            else:
                pics.append(down_pic(pic, '/assets/img/members/student', format_filename(name) + '.jpg',OVER_WRITE_PICS=True))
        
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
        'state': states
    }
    
    writer = pd.ExcelWriter('final_new.xlsx')
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
    for person in tqdm(data):
        if person[0] == '樊帅':
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
