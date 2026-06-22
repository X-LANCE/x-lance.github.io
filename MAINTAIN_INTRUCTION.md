# 网站维护与更新流程

## 维护原则

- 官网以 `x-lance.github.io/` 为当前维护目录。
- 常规更新应先修改数据源或维护表，再重新生成对应的静态数据文件。
- 发布前需要本地检查页面是否能正常打开、语言切换是否正常、图片与链接是否可访问。
- 涉及人员、论文、新闻、招生、公众号迁移等内容时，先由宣传委员收集材料，再请相关老师或负责人确认后发布。

## 目录说明

- `index.html`：首页。
- `news.html`：新闻与动态。
- `papers.html`：论文列表。
- `members.html`、`faculty.html`、`students.html`：成员相关页面。
- `repositories.html`：项目与代码仓库。
- `videos.html`：Xmart / Bilibili 视频内容。
- `join.html`：招生与加入我们。
- `research.html`：兼容入口，会跳转到首页研究方向部分。
- `css/style.css`：页面样式。
- `js/runtime.js`：主要页面渲染逻辑。
- `js/data.js`：首页、新闻、项目、视频等静态内容。
- `js/members-data.js`：教师与成员基础数据。
- `assets/data/members.json`：自动生成的成员数据。
- `assets/data/papers.json`：自动生成的论文数据。
- `assets/img/`：原始图片资源。
- `assets/img-webp/`：脚本生成的 WebP 图片资源，页面会优先加载这里的图片。
- `site-build/`：成员与论文数据构建脚本及维护表。

## 常规更新流程

1. 确认更新内容

   收集需要更新的材料，包括新闻标题、日期、正文摘要、图片、论文信息、成员信息、项目链接、视频链接等。缺失信息先留空，不要自行补写无法确认的内容。

2. 修改对应数据源

   - 首页、新闻、研究方向、项目、视频等内容：优先修改 `js/data.js`。
   - 成员信息：优先更新 `site-build/members/final.xlsx`。
   - 论文信息：优先更新 `site-build/papers/paper_list.txt` 或 `site-build/papers/papers.xlsx`。
   - 页面结构或交互：修改对应 HTML 文件和 `js/runtime.js`。
   - 视觉样式：修改 `css/style.css`。

3. 重新生成数据

   运行：

   ```bash
   /home/james/miniconda3/envs/BASE3.14/bin/python site-build/papers/build_papers_excel.py
   /home/james/miniconda3/envs/BASE3.14/bin/python site-build/build_site_json.py
   /home/james/miniconda3/envs/BASE3.14/bin/python site-build/convert_images_to_webp.py --clean --force
   ```

   第一条命令会从论文列表生成或更新 `site-build/papers/papers.xlsx`。第二条命令会生成 `assets/data/members.json`、`assets/data/papers.json`，并同步更新浏览器直开时使用的 JS 兜底数据。第三条命令会把 `assets/img/` 中的栅格图片转换到 `assets/img-webp/`，其中成员头像最长边为 500px，合照和视频封面最大宽度为 1000px，WebP 质量为 75。

4. 本地预览

   在站点目录启动本地服务：

   ```bash
   python3 -m http.server 8000
   ```

   浏览器访问 `http://localhost:8000/`，重点检查首页、成员页、论文页、新闻页、项目页、视频页和加入我们页面。

5. 发布前检查

   - 页面无明显排版错位。
   - 中英文切换正常。
   - 明暗模式切换正常。
   - 图片、二维码、成员头像、视频封面正常显示。
   - 外部链接可以打开。
   - 新增论文、成员、新闻的日期和文字无误。
   - HTML 中 JS 引用的缓存参数保持同一批更新时间戳。

6. 发布与记录

   发布后更新 `README.md` 中的最近更新日期，并在宣传委员维护记录中说明本次修改内容、负责人和确认人。

## 分模块维护说明

### 新闻与首页重点内容

- 适合放置重要获奖、项目进展、代表成果、招生信息、公众号精选内容等。
- 更新入口通常在 `js/data.js`。
- 新闻内容应保持短而明确，避免把公众号长文原样复制到首页。

### 论文

- 代表性论文需要人工筛选，不建议全量堆叠。
- 新论文优先补齐标题、作者、年份、venue、链接、奖项或附件信息。
- 如果论文链接暂未公开，相关字段留空。

### 成员

- 教师、博士后、博士、硕士、本科生、校友信息应从维护表统一更新。
- 新增头像后检查文件名和路径是否与数据表一致。
- 成员身份变化时，同步检查当前成员页和校友列表。

### 项目与视频

- 项目链接优先使用公开 GitHub、Hugging Face、项目主页等稳定链接。
- 视频内容优先使用 Bilibili 稳定链接，并确认封面图是否正常显示。

### 样式与页面结构

- 小范围文字与数据更新不应改动页面结构。
- 视觉或布局调整后需要同时检查桌面端和移动端。
- 若新增页面，需要同步检查导航、页脚和中英文文案。

## 建议维护周期

- 新闻、招生、视频：有重要事件时即时更新。
- 论文：每月或每个重要会议截稿 / 录用周期后检查一次。
- 成员：每学期初和毕业季集中检查一次。
- 首页重点内容：每学期至少检查一次，避免长期停留在旧信息。
- 外部链接与图片：每季度抽查一次。
