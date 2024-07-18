list_file = "news_list.md"

news_index = 0

# get news list
with open(list_file, 'r', encoding='utf-8') as f:
    news_list = f.read()

news_list = news_list.split("\n# 20")
for news in news_list:
    info = news.strip().split("\n\n")
    assert len(info) == 3
    news_index += 1
    date = "20" + info[0].replace('#','').strip()
    content_en = info[1].strip()
    content_zh = info[2].strip()

    news_en = f"""---
layout: post
date: {date}
inline: true
related_posts: false
---

{content_en}
"""
    news_zh = f"""---
layout: post
date: {date}
inline: true
related_posts: false
---

{content_zh}
"""
    with open(f"./en/announcement_{news_index}.md", 'w', encoding='utf-8') as f:
        f.write(news_en)
    with open(f"./zh/announcement_{news_index}.md", 'w', encoding='utf-8') as f:
        f.write(news_zh)
