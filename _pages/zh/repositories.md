---
page_id: repositories
layout: page
permalink: /repositories/
title: 🗃️仓库
description: X-LANCE的部分GitHub项目仓库
nav: true
nav_order: 5
---

[//]: # (## GitHub 主页)

[//]: # ()
[//]: # ({% if site.data.repositories.github_users %})

[//]: # ()
[//]: # (<div class="repositories d-flex flex-wrap flex-md-row flex-column justify-content-between align-items-center">)

[//]: # (  {% for user in site.data.repositories.github_users %})

[//]: # (    {% include repository/repo_user.liquid username=user %})

[//]: # (  {% endfor %})

[//]: # (</div>)

[//]: # ()
[//]: # (---)

[//]: # ()
[//]: # ({% if site.repo_trophies.enabled %})

[//]: # ({% for user in site.data.repositories.github_users %})

[//]: # ({% if site.data.repositories.github_users.size > 1 %})

[//]: # ()
[//]: # (<h4>{{ user }}</h4>)

[//]: # (  {% endif %})

[//]: # (  <div class="repositories d-flex flex-wrap flex-md-row flex-column justify-content-between align-items-center">)

[//]: # (  {% include repository/repo_trophies.liquid username=user %})

[//]: # (  </div>)

[//]: # ()
[//]: # (---)

[//]: # ()
[//]: # ({% endfor %})

[//]: # ({% endif %})

[//]: # ({% endif %})

[//]: # (## GitHub 项目)

{% if site.data.repositories.github_repos %}

<div class="repositories d-flex flex-wrap flex-md-row flex-column justify-content-between align-items-center">
  {% for repo in site.data.repositories.github_repos %}
    {% include repository/repo.liquid repository=repo %}
  {% endfor %}
</div>
{% endif %}
