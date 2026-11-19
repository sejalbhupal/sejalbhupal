from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt

md = r'''<h2 align="left">👋 About Me:</h2>

Computer Engineering undergraduate at **St. Vincent Pallotti College of Engineering & Technology, Nagpur (2023–2027)**.

Interested in **Data Analytics, Business Intelligence, and Software Development**. I enjoy working with data, building dashboards, and developing practical software solutions.

- 📊 Exploring data analysis, visualization, and business intelligence.
- 🐍 Working with Python, Pandas, SQL, and Power BI.
- 💻 Building programming and problem-solving skills with C++ and DSA.
- 🧠 Interested in practical applications of analytics and software engineering.
- ☁️ Learning cloud technologies and AWS fundamentals.
- 📩 Contact: [sejalbhupal15@gmail.com](mailto:sejalbhupal15@gmail.com)

<h2 align="left">🌐 Socials:</h2>

<p align="left">
<a href="https://www.linkedin.com/in/sejal-bhupal"><img src="https://img.shields.io/badge/LinkedIn-Connect-0A66C2?style=flat-square&logo=linkedin&logoColor=white" /></a>
<a href="https://github.com/sejalbhupal"><img src="https://img.shields.io/badge/GitHub-Follow-181717?style=flat-square&logo=github&logoColor=white" /></a>
<a href="https://leetcode.com/u/sejal_bhupal"><img src="https://img.shields.io/badge/LeetCode-Practice-FFA116?style=flat-square&logo=leetcode&logoColor=black" /></a>
<a href="mailto:sejalbhupal15@gmail.com"><img src="https://img.shields.io/badge/Email-Contact-EA4335?style=flat-square&logo=gmail&logoColor=white" /></a>
</p>

<h2 align="left">💻 Tech Stack:</h2>

<p align="left">
<img src="https://img.shields.io/badge/C%2B%2B-00599C?style=for-the-badge&logo=cplusplus&logoColor=white" />
<img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white" />
<img src="https://img.shields.io/badge/Java-ED8B00?style=for-the-badge&logo=openjdk&logoColor=white" />
<img src="https://img.shields.io/badge/SQL-4479A1?style=for-the-badge&logo=postgresql&logoColor=white" />
<img src="https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white" />
<img src="https://img.shields.io/badge/CSS3-1572B6?style=for-the-badge&logo=css3&logoColor=white" />
<img src="https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black" />
<img src="https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" />
<img src="https://img.shields.io/badge/NumPy-013243?style=for-the-badge&logo=numpy&logoColor=white" />
<img src="https://img.shields.io/badge/Matplotlib-11557C?style=for-the-badge&logo=plotly&logoColor=white" />
<img src="https://img.shields.io/badge/Power_BI-F2C811?style=for-the-badge&logo=powerbi&logoColor=black" />
<img src="https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white" />
<img src="https://img.shields.io/badge/React-20232A?style=for-the-badge&logo=react&logoColor=61DAFB" />
<img src="https://img.shields.io/badge/Node.js-339933?style=for-the-badge&logo=nodedotjs&logoColor=white" />
<img src="https://img.shields.io/badge/FastAPI-009688?style=for-the-badge&logo=fastapi&logoColor=white" />
<img src="https://img.shields.io/badge/MySQL-005C84?style=for-the-badge&logo=mysql&logoColor=white" />
<img src="https://img.shields.io/badge/PostgreSQL-316192?style=for-the-badge&logo=postgresql&logoColor=white" />
<img src="https://img.shields.io/badge/Git-F05032?style=for-the-badge&logo=git&logoColor=white" />
<img src="https://img.shields.io/badge/GitHub-181717?style=for-the-badge&logo=github&logoColor=white" />
<img src="https://img.shields.io/badge/AWS-232F3E?style=for-the-badge&logo=amazonaws&logoColor=white" />
</p>

<h2 align="left">🚀 Featured Projects:</h2>

<table>
<tr>
<td width="50%" valign="top">

<h3>🚍 Smart City Transit Optimizer</h3>

**Tech:** C++, Python, Pandas, Power BI

- Transit simulation using core DSA algorithms.
- Processed 500+ trip logs with Pandas.
- Built a dashboard with route-wise KPIs.

<a href="https://github.com/sejalbhupal/SmartCityTransitOptimizer">🔗 View Repository</a>

</td>
<td width="50%" valign="top">

<h3>🏥 Healthcare Analytics Dashboard</h3>

**Tech:** SQL, Power BI, Data Analytics

- Analyzed healthcare operational data.
- Visualized patient flow and waiting times.
- Generated operational KPIs and insights.

<a href="https://github.com/sejalbhupal/healthcare-analytics-dashboard">🔗 View Repository</a>

</td>
</tr>
<tr>
<td width="50%" valign="top">

<h3>📈 Sales Performance & Profitability Intelligence</h3>

**Tech:** Power BI, Data Analysis

- Analyzed 50,000+ sales records.
- Created KPI cards and drill-through filters.
- Explored product and regional performance.

<a href="https://github.com/sejalbhupal/Sales-Performance-and-Profitability-Intelligence">🔗 View Repository</a>

</td>
<td width="50%" valign="top">

<h3>🏙️ Civic AI — Grievance Management</h3>

**Tech:** React, Node.js, PostgreSQL, ML

- Full-stack citizen complaint management system.
- Complaint classification and department assignment.
- Admin review and grievance workflow.

<a href="https://github.com/aryankale316-oss/civic-ai">🔗 View Repository</a>

</td>
</tr>
<tr>
<td width="50%" valign="top">

<h3>🔄 Sorting Visualizer</h3>

**Tech:** HTML, CSS, JavaScript

- Interactive sorting algorithm visualizer.
- Bubble, Selection, Insertion, Merge, Quick, and Heap Sort.
- Real-time color-coded animations.

<a href="https://github.com/sejalbhupal/Sorting_Visualizer">🔗 View Repository</a>

</td>
<td width="50%" valign="top">

<h3>💡 More Projects</h3>

I keep exploring practical ideas in analytics, automation, algorithms, and software development.

<a href="https://github.com/sejalbhupal?tab=repositories">🔗 Explore All Repositories</a>

</td>
</tr>
</table>

<h2 align="left">💼 Experience:</h2>

**Data Analytics Intern — Codec Technologies Pvt. Ltd.**  
*June 2026 · Remote*

- Worked on healthcare predictive analytics.
- Performed exploratory data analysis and feature engineering.
- Developed dashboards to track healthcare KPIs and trends.

**Data Analyst Intern — NullClass**  
*May 2025 – June 2025 · Remote*

- Built a Twitter Analytics Dashboard using Python, Pandas, and Power BI.
- Analyzed user engagement, tweet activity, and trends.
- Worked on data cleaning and reporting workflows.

<h2 align="left">📜 Certifications & Achievements:</h2>

- AWS Cloud Practitioner course — Udemy
- Data Science & Analytics — HP
- Introduction to Data Science — Infosys Springboard
- Python for Data Analysis — Udemy
- AWS APAC Solutions Architecture Job Simulation — Forage
- Solved **150+ coding problems** on LeetCode and GeeksforGeeks.
- Secured **2nd Rank** in Bidcas, a department-level competitive event.
- Participated in **Technex 2025 Hackathon**.

<h2 align="left">📊 GitHub Stats:</h2>

<p align="left">
<img height="180em" src="https://github-readme-stats.vercel.app/api?username=sejalbhupal&show_icons=true&theme=dark&hide_border=false&include_all_commits=true&count_private=true" alt="Sejal's GitHub Stats" />
</p>

<p align="left">
<img height="180em" src="https://github-readme-stats.vercel.app/api/top-langs/?username=sejalbhupal&layout=compact&theme=dark&hide_border=false" alt="Most Used Languages" />
</p>

<p align="left">
<img height="180em" src="https://streak-stats.demolab.com?user=sejalbhupal&theme=dark&hide_border=false" alt="GitHub Streak" />
</p>

<h2 align="left">🏆 GitHub Trophies:</h2>

<p align="left">
<img src="https://github-profile-trophy.vercel.app/?username=sejalbhupal&theme=darkhub&no-frame=false&no-bg=false&margin-w=4" alt="GitHub Trophies" />
</p>

<h2 align="left">🔝 Top Contributed Repo:</h2>

<p align="left">
<img src="https://github-contributor-stats.vercel.app/api?username=sejalbhupal&limit=5&theme=dark&combine_all_yearly_contributions=true" alt="Top Contributed Repositories" />
</p>

<h2 align="left">📈 Contribution Activity:</h2>

<p align="left">
<img src="https://github-readme-activity-graph.vercel.app/graph?username=sejalbhupal&bg_color=0d1117&color=58a6ff&line=1f6feb&point=ffffff&area=true&hide_border=true" width="100%" alt="GitHub Contribution Activity Graph" />
</p>

---

<p align="left">
<img src="https://komarev.com/ghpvc/?username=sejalbhupal&label=Profile%20Views&color=0e75b6&style=flat" alt="Profile Views" />
</p>

<div align="center">

**Thanks for visiting my profile!**  
*Always learning. Always building. Always improving.*

</div>
'''

out = Path("/mnt/data")
md_path = out / "README.md"
docx_path = out / "Sejal_README_Gouri_Style_Projects_Activity.docx"
md_path.write_text(md, encoding="utf-8")

doc = Document()
sec = doc.sections[0]
sec.top_margin = Inches(0.55)
sec.bottom_margin = Inches(0.55)
sec.left_margin = Inches(0.6)
sec.right_margin = Inches(0.6)
doc.add_heading("Sejal Bhupal — GitHub README", 0)
p = doc.add_paragraph()
p.add_run("Updated version: ").bold = True
p.add_run("Gouri-style layout with featured projects, GitHub statistics, trophies, top contributed repositories, and a contribution activity graph.")
doc.add_paragraph("Copy the raw source below into README.md in your profile repository (sejalbhupal/sejalbhupal).")
for line in md.splitlines():
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(line if line else " ")
    r.font.name = "Consolas"
    r.font.size = Pt(7.2)
doc.save(docx_path)

print(f"Created {md_path}")
print(f"Created {docx_path}")
