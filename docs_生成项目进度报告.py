# -*- coding: utf-8 -*-
"""生成《OEM 参数单生成器 项目进度报告》Word 文档"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE = '/Users/bob/.openclaw/workspace/projects/参数单生成器'
OUT = os.path.join(BASE, 'OEM参数单生成器_项目进度报告_20261017.docx')

BLUE = RGBColor(0x1F, 0x5C, 0xA8)
ORANGE = RGBColor(0xD9, 0x6B, 0x1E)
DARK = RGBColor(0x33, 0x33, 0x33)
GREY = RGBColor(0x77, 0x77, 0x77)
RED = RGBColor(0xC0, 0x30, 0x30)

doc = Document()

style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(10.5)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
for s in ['Heading 1', 'Heading 2', 'Heading 3', 'Title']:
    st = doc.styles[s]
    st.font.name = 'Arial'
    st.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    st.font.color.rgb = BLUE if s != 'Heading 3' else DARK

sec = doc.sections[0]
sec.top_margin = Cm(2.2)
sec.bottom_margin = Cm(2.0)
sec.left_margin = Cm(2.4)
sec.right_margin = Cm(2.4)


def p(text='', size=10.5, bold=False, color=None, align=None, space_after=6, indent=None):
    par = doc.add_paragraph()
    run = par.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color
    if align is not None:
        par.alignment = align
    par.paragraph_format.space_after = Pt(space_after)
    par.paragraph_format.line_spacing = 1.35
    if indent is not None:
        par.paragraph_format.left_indent = Cm(indent)
    return par


def bullet(text, bold_head=None):
    par = doc.add_paragraph(style='List Bullet')
    par.paragraph_format.space_after = Pt(3)
    par.paragraph_format.line_spacing = 1.3
    if bold_head:
        r = par.add_run(bold_head)
        r.bold = True
        r.font.size = Pt(10.5)
    r2 = par.add_run(text)
    r2.font.size = Pt(10.5)
    return par


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), hexcolor)
    tcPr.append(shd)


def table(headers, rows, widths=None, head_fill='DCE9F8'):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = ''
        par = c.paragraphs[0]
        run = par.add_run(h)
        run.bold = True
        run.font.size = Pt(9.5)
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        shade(c, head_fill)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ''
            par = cells[i].paragraphs[0]
            run = par.add_run(str(v))
            run.font.size = Pt(9.5)
            par.paragraph_format.space_after = Pt(2)
            par.paragraph_format.line_spacing = 1.15
    if widths:
        for r in t.rows:
            for i, w in enumerate(widths):
                r.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


# ================= 封面 =================
p('OEM 参数单生成器', size=25, bold=True, color=BLUE, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
p('ParamSheet Generator', size=13, bold=True, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
p('项 目 进 度 报 告', size=17, bold=True, color=DARK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
p('Datasheet 一上传，参数单自动出（中英双语 PPTX / PDF）', size=10.5, color=GREY,
  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=22)

table(['项目名称', 'OEM 参数单生成器（ParamSheet Generator V2）'],
      [['报告日期', '2026 年 10 月 17 日'],
       ['开发周期', '2026-09-14 模板测绘 → 2026-10-16 V2 全量开发（单日 11 次提交）'],
       ['项目状态', '✅ V2 已上线运行，核心功能完成，等待用户实测反馈'],
       ['代码仓库', 'github.com/Bob-AGG/paramsheet-generator（main 分支）'],
       ['线上地址', 'https://139.224.66.88.sslip.io/paramsheet/ (末尾斜杠必须带)']],
      widths=[3.2, 13.3])

# ================= 一、项目概述 =================
doc.add_heading('一、项目概述', level=1)

p('1.1 一句话定位', bold=True, space_after=3)
p('把客户/厂家给的发动机或发电机组 Datasheet（PDF、Excel、TXT）上传上去，AI 自动把参数抓出来填进表单，'
  '再一键生成符合公司模板的中英文「参数单」PPT ——排版、表格、图片、油耗、标题全部自动完成，可再转 PDF 发给客户。', indent=0.5)

p('1.2 解决的痛点', bold=True, space_after=3)
bullet('以往做一份 OEM 参数单，要人工从 PDF 里一条条抄几十个参数到 PPT，一份半小时到一小时，且容易抄错。', '人工抄录：')
bullet('错一个 kW/kVA 数字，客户拿去做方案就出问题，返工成本高。', '易出错：')
bullet('国外客户要英文版、国内要中文版，以前要维护两套。', '中英双版：')
bullet('新版方案：上传 → AI 提取 → 人工扫一眼核对 → 出 PPT/PDF，压缩到 3～5 分钟。', '')

# ================= 二、整体完成度 =================
doc.add_heading('二、整体完成度', level=1)
table(['模块', '完成度', '说明'],
      [['PPTX 出件引擎', '95%', '英/中双模板，66 张表 + 独立文本框逐格回填，出件自校验'],
       ['AI 参数提取', '85%', 'DeepSeek 双频结构 + 规则兜底；扫描件/图片型 Datasheet 未支持'],
       ['前端界面', '90%', '上传、表单、频率切换、图片选择、导出全部可用'],
       ['图片处理', '90%', '自动抠白底 + 等比居中，OPEN / SILENT 双图自动匹配'],
       ['部署运维', '95%', 'systemd 常驻 + nginx 子路径 + LibreOffice PDF 转换'],
       ['模板对齐', '90%', '第 3、4 页与用户模板零差异；第 1、2 页待用户最终验收']],
      widths=[3.8, 2.0, 10.7])

# ================= 三、已完成工作 =================
doc.add_heading('三、已完成工作明细', level=1)

p('3.1 出件引擎（build_pptx.py，464 行）', bold=True, space_after=3)
bullet('用 python-pptx 打开公司模板，按坐标映射 JSON（template_map.en.json / .zh.json）把数据写进对应表格单元格和文本框，而不是重新画版式 —— 出来的文件与模板结构、字体、配色完全一致，客户可直接编辑。', '原理：')
bullet('英文模板 4 页共 66 行表格映射（3 张表 + 8 张表 + 2 张表）+ 3 个独立文本框（声压级、频率小格、燃料小格）。', '映射范围：')
bullet('写完后逐格回读，确认数据真的落位；第 3、4 页与模板逐 shape 比对差异 = 0。', '出件自校验：')
bullet('中英双语共用同一套坐标，切换语言即可出英文版/中文版。', '双语：')

p('3.2 AI 参数提取（datasheet_ai.py + datasheet_parser.py）', bold=True, space_after=3)
bullet('PDF/XLSX/XLS/TXT 四种输入，先抽文本再交 DeepSeek 按固定字段表提取，返回 JSON 直接灌进表单。', '流程：')
bullet('一份资料同时含 50Hz / 60Hz 两套数据时分开放（功率、转速、油耗随频率变），界面切换频率各填一套、各出一份。', '50Hz / 60Hz 双频：')
bullet('单位自动归一（mbar→kPa、m³/h→m³/min、温度统一 °C）；「TBD / N.A.」自动过滤；多档取值规则明确。', '数据清洗：')
bullet('字段白名单校验 + 反幻觉提示词，防止 AI 编造 Datasheet 里没有的数字。', '防幻觉：')
bullet('AI 不可用或网络失败时，走正则规则解析器兜底（357 行），保证至少能抽出主要参数。', '规则兜底：')

p('3.3 前端与图片处理', bold=True, space_after=3)
bullet('单文件原生 JS，无框架依赖；表单按「🏭 机组 / 🔧 动力 / ⚡ 电机」三大类分组折叠，56 个字段一眼能找。', '界面：')
bullet('每个型号两张图（型号_OPEN.png 开架 / 型号_SILENT.png 静音），自动匹配 + 手动选择 + 在线上传。', '机组图片：')
bullet('只抠四周连通的纯白背景（不动图内白色），再按目标框等比居中裁切，避免图片被拉变形。', '图片处理：')

p('3.4 部署', bold=True, space_after=3)
bullet('阿里云 139.224.66.88，`/root/paramsheet/`（Python3.11 venv），systemd 常驻 `paramsheet.service` 监听 127.0.0.1:8010。', '服务器：')
bullet('nginx 443 分流 `/paramsheet/` → 8010，与发电机选型 AI 后端（genset-ai）共存同一台机器。', '路由：')
bullet('服务器装 LibreOffice，出件时可同时导出 PPTX 与 PDF。', 'PDF 转换：')

# ================= 四、关键规则 =================
doc.add_heading('四、关键规则与设计决策（已定稿）', level=1)

p('4.1 模板填充规则：「12345」标注法', bold=True, space_after=3)
p('用户在模板用手写标了「12345」的格子 = 要由数据填充的格子；没标的格子一律保持模板原值不动（不删格、不填数）。', indent=0.5)
table(['页面', '填充规则', '备注'],
      [['第 1 页', '按标注重填（机组 + 动力两部分）', '大标题、型号字、功率、电压、尺寸重量、声压级等'],
       ['第 2 页', '按标注重填（动力 + 电机）', '基础数据、排气、燃油、冷却、电气、进气、发电机规格'],
       ['第 3 页', '完全不填，保持模板原样', '无标注'],
       ['第 4 页', '完全不填，features 整段保留模板', '无标注（曾误填导致 ◆ 符号丢失、文字挤成一团）']],
      widths=[2.3, 6.4, 7.8])

p('4.2 其他决策', bold=True, space_after=3)
bullet('第 2、3、4 页版式必须与用户原模板一模一样，第 1、2 页允许随数据变化。', '版式一致性：')
bullet('油耗四行只填数字（100% ESP / 100% PRP / 75% PRP / 50% PRP），标签模板左列已有，重复会叠字。', '油耗格式：')
bullet('用户原模板 `参数单模板.pptx` 内已含 D17/DEUTZ 示例数据，不是纯底板；实际使用的是 templates/ 下的英文底板。', '模板区分：')

# ================= 五、本日修复的问题 =================
doc.add_heading('五、已修复问题清单（2026-10-16，共 11 次提交）', level=1)

table(['#', '问题现象', '根因', '解决'],
      [['1', '上传 Datasheet 后页面无反应', 'URL 无尾斜杠时相对路径被打到错误前缀', '前端改用绝对 API 前缀'],
       ['2', '上传后数据填不进表单', 'applyFields 不认双频结构', '兼容 freqs/common 结构'],
       ['3', '大标题、型号字、油耗三项填了不生效', '出件只认模板串，从没读用户输入；油耗未组装成多行', '逐项打通并重写标题框定位'],
       ['4', '第 2/3/4 页被误改', '标题替换在每一页都执行', '标题替换只在第 1 页'],
       ['5', '第 4 页 features 文字挤成一团、断行错乱', '模板里「◆ 」单独占一个 run，被数据覆盖吃掉', '保留模板符号前缀 run + 清洗保留行首'],
       ['6', '单位混乱、AI 编造数值', '无单位归一、无白名单', '单位归一 + 反幻觉约束 + 字段白名单'],
       ['7', '未标注格子被清空', '字段全量映射', '按 12345 标注重建映射，未标格保留原值']],
      widths=[0.9, 5.4, 6.2, 4.0])

# ================= 六、当前状态 =================
doc.add_heading('六、当前状态与风险', level=1)

p('6.1 已验证通过', bold=True, space_after=3)
bullet('本机出件逐格对比：第 3、4 页与模板差异 = 0；第 1、2 页只有标注格变化，未标注格原值（Compression ratio、Type Brushless、Coupling Flexible disc、Over speed 2250 RPM、IP rating IP23）全部保留。', '')
bullet('代码已推送 GitHub（commit 7a9773f）+ 上传服务器 + 重启 paramsheet.service。', '')

p('6.2 待验证 / 风险', bold=True, space_after=3)
bullet('用户尚未实测最新版（0f807e3 与 7a9773f 两轮均未反馈）。需强刷 Ctrl/Cmd+Shift+R 并清浏览器本地数据（字段结构已变）。', '⚠️ 用户验收：')
bullet('模板对齐依赖坐标映射 JSON，若客户换模板版本，需要重新测绘映射（template_structure.json 工具已在）。', '模板耦合：')
bullet('扫描件 / 图片型 Datasheet 不支持，需先 OCR。', 'AI 提取局限：')
bullet('与发电机选型服务共用一台 2 核 2G 服务器，大 PPT（中文模板 32MB）转换时内存吃紧，高并发需拆分。', '服务器资源：')
bullet('AI 提取是「尽力而为」，README 已明确提示导出前必须人工核对，避免直接发客户出错。', '数据准确性：')

# ================= 七、待办 =================
doc.add_heading('七、待办事项', level=1)
table(['#', '优先级', '事项', '说明'],
      [['1', 'P0', '用户实测验收最新版', '强刷 + 清缓存，确认第 4 页 ◆ 回来、第 1 页数据正确'],
       ['2', 'P1', '按实测差异继续打磨第 1、2 页', '用「PDF 渲染对比法」逐页挖差异'],
       ['3', 'P1', '补充历史型号数据模板', '已积累 D17、WP715D5、4HTAA4.5-G5E1/G5E3/G5E5、6DTAA8.9-G5E4、6HTAA7.2-G5E1、MTU 18V2000B76、洋马 4TNV98C-IYE 等'],
       ['4', 'P2', '扫描件 OCR 支持', '图片型 Datasheet 目前需人工'],
       ['5', 'P2', '批量出件', '一次上传多份 Datasheet 批量生成'],
       ['6', 'P3', '出件历史记录 / 一键重下', '便于重复发版']],
      widths=[0.9, 1.6, 5.6, 8.4])

# ================= 八、产出与资产 =================
doc.add_heading('八、已产出资产', level=1)
table(['类别', '内容'],
      [['代码', 'index.html（前端 556 行）/ build_pptx.py（出件 464 行）/ datasheet_ai.py / datasheet_parser.py / image_utils.py / server.py'],
       ['模板', 'template_en.pptx、template_zh.pptx（32MB 中文版）+ 两份坐标映射 JSON'],
       ['已生成参数单', 'D17（英/中）、WP715D5-6M33D633E200（英/中）共 4 份样例 PPTX'],
       ['知识沉淀', 'memory/oem-template.md 动力参数提取模板 + 单位换算规则 + 表映射记录'],
       ['同类积累', 'workspace/oem/ 下 15+ 份已提取的发动机参数 Markdown（4HTAA4.5、6DTAA8.9、MTU、洋马、上柴、谨冠等）']],
      widths=[2.6, 13.9])

doc.add_paragraph()
p('—— 报告完 ——', size=9, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)

doc.save(OUT)
print('saved:', OUT, os.path.getsize(OUT), 'bytes')
