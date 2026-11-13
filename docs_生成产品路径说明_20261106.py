# -*- coding: utf-8 -*-
"""生成《OEM 参数单生成器 产品路径说明》Word 文档（2026-11-06）"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

BASE = '/Users/bob/.openclaw/workspace/projects/参数单生成器'
OUT = os.path.join(BASE, 'OEM参数单生成器_产品路径说明_20261106.docx')

BLUE = RGBColor(0x1F, 0x5C, 0xA8)
ORANGE = RGBColor(0xD9, 0x6B, 0x1E)
DARK = RGBColor(0x33, 0x33, 0x33)
GREY = RGBColor(0x77, 0x77, 0x77)
RED = RGBColor(0xC0, 0x30, 0x30)
GREEN = RGBColor(0x2E, 0x7D, 0x32)

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


def h1(text):
    h = doc.add_heading(text, level=1)
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after = Pt(8)
    return h


def h2(text):
    h = doc.add_heading(text, level=2)
    h.paragraph_format.space_before = Pt(10)
    h.paragraph_format.space_after = Pt(6)
    return h


def table(headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Light Grid Accent 1'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, htext in enumerate(headers):
        hdr[i].text = ''
        r = hdr[i].paragraphs[0].add_run(htext)
        r.bold = True
        r.font.size = Pt(9.5)
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = ''
            r = cells[i].paragraphs[0].add_run(str(val))
            r.font.size = Pt(9.5)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


# ==================== 封面标题 ====================
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = title.add_run('OEM 参数单生成器')
r.font.size = Pt(24)
r.bold = True
r.font.color.rgb = BLUE

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = sub.add_run('产品路径说明书')
r.font.size = Pt(16)
r.font.color.rgb = DARK

info = doc.add_paragraph()
info.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = info.add_run('版本 v1.0　｜　2026-11-06　｜　状态：路径已定稿，待恢复开发')
r.font.size = Pt(10)
r.font.color.rgb = GREY

doc.add_paragraph()

# ==================== 摘要 ====================
h1('一、摘要')

p('本工具定位已确定为「销售内部使用的参数单生成工具」，不是对外交付产品。'
  '核心目标是把销售每天重复填写参数单的工作自动化：从选型工具选定机型后，一键生成标准格式的 PPT 参数单。', space_after=8)

p('产品路径的六个关键问题（谁用、入口在哪、单台还是批量、型号怎么找、PPT 存不存、中英文怎么选）'
  '已于 2026-11-06 全部拍板，开发按本文档执行，不再反复讨论方向。', space_after=8)

p('当前状态：技术侧已有 5 项改进修好但未上线；产品路径已明晰；下一步先部署已有改动，再补选型工具跳转。',
  bold=True, color=ORANGE, space_after=10)

# ==================== 定位 ====================
h1('二、产品定位')

h2('2.1 用户选择')
p('在「内部工具」与「对外客户交付物」两个方向中，明确选择前者。', space_after=6)

table(
    ['维度', '结论'],
    [
        ['使用者', '销售（内部人员）'],
        ['访问方式', '内部网络直接访问，不需要账号权限'],
        ['优先级', '速度优先 —— 能用比好看重要'],
        ['容错原则', '允许界面粗糙，但数字必须准确'],
        ['典型场景', '销售给客户报价前，按客户需求选好机型，生成标准参数单'],
    ],
    widths=[4.0, 11.0]
)

h2('2.2 为什么不做对外版本')
bullet('2058 个型号数据都在本地库，销售每天填参数单是明确存在的痛点，先解决这个。', '痛点明确：')
bullet('对外交付物涉及权限、访问控制、版式美观、错误零容忍，复杂度高一个量级，等内部用顺了再评估。', '成本差异：')
bullet('真实的用户路径需要先用起来才知道哪里不顺，闭门设计对外版容易白做。', '先跑通再加码：')

# ==================== 六个答案 ====================
h1('三、六个关键问题的结论')

h2('3.1 入口路径：从选型工具一键跳转')
p('销售在选型工具里选定机型后，结果卡片上出现「生成参数单」按钮，点击自动打开参数单页面并预填全部信息。', space_after=6)
table(
    ['项', '方案'],
    [
        ['跳转方式', '新窗口打开，通过 URL 参数传递'],
        ['传递字段', '品牌、型号、功率、电压'],
        ['参数单页面行为', '自动预填表单 + 自动选中型号'],
        ['兜底', '保留独立地址，允许直接打开手动填写'],
    ],
    widths=[4.0, 11.0]
)

h2('3.2 单台 vs 批量：先做单台')
bullet('销售真实场景 99% 是单台报价，批量需求未经过验证。', '原因一：')
bullet('批量涉及多型号并行、图片匹配、错误处理，工作量是单台的数倍。', '原因二：')
p('结论：保持现状，只做单台。后续如出现明确的批量需求再评估。', bold=True, space_after=8)

h2('3.3 型号查找：搜索框 + 品牌筛选')
table(
    ['能力', '状态'],
    [
        ['搜索框聚焦即显示全部型号', '已实现（未部署）'],
        ['型号模糊匹配（输入 AF11 可搜到 AF11D5）', '需确认是否已支持'],
        ['品牌下拉筛选', '已实现'],
        ['分类树', '不做'],
        ['最近使用', '不做（增加状态维护成本，内部工具不需要）'],
    ],
    widths=[8.0, 7.0]
)

h2('3.4 PPT 存档：不存历史、不二次编辑')
bullet('生成后直接下载到本地，服务器不留文件。', '生成即下载：')
bullet('需要修改 → 改表单重新生成，简单直接。', '修改方式：')
bullet('避免服务器磁盘堆积和历史文件清理的负担。', '不存档的理由：')

h2('3.5 中英文版本：手动选择')
bullet('客户注册地不等于使用语言，自动判断容易出错。', '不做自动判断的原因：')
bullet('表单上提供中英文切换，默认中文。', '实现方式：')

h2('3.6 数据准确性（内部工具最重要的一条）')
bullet('字段缺失时留空，不猜测、不填 0。', '原则一：')
bullet('生成前做校验，必填项缺失给警告，但不阻断生成。', '原则二：')
bullet('参数单要发给客户或工厂，一个错数字的代价高于慢十分钟。', '原则三：')

# ==================== 当前状态 ====================
h1('四、当前进度')

h2('4.1 代码与部署状态')
table(
    ['项', '状态', '说明'],
    [
        ['本地代码', '✅ 已提交', '最新提交 9879049（产品路径文档）'],
        ['工作区', '✅ 干净', '无未提交改动'],
        ['服务器服务', '✅ 运行中', '139.224.66.88，paramsheet.service 状态 active'],
        ['服务器版本', '⚠️ 落后', '运行的是 2026-10-26 13:50 版本'],
        ['本地领先内容', '⏳ 未部署', '5 项改进已修好，尚未上线'],
    ],
    widths=[4.0, 3.0, 8.0]
)

h2('4.2 已修好但未上线的 5 项改进')
table(
    ['序号', '改进内容', '价值'],
    [
        ['1', '频率 / 相数 / 功率因数拆成三个独立字段', '原先复合框填不进去，销售填单卡在这里'],
        ['2', '搜索框聚焦即显示全部型号', '原先必须先打字才能看到型号，找不到就不知道有什么'],
        ['3', '新增「压缩比」输入字段', '参数单必填项缺失'],
        ['4', '文件名安全化（型号含 / \\ : 等符号不再报错）', 'AF11D5/1P 这类型号原先直接生成失败'],
        ['5', '同名型号加配置指纹防覆盖', 'AF17D6/1P 两个款清洗后同名，会静默互相覆盖'],
    ],
    widths=[1.5, 8.0, 5.5]
)

# ==================== 待办 ====================
h1('五、后续动作')

p('恢复开发按四个阶段推进，前一阶段完成才进入下一阶段。', space_after=8)

h2('阶段一：部署已有改动（技术债清零）')
p('把本地 5 项改进同步到服务器并重启服务。这一步只涉及传文件，无需改代码。', space_after=4)
p('验收标准：打开页面确认①频率相数PF三字段独立　②搜索框聚焦显示全部型号　③压缩比字段存在',
  color=GREEN, space_after=8)

h2('阶段二：压缩比数据灌库（需用户侧执行）')
p('2058 份原始参数单不在本机，需要用户在公司电脑上运行提取脚本。', space_after=4)
bullet('运行 tools/convert_extract.py（Windows 下有「一键提取.bat」）', '动作一：')
bullet('报出带压缩比字段的记录条数', '动作二：')
bullet('灌进 data/engine_50hz.json 与 engine_60hz.json', '动作三：')
p('⚠️ 此步骤开发侧无法替代，必须用户在公司电脑完成。', bold=True, color=RED, space_after=8)

h2('阶段三：选型工具到参数单的跳转')
p('在选型工具（genset-selector）的结果卡片上增加「生成参数单」按钮，并按 URL 参数传递信息。', space_after=4)
bullet('选型工具侧：增加按钮 + 拼接 URL 参数', '改动一：')
bullet('参数单侧：接收参数 + 自动预填 + 自动选型号', '改动二：')

h2('阶段四：完整验收')
p('从选型工具选一台机型 → 点击生成按钮 → 参数单自动带出全部信息 → 导出 PPT → 核对数字与选型结果一致。',
  space_after=8)

# ==================== 不做清单 ====================
h1('六、明确不做的事')

p('以下内容不在本项目范围内，避免开发过程中范围蔓延。', space_after=6)

table(
    ['不做的事', '原因'],
    [
        ['用户账号 / 权限系统', '内部工具，内网访问即可'],
        ['批量生成', '需求未经验证，工作量数倍'],
        ['PPT 历史存档', '服务器不留文件，避免清理负担'],
        ['线上二次编辑', '改表单重新生成更快'],
        ['型号分类树', '搜索 + 品牌筛选已够'],
        ['中英文自动切换', '自动判断易出错'],
        ['对外客户访问入口', '定位是内部工具'],
    ],
    widths=[6.0, 9.0]
)

# ==================== 附录 ====================
h1('七、关键文件索引')

table(
    ['文件', '作用'],
    [
        ['index.html', '前端单页（表单 + 型号搜索 + 生成）'],
        ['build_pptx.py', 'PPT 生成核心逻辑'],
        ['server.py', 'FastAPI 后端（单端口托管前端 + 生成接口）'],
        ['templates/template_map.zh.json', '中文模板单元格映射表'],
        ['templates/template_map.en.json', '英文模板单元格映射表'],
        ['tools/convert_extract.py', '原始参数单批量提取脚本（Windows 运行）'],
        ['paramsheet.db', '型号数据库（2.8MB）'],
        ['PRODUCT_PATH.md', '本文档的 Markdown 源文件（仓库内）'],
        ['ARCHIVE_20261027.md', '2026-10-27 归档时的技术状态记录'],
    ],
    widths=[6.0, 9.0]
)

doc.add_paragraph()
end = doc.add_paragraph()
end.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = end.add_run('— 文档结束 —')
r.font.size = Pt(9)
r.font.color.rgb = GREY

doc.save(OUT)
print('已生成:', OUT)
