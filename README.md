# 参数单生成器 ParamSheet Generator

把发电机组/发动机的 Datasheet 参数，一键生成中英双语「参数单」PPTX / PDF。

## 功能

- **AI 自动提取**：上传 Datasheet（PDF / Excel / TXT）→ DeepSeek 按固定字段表提取 → 自动填表（读不到的留空手填）
- **50Hz / 60Hz 双频**：同一份资料里的两个频率数据分开存，界面切换频率、各填一套、按频率导出
- **中英双语**：共享坐标映射，切换语言即可出英文版 / 中文版参数单
- **机组图片**：每个型号两张图（`型号_OPEN.png` 开架 / `型号_SILENT.png` 静音），自动匹配 + 手动选择 + 上传图片库；自动抠白底（只抠四周连通背景）+ 等比居中裁切
- **导出**：原生可编辑 PPTX + PDF（服务器 LibreOffice 转换）

## 架构

```
index.html           前端（单文件，原生 JS，工具风 UI）
server.py            FastAPI 后端：/api/parse（AI提取）、/api/build（出件）、/api/upload-image、/api/images
datasheet_ai.py      AI 提取器（DeepSeek，双频结构 + 规则兜底）
datasheet_parser.py  规则解析器（PDF/Excel 抽文本 + 正则提取，兜底用）
build_pptx.py        PPTX 出件引擎（逐格回读校验）
image_utils.py       图片处理（抠白底、裁切、匹配）
glossary.json        中英术语对照表（动力/电机/通用）
templates/           模板（template_en.pptx / template_zh.pptx + 坐标映射 json）
```

## 本地运行

```bash
pip install -r requirements.txt
export DEEPSEEK_API_KEY=sk-xxxx        # AI 提取需要
python server.py                        # → http://127.0.0.1:8010
```

## 部署（服务器）

```bash
# 服务器：Python3.11 venv + LibreOffice（PDF 转换）
python3 -m venv venv && ./venv/bin/pip install -r requirements.txt
systemctl start paramsheet              # uvicorn 127.0.0.1:8010
```

## 字段表（24 字段，前端表单一致）

发动机信息 / 油耗 / 基础数据 / 排气 / 燃油 / 冷却 / 电气 / 进气 / 功率 / 尺寸 / 发电机规格 / 配置项。

其中 **随频率变化的字段**：`freq, prp_kva, eps_kva, prp_kw, eps_kw, rpm, fc_esp, fc_prp, fc_75, fc_50`（放 `freqs[频率]`），其余放 `common`。

## 数据结构

```json
{
  "model": "4HTAA4.5-G5E1",
  "curFreq": "50Hz",
  "freqs": {
    "50Hz": { "prp_kva": "50", "rpm": "1500rpm", "fc_prp": "12.5" },
    "60Hz": { "prp_kva": "57", "rpm": "1800rpm", "fc_prp": "14.8" }
  },
  "common": { "e_brand": "...", "displ": "4.54L", "e_cool": "Water-cooled" }
}
```

## 注意

- AI 提取为「尽力而为」，**导出前务必人工核对**
- DeepSeek key 只存服务器 `.env`（600 权限），**不进仓库**
- 扫描件/图片版 Datasheet 暂不支持（需 OCR）
