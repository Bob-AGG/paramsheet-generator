# 批量提取历史参数单 — 使用说明

## 一句话
在公司电脑上跑一次，把 NAS 里上千份参数单（PDF/PPT/PPTX）批量读成小 JSON，**原始文件不出内网**。

## 装依赖（一次）
```
pip install pdfplumber pypdf python-pptx
```

## 试跑先看效果（20~30 份）
```
python extract_batch.py --src "Z:\参数单\历史" --out "D:\paramsheet_extract" --limit 30
```

## 全量跑
```
python extract_batch.py --src "Z:\参数单\历史" --out "D:\paramsheet_extract"
```

## 跑到一半断了？
```
python extract_batch.py --src "Z:\参数单\历史" --out "D:\paramsheet_extract" --skip-existing
```

## 输出长什么样
```
D:\paramsheet_extract\
  index.json         汇总清单（总数/成功/失败/每份型号字数）
  report.txt         人看的报告：失败清单、同型号重复
  items\XXX.json     每份参数单一份（文本 + 表格 + 来源）
```

## 怎么交给我
把 `D:\paramsheet_extract` 整个文件夹压成 zip 发我。1000 份大约 **5~15 MB**。

## 常见情况
- **老 .ppt 格式**：本机没装 LibreOffice 会记为失败。装了会自动转 pptx 再读。
- **文件名没型号**：脚本会从正文找，找不到标 `(未识别)`，不影响下游，我这边还能再捞。
- **同型号多份**：`report.txt` 里会列出来，你自己判断是重复还是新老版本。
