# 批量提取历史参数单 — 傻瓜操作说明（Windows）

**一句话：** 双击一个 .bat，跑完把输出文件夹打包发我。

原始文件不出你的电脑，脚不联网、不调 AI。

---

## 第一次使用（只做一次）

### 1. 装 Python（如果电脑上已有就跳过）

1. 打开 <https://www.python.org/downloads/> ，点黄色大按钮下载
2. 双击安装，**第一页最下面一定要勾上** `Add python.exe to PATH`
3. 装完点 Close

验证：按 `Win + R`，输入 `cmd` 回车，在黑窗口里输入：

```
python --version
```

出现 `Python 3.x.x` 就 OK。（如果提示「不是内部或外部命令」，重装一次，记得勾 PATH）

### 2. 装依赖（一次就够）

在同一个黑窗口里粘贴这一行，回车，等它跑完（1~3 分钟）：

```
pip install pdfplumber pypdf python-pptx
```

### 3. 解压脚本

把 `extract_batch.py` 和 `一键提取.bat` 放到同一个文件夹，比如 `D:\参数单提取\`

---

## 每次使用

### 方式 A：双击（推荐）

在 `一键提取.bat` 上点**右键 → 编辑**，把里面两行路径改成你的实际路径：

```bat
set SRC=Z:\参数单\康明斯
set OUT=D:\参数单提取\输出
```

保存后**双击** `一键提取.bat`。跑完窗口会停住，看一眼最后几行。

### 方式 B：命令行

```bat
cd /d D:\参数单提取
python extract_batch.py --src "Z:\参数单\康明斯" --out "D:\参数单提取\输出"
```

---

## 建议分两步跑

### 第一步：先跑 30 份试跑（1 分钟）

```bat
python extract_batch.py --src "Z:\参数单\康明斯" --out "D:\试跑30" --limit 30
```

跑完把这俩发我：`D:\试跑30\report.txt` + `D:\试跑30\index.json`

我看过提取质量再让你跑全量。

### 第二步：全量

```bat
python extract_batch.py --src "Z:\参数单\康明斯" --out "D:\全量"
```

按 0.1~0.5 秒/份算，1000 份大概 3~8 分钟。

**中途断了？** 加一个参数，跳过已跑过的：

```bat
python extract_batch.py --src "Z:\参数单\康明斯" --out "D:\全量" --skip-existing
```

---

## 输出是什么样的

```
D:\全量\
  index.json        汇总清单（每份的型号/字数/表格数/是否成功）
  report.txt        人看的报告：失败清单、同型号重复
  items\
    CU110D5-6BT5.9-G2.json
    CU1375D5-KTA50-G3.json
    ...
```

每份 JSON 约 5~15KB。1000 份大概 **10~30MB**，打包成 zip 微信能直接发。

---

## 怎么交给我

把输出文件夹整个右键 → 压缩成 zip → 发我。或者只发 `items` 文件夹也行。

---

## 常见问题

**Q: 提示「老 .ppt 格式，且本机未装 LibreOffice」**
装个 LibreOffice（免费）：<https://www.libreoffice.org/download/>，装完重跑一次即可，脚本会自动转格式再读。

**Q: report.txt 里一堆「(未识别)」型号**
正常，说明文件名和正文都没认出型号。不影响提取，我这边能再捞。

**Q: 跑得慢**
NAS 网络盘读文件慢。可以先把文件夹整个复制到本地 D 盘再跑。

**Q: 提示「pip 不是内部或外部命令」**
`python -m pip install pdfplumber pypdf python-pptx`（用这一行代替）

**Q: 扫描件 PDF 提取不出来**
扫描件是图片，没文字层，脚本读不出来（report.txt 里会标 `!!` 失败）。这类需要 OCR，先不用管，跑完报告里看有多少再说。

---

## 2026-10-20 新版：多线程 + 跳过 PDF（提速用）

原来的版本一个文件一个文件串行解析，NAS 网络盘上特别慢。新版加了两项：

- `--pptx-only` 只解析 PPT/PPTX，跳过 PDF（PDF 表格会被压成文字流，本来就只当兜底）
- `--jobs 8` 8 个线程同时跑（机械硬盘建议 2~4，SSD 建议 8~16）

`run_me.bat` 里已经默认开了这两项。**把源目录改成你复制到本地 D 盘的路径**即可。

`--skip-existing` 现在是真实现了：输出目录里已有同名结果的会自动跳过，跟前一次跑完的可以无缝接着跑。

输出格式和旧版完全一致（`items/*.json`、`index.json`、`report.txt`），可以混着用。
