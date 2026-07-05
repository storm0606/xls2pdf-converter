# xls2pdf-converter 开发文档

## 项目概述

xls2pdf-converter 是一个 Python 后端服务，用于将 XLS 格式的**巡更系统报表**精确转换为 PDF 文件。它通过 xlrd 解析 Excel BIFF 格式，使用 reportlab 生成 PDF，并利用 fontTools 对嵌入字体做子集化以减小文件体积。

### 项目定位

- **P0 — 核心功能**：XLS→PDF 像素级还原（对齐、边框、字体、背景色、合并单元格）
- **P1 — API 服务**：Flask HTTP API，支持 token 中转与直出两种模式
- **P2 — 部署运维**：systemd 服务，字体目录管理

### 架构

```
上传 XLS → parse_xls() → _collect_chars() → _register_subsets()
                         ↓
                   _draw_cell() × N — 逐单元格绘制（背景、边框、文本）
                         ↓
                   xls_to_pdf() — 分页、输出 PDF
```

---

## 核心模块

### converter.py — 核心转换引擎

#### 数据流

1. **parse_xls(xls_path)** → `(sheets, wb)`
   - 用 `xlrd.open_workbook(xls_path, formatting_info=True)` 打开
   - 遍历每个 sheet，提取：
     - 列宽 `col_widths`（xlrd 单位：1/256 字符宽度）
     - 行高 `row_heights`（xlrd 单位：twips，1 twip = 1/20 pt）
     - 单元格样式 `cell_styles`（字体、对齐、边框、颜色）
     - 合并单元格映射 `merged_map` / `merged_slave`

2. **_register_subsets(sheets)** → `font_map`
   - 收集所有用到的字符 → `_collect_chars()`
   - 对每种字体调用 `font_subsetter.subset_font()` 只保留用到的字符
   - 注册到 reportlab 的 `pdfmetrics`

3. **_draw_cell(c, ri, ci, sheet, col_x, row_y, col_w, row_h, font_map)**
   - 绘制**背景色**（矩形填充）
   - 绘制**边框**（四条边，支持线宽 BORDER_THIN / BORDER_MEDIUM）
   - 绘制**文本**（按对齐参数计算位置）

4. **xls_to_pdf(xls_path, output_path)** → `output_path`
   - 计算最大页宽（支持超宽表格页）
   - 分页逻辑：按 USABLE_H 换页
   - 逐页绘制所有单元格
   - 保存 PDF

#### 对齐常量参考表

**水平对齐（horz_align）：**

| xlrd 值 | 常量名     | 含义    | PDF 实现                                     |
|---------|-----------|---------|---------------------------------------------|
| 0       | general   | 通用    | 数字(ctype=2)右对齐，文本/其他左对齐         |
| 1       | left      | 左对齐  | `tx = x1 + pad`                             |
| 2       | center    | 居中    | `tx = x1 + (w - tw) / 2.0`                  |
| 3       | right     | 右对齐  | `tx = x2 - tw - pad`                        |

**垂直对齐（vert_align）：**

| xlrd 值 | 常量名     | 含义    | PDF 实现                                     |
|---------|-----------|---------|---------------------------------------------|
| 0       | top       | 靠上    | `ty = y1 - font_size - pad`                 |
| 1       | middle    | 居中    | `ty = y2 + (h - font_size) / 2.0`           |
| 2       | bottom    | 靠下    | `ty = y2 + pad`                             |

> **注意**：v4.0 之前水平对齐映射完全颠倒（horz=2 误为右对齐、horz=1 误为居中）。v4.0 已修复。

#### 关键常量

| 常量            | 值               | 说明                                          |
|----------------|------------------|----------------------------------------------|
| COL_UNIT_PT    | 0.0205078125    | 列宽转换系数：xlrd 单位 → pt                  |
| BORDER_THIN    | 0.5              | 细边框线宽（pt）                              |
| BORDER_MEDIUM  | 1.0              | 中等边框线宽（pt）                            |
| MARGIN         | 14               | 页面边距（pt）                                |
| A4_W           | 595.28           | A4 宽度（pt）                                 |
| A4_H           | 841.89           | A4 高度（pt）                                 |
| USABLE_H       | 813.89           | 可用高度 = A4_H - MARGIN × 2                  |

#### COL_UNIT_PT 推导

```
xlrd 列宽单位 = 1/256 字符宽度
Excel 默认字符宽度 ≈ 7 像素 (96 DPI)
1 字符宽度 = 7/96 英寸 = 7/96 × 72 pt = 5.25 pt
1 单位 = 5.25 / 256 ≈ 0.0205078125 pt

COL_UNIT_PT = 7.0 / 256.0 * 72.0 / 96.0  =  0.0205078125
```

#### 默认行高

行高默认值 300 twips = 15 pt（Excel 默认行高）。转换公式：`twips / 20.0 = pt`。

---

### server.py — Flask API 服务

#### 路由

| 方法   | 路径                           | 说明                              |
|--------|-------------------------------|-----------------------------------|
| POST   | `/convert`                    | 上传 XLS，返回 `{token, size}`    |
| GET    | `/download/<token>.pdf`       | 用 token 下载 PDF（5 分钟有效）   |
| POST   | `/convert/direct`             | 上传 XLS，直接返回 PDF 二进制     |
| GET    | `/api/info`                   | 返回服务信息                      |
| GET    | `/health`                     | 健康检查                          |

#### Token 存储机制

使用内存字典 `_export_store` 存储已转换的 PDF 字节。线程锁 `_export_lock` 保证并发安全。每个 token 有效期 5 分钟，超时后自动清理。

---

### font_subsetter.py — 字体子集化

#### 字体查找优先级

1. `fonts/` 目录下的本地字体文件（精确文件名匹配）
2. 系统字体目录中的精确文件名匹配
3. 系统字体目录中的模糊文件名匹配
4. Fallback 到替代字体（中文字体 → WenQuanYi Micro Hei）
5. 已知的 fallback 字体路径

#### 字体名称映射（_FONT_NAME_MAP）

| Excel 字体名     | 文件名           | 实际字体                        |
|-----------------|-----------------|--------------------------------|
| 微软雅黑         | msyh.ttf        | Microsoft YaHei                |
| microsoft yahei | msyh.ttf        | Microsoft YaHei                |
| arial           | arial.ttf       | LiberationSans-Regular（替代）  |
| arial unicode ms| arialuni.ttf    | —                              |
| simsun          | simsun.ttc      | Noto Serif CJK SC（替代）      |
| 宋体             | simsun.ttc      | Noto Serif CJK SC（替代）      |
| simhei          | simhei.ttc      | Noto Serif CJK SC Bold（替代） |
| 黑体             | simhei.ttc      | Noto Serif CJK SC Bold（替代） |

#### 子集化流程

```
subset_font(font_path, chars, font_index=0)
  ├── TTC 文件先提取单个字体到临时 TTF
  ├── fontTools.Subsetter 只保留 chars 中的字符
  ├── 保存子集化后的字体
  └── 返回路径
```

---

## 部署与依赖

### Python 依赖

```
flask>=2.0
xlrd>=1.2.0
reportlab>=4.0
fonttools>=4.0
```

### 启动

```bash
cd /path/to/xls2pdf-converter
python3 server.py
# 监听 0.0.0.0:28891
```

### systemd 服务

```ini
[Unit]
Description=xls2pdf-converter
After=network.target

[Service]
Type=simple
User=storm
WorkingDirectory=/home/storm/.openclaw/workspace/xls2pdf-converter
ExecStart=/usr/bin/python3 server.py
Restart=always

[Install]
WantedBy=multi-user.target
```

### 字体目录

字体文件放在项目 `fonts/` 目录下：

```
fonts/
├── msyh.ttf              # 微软雅黑
├── simsun.ttc            # 宋体（Noto Serif CJK SC 替代）
├── simhei.ttc            # 黑体（Noto Serif CJK Bold 替代）
├── arial.ttf             # Arial（LiberationSans 替代）
├── arialbd.ttf           # Arial Bold（LiberationSans Bold 替代）
├── wqy-microhei.ttc      # WenQuanYi Micro Hei（通用 fallback）
└── LiberationSans-*.ttf  # Liberation Sans（Arial 系列替代）
```

---

## 测试验证方法

### 1. 基础转换测试

```bash
cd /path/to/xls2pdf-converter
python3 -c "
import sys; sys.path.insert(0, '.')
from converter import xls_to_pdf
result = xls_to_pdf('test.xls', '/tmp/output.pdf')
print(f'PDF: {result}')
"
pdfinfo /tmp/output.pdf
pdftotext /tmp/output.pdf - | head -30
```

### 2. API 测试

```bash
# Token 模式
curl -F "file=@test.xls" http://localhost:28891/convert

# 直出模式
curl -F "file=@test.xls" http://localhost:28891/convert/direct -o output.pdf

# 服务信息
curl http://localhost:28891/api/info

# 健康检查
curl http://localhost:28891/health
```

### 3. 验证重点

- [ ] 单元格文字位置与 Excel 一致（对齐映射正确）
- [ ] 合并单元格正常渲染
- [ ] 中文字体正常子集化与嵌入
- [ ] 列宽行高与 Excel 显示一致
- [ ] 分页正确，无内容截断
- [ ] 背景色渲染正确
- [ ] 边框线条位置准确
