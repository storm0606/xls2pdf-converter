# xls2pdf-converter 二次开发计划

## 项目路径
- 代码：/home/storm/.openclaw/workspace/xls2pdf-converter/
- 主文件：converter.py（XLS解析+PDF渲染）、server.py（Flask API）、font_subsetter.py（字体子集化）
- 字体目录：fonts/（msyh.ttf, simsun.ttc, arial.ttf, simhei.ttc, wqy-microhei.ttc）

## 需求清单

### P0 — 核心功能修复
1. **修复对齐 bug**：`_draw_cell` 中 horz_align / vert_align 映射完全颠倒
   - 当前：horz=2(center)→右对齐, horz=1(left)→居中 → **完全反了**
   - 正确映射：
     - 0 (general): 数字右对齐，文本左对齐
     - 1 (left): 左对齐（x1 + pad）
     - 2 (center): 居中（x1 + (w - tw) / 2）
     - 3 (right): 右对齐（x2 - tw - pad）
   - 垂直对齐同理修正：
     - 0 (top): ty = y1 - font_size - pad
     - 1 (middle): ty = y2 + (h - font_size) / 2
     - 2 (bottom): ty = y2 + pad

2. **修正单元格高度/宽度**：确保 PDF 中行列尺寸与 Excel 一致
   - 列宽：xlrd 的 colinfo_map 返回单位为 1/256 字符宽度，1字符≈256单位≈6pt，需要验证 COL_UNIT_PT 常量
   - 行高：xlrd 的 rowinfo_map 返回 twips（1/20 pt），当前 /20.0 转换正确
   - 需要对比 Excel 原始尺寸与 PDF 输出，确认缩放系数

3. **字体支持**：已就位
   - 微软雅黑: msyh.ttf ✅
   - 宋体: simsun.ttc (Noto Serif CJK SC替代) ✅
   - Arial: arial.ttf (Liberation Sans替代) ✅
   - 黑体: simhei.ttc (Noto Serif CJK Bold替代) ✅

### P1 — 文档与API
4. **开发文档**：写完整的 DEVELOPMENT.md
   - 项目架构说明
   - 核心模块说明（converter.py / server.py / font_subsetter.py）
   - 数据流：XLS → parse_xls → _register_subsets → _draw_cell → PDF
   - 对齐常量参考表
   - 字体映射配置说明
   - 部署与依赖

5. **开放 API 接口** + 使用说明
   - 当前 API：POST /convert (上传XLS→返回token) + GET /download/<token>.pdf
   - 新增：POST /convert/json (上传XLS→直接返回PDF base64)
   - 新增：GET /api/info (服务信息)
   - 写 API.md 使用说明文档

## 验证标准
- 用 /home/storm/.openclaw/workspace/patrol-server-2.0/template/12.9.xls 作为测试文件
- 对比 Excel 原始文件与 PDF 输出：
  - 所有单元格文字必须在单元格内居中（与 Excel 一致）
  - 列宽行高比例正确
  - 中文字体（宋体、微软雅黑）正常渲染
  - 英文字体（Arial）正常渲染

## 技术栈
- Python 3, xlrd, reportlab, fontTools, Flask
- 服务端口：28891
- systemd 服务：xls2pdf-converter
