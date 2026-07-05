# xls2pdf-converter API 使用说明

## 服务地址

- **地址**：`http://<host>:28891`
- **端口**：28891
- **协议**：HTTP

---

## 接口清单

| 方法   | 路径                           | 说明                              |
|--------|-------------------------------|-----------------------------------|
| POST   | `/convert`                    | 上传 XLS，返回 Token + 大小       |
| POST   | `/convert/direct`             | 上传 XLS，直接返回 PDF 二进制     |
| GET    | `/download/<token>.pdf`       | 用 Token 下载 PDF                  |
| GET    | `/api/info`                   | 返回服务信息                      |
| GET    | `/health`                     | 健康检查                          |

---

## 1. POST /convert — XLS 上传转 PDF（Token 模式）

### 请求格式

`multipart/form-data`

| 字段         | 类型   | 必填 | 默认值   | 说明                |
|-------------|--------|------|---------|---------------------|
| file        | File   | 是   | —       | 要转换的 .xls 文件  |
| route_name  | String | 否   | export  | 下载时的文件名前缀  |

### 响应格式

```json
{
  "token": "a1b2c3d4e5f6...",
  "size": 151106
}
```

| 字段    | 类型   | 说明                         |
|--------|--------|------------------------------|
| token  | String | 文件标识，用于 /download 下载 |
| size   | Int    | PDF 文件大小（字节）          |

### curl 示例

```bash
curl -F "file=@12.9.xls" -F "route_name=report" http://localhost:28891/convert
```

### Python 示例

```python
import requests

url = 'http://localhost:28891/convert'
files = {'file': open('12.9.xls', 'rb')}
data = {'route_name': 'report'}

resp = requests.post(url, files=files, data=data)
result = resp.json()
print(f"Token: {result['token']}, Size: {result['size']}")

# 下载 PDF
pdf_url = f"http://localhost:28891/download/{result['token']}.pdf"
pdf_resp = requests.get(pdf_url)
with open('output.pdf', 'wb') as f:
    f.write(pdf_resp.content)
```

---

## 2. POST /convert/direct — XLS 上传转 PDF（直出模式）

### 请求格式

`multipart/form-data`

| 字段         | 类型   | 必填 | 默认值   | 说明                |
|-------------|--------|------|---------|---------------------|
| file        | File   | 是   | —       | 要转换的 .xls 文件  |
| route_name  | String | 否   | export  | Content-Disposition 文件名  |

### 响应格式

`Content-Type: application/pdf`

直接返回 PDF 文件二进制流。

### curl 示例

```bash
# 下载到文件
curl -F "file=@12.9.xls" -F "route_name=report" \
  http://localhost:28891/convert/direct -o report.pdf

# 浏览器查看（inline）
curl -F "file=@12.9.xls" http://localhost:28891/convert/direct | \
  xdg-open /dev/stdin
```

### Python 示例

```python
import requests

url = 'http://localhost:28891/convert/direct'
files = {'file': open('12.9.xls', 'rb')}

resp = requests.post(url, files=files)
with open('output.pdf', 'wb') as f:
    f.write(resp.content)
```

---

## 3. GET /download/<token>.pdf — 下载 PDF

### 请求格式

URL 路径参数。

| 参数      | 说明                        |
|----------|-----------------------------|
| token    | `/convert` 返回的 token 值  |

### 响应格式

`Content-Type: application/pdf`

直接返回 PDF 文件二进制流。

### curl 示例

```bash
curl http://localhost:28891/download/a1b2c3d4e5f6.pdf -o report.pdf
```

### 注意事项

- Token **有效期 5 分钟**，过期后返回 404
- Token 一次性使用，下载后可再次请求（有效期重新计算）

---

## 4. GET /api/info — 服务信息

### 请求格式

无参数。

### 响应格式

```json
{
  "version": "4.0",
  "supported_formats": ["xls", "xlsx (via xlrd)"],
  "fonts_available": [
    "arial.ttf",
    "arialbd.ttf",
    "msyh.ttf",
    "simhei.ttc",
    "simsun.ttc",
    "wqy-microhei.ttc"
  ],
  "endpoints": [
    {"method": "POST", "path": "/convert", "desc": "上传XLS，返回JSON {token, size}"},
    {"method": "GET",  "path": "/download/<token>.pdf", "desc": "用token下载PDF"},
    {"method": "POST", "path": "/convert/direct", "desc": "上传XLS，直接返回PDF二进制"},
    {"method": "GET",  "path": "/api/info", "desc": "返回服务信息"},
    {"method": "GET",  "path": "/health", "desc": "健康检查"}
  ]
}
```

### curl 示例

```bash
curl http://localhost:28891/api/info | python3 -m json.tool
```

---

## 5. GET /health — 健康检查

### 请求格式

无参数。

### 响应格式

```json
{"status": "ok"}
```

### curl 示例

```bash
curl http://localhost:28891/health
```

---

## 错误码说明

| HTTP 状态码 | 错误信息                        | 说明                     |
|------------|--------------------------------|--------------------------|
| 400        | `{"error": "No file provided"}` | 请求未包含 file 字段     |
| 404        | `{"error": "Token expired or invalid"}` | Token 无效或已过期 |
| 500        | `{"error": "<异常信息>"}`         | 转换过程出错             |

---

## 完整工作流示例

### 批量转换（Python）

```python
import requests
import os

API_BASE = 'http://localhost:28891'

xls_dir = '/path/to/xls/files'
for fname in os.listdir(xls_dir):
    if not fname.endswith('.xls'):
        continue

    xls_path = os.path.join(xls_dir, fname)
    pdf_name = fname.replace('.xls', '.pdf')

    # 直出模式
    resp = requests.post(
        f'{API_BASE}/convert/direct',
        files={'file': open(xls_path, 'rb')},
        data={'route_name': pdf_name.replace('.pdf', '')}
    )
    with open(pdf_name, 'wb') as f:
        f.write(resp.content)

    print(f'{fname} → {pdf_name} ({len(resp.content)} bytes)')
```

### 配合前端（Token 模式）

```javascript
// 上传
const formData = new FormData();
formData.append('file', fileInput.files[0]);
formData.append('route_name', 'report');

const resp = await fetch('http://host:28891/convert', {
  method: 'POST',
  body: formData
});
const { token } = await resp.json();

// 下载
const pdfResp = await fetch(`http://host:28891/download/${token}.pdf`);
const blob = await pdfResp.blob();
const url = URL.createObjectURL(blob);
window.open(url);
```
