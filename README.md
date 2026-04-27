# QCP质量计划 → CNPE Excel 转换工具

**访问密码：** `KSB2026`

## 功能说明

将 QCP 质量计划 PDF 自动识别、提取工序数据，填充到 CNPE 格式 Excel 模板。

- 19位编码自动识别
- 数字原生 PDF（pdfplumber） + 扫描件（MiniMax OCR）
- 密码保护访问
- 工序数据自动映射到 CNPE 模板

## 部署到 Vercel

```bash
npm i -g vercel
cd qcp-to-cnpe-web
vercel --yes
```

## 环境变量

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `QCP_PASSWORD` | `KSB2026` | 访问密码，部署前请修改 |
| `SECRET_KEY` | `qcp-cnpe-secret-key-2026` | Flask会话密钥 |

## API

- `POST /login` — 密码验证，请求体 `{"password": "xxx"}`
- `POST /upload` — 上传 PDF 转换，需先登录
- `GET /health` — 健康检查
