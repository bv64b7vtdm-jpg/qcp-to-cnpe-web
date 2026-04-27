# QCP质量计划 → CNPE Excel 转换工具

将QCP（质量控制计划）PDF文件转换为CNPE系统可导入的Excel格式。

## 功能特性

- **自动PDF类型检测**：识别电子版（文字可提取）和扫描版（图片型）
- **智能工序提取**：从QCP表格中提取工序编号、名称、见证点信息
- **保留隐藏列**：XML级别操作，完整保留原始模板的格式、样式和隐藏列
- **批量填充**：支持自动填充D/F/H/I/K/P/Q/R列

## 使用方法

1. 访问Web界面
2. 填写19位编码（必填）
3. 填写零件号（可选）
4. 填写厂家物项编号（必填）
5. 上传QCP PDF文件
6. 点击"开始转换"下载Excel文件

## 字段填充规则

| 列 | 字段 | 规则 |
|---|---|---|
| D | 排序号 | 10, 20, 30... |
| F | 厂家物项编码 | 用户输入的值 |
| H | 工序编号 | 从PDF提取（如A19.0） |
| I | 工序名称 | 从PDF提取（如原材料检验） |
| K | 选点S | H点/W点/R点/- |
| P | 是否产生报告 | 始终为"否" |
| Q | 是否有选点 | 先决条件检查/质量计划关闭 → N，有见证点 → Y |
| R | 选点A1 | H点/W点/R点/- |

## 本地运行

```bash
cd qcp-to-cnpe-web
pip install -r requirements.txt
python app.py
```

访问 http://localhost:5000

## 部署（Render）

```bash
# 直接推送到Git，Render会自动构建
```

## 技术栈

- Flask：Web框架
- pdfplumber：PDF文字提取
- openpyxl：Excel操作（备用）
- gunicorn：生产服务器

## 项目结构

```
qcp-to-cnpe-web/
├── app.py                    # Flask主程序
├── requirements.txt           # 依赖
├── render.yaml               # 部署配置
├── templates/
│   ├── index.html           # 上传页面
│   └── CNPE_质量计划导入Excel模板.xlsx  # Excel模板
├── static/
│   └── style.css            # 样式
├── utils/
│   ├── pdf_parser.py        # PDF解析
│   └── excel_filler.py      # Excel填充
└── README.md
```