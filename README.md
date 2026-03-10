# 一键实证分析 Web 应用

基于 Streamlit 构建的实证分析可视化工具，支持多种数据格式，提供完整的回归分析流程。

## 功能特性

- **数据支持**: .dta (Stata)、.xlsx/.xls (Excel)、.csv、.sav (SPSS)
- **变量识别**: 自动识别因变量、自变量、控制变量
- **分析方法**:
  - 描述性统计
  - 基准回归 (OLS + 稳健标准误)
  - 异质性分析 (分组回归)
  - 稳健性检验 (Winsorize、不同标准误)
  - 固定效应模型
  - 诊断检验 (VIF、BP检验、DW检验)
- **可视化**: 交互式图表展示
- **报告生成**: Word 格式分析报告

---

## 快速开始

### 方式一：本地运行

1. **安装依赖**
```bash
cd D:\econometric-web
pip install -r requirements.txt
```

2. **启动应用**
```bash
streamlit run app.py
```

3. **访问应用**

打开浏览器访问 `http://localhost:8501`

---

### 方式二：Docker 部署

1. **构建镜像**
```bash
cd D:\econometric-web
docker build -t econometric-web .
```

2. **运行容器**
```bash
docker run -d -p 8501:8501 --name econometric-app econometric-web
```

3. **访问应用**

打开浏览器访问 `http://localhost:8501`

---

### 方式三：服务器部署

#### Streamlit Cloud (推荐)

1. 将代码推送到 GitHub 仓库
2. 访问 [share.streamlit.io](https://share.streamlit.io)
3. 连接 GitHub 仓库并部署

#### 自建服务器

```bash
# 使用 Docker Compose
docker-compose up -d

# 或使用 nginx 反向代理
# nginx 配置示例:
# server {
#     listen 80;
#     server_name your-domain.com;
#     location / {
#         proxy_pass http://127.0.0.1:8501;
#         proxy_http_version 1.1;
#         proxy_set_header Upgrade $http_upgrade;
#         proxy_set_header Connection "upgrade";
#     }
# }
```

---

## 使用说明

### 1. 上传数据

支持格式：
- Stata 数据文件 (.dta) - **推荐**
- Excel 文件 (.xlsx, .xls)
- CSV 文件 (.csv)
- SPSS 数据文件 (.sav)

### 2. 变量选择

应用会自动识别变量类型：
- **因变量**: 结果变量 (Y)
- **核心自变量**: 处理变量 (X)
- **控制变量**: 协变量
- **固定效应**: 地区、时间等
- **分组变量**: 用于异质性分析

### 3. 分析与下载

点击"开始分析"后，可以：
- 查看交互式图表
- 浏览回归结果表格
- 下载 Word 分析报告

---

## 项目结构

```
D:\econometric-web\
├── app.py              # Streamlit 主应用
├── requirements.txt    # Python 依赖
├── Dockerfile          # Docker 配置
└── README.md           # 使用说明
```

---

## 技术栈

- **前端**: Streamlit
- **数据处理**: Pandas, NumPy
- **统计分析**: Statsmodels
- **可视化**: Plotly
- **报告生成**: python-docx

---

## 注意事项

1. **数据格式**: 推荐使用 Stata (.dta) 格式，支持最好
2. **样本量**: 基准回归建议 N > 100
3. **缺失值**: 检查缺失比例，>30% 建议剔除
4. **中文路径**: 已支持中文文件路径

---

## 更新日志

### v1.0.0 (2024-03-10)
- 初始版本发布
- 支持完整实证分析流程
- Web 可视化界面
- Word 报告生成
- Docker 部署支持