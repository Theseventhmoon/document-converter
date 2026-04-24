# 文档转换工具

一个支持多种文档格式转换的在线工具，类似于 https://www.aconvert.com/cn/ebook/。

## 功能特性

- 支持多种文档格式的上传和转换
- 拖拽上传和批量上传功能
- 多语言界面支持（中文、英文）
- 实时转换进度显示
- 转换结果下载
- 安全的文件处理

## 支持的格式

### 输入格式
- PDF, Word (doc, docx), Excel (xls, xlsx), CSV, TXT
- 电子书格式：AZW, CBZ, CBR, CBC, CHM, DJVU, EPUB, FB2, HTML, LIT, LRF, MOBI, ODT, PRC, PDB, PML, RTF, SNB, TCR

### 转换选项
- PDF → DOCX, TXT, HTML
- DOCX → PDF, TXT, HTML
- XLSX → CSV, PDF
- EPUB → PDF, MOBI, TXT
- MOBI → EPUB, PDF
- TXT → PDF, DOCX

## 技术栈

- 前端：HTML5, CSS3, JavaScript
- 后端：Python Flask
- 文档处理：PyPDF2, python-docx
- 安全：python-magic

## 快速开始

### 本地开发

1. **启动后端服务**

```bash
# 进入后端目录
cd backend

# 激活虚拟环境
venv\Scripts\activate  # Windows
# 或
source venv/bin/activate  # Linux/Mac

# 启动服务
python app.py
```

后端服务将运行在 http://localhost:5000

2. **访问前端页面**

直接打开 `frontend/index.html` 文件即可。

### 部署到云服务器

1. **准备服务器**

- 选择云服务器（如AWS EC2、阿里云ECS等）
- 安装Python 3.7+
- 安装依赖：`pip install flask flask-cors python-multipart PyPDF2 python-docx python-magic-bin`

2. **上传代码**

将项目文件上传到服务器。

3. **配置服务**

- 修改 `app.py` 中的 `host` 和 `port` 配置
- 设置环境变量（可选）

4. **启动服务**

```bash
# 后台运行
nohup python app.py > server.log 2>&1 &
```

5. **配置域名和SSL**（可选）

- 绑定域名
- 配置HTTPS

## 安全措施

- 文件类型验证
- 文件大小限制（默认100MB）
- 文件名清理
- 文件类型检测
- 唯一文件名生成

## 性能优化

- 文件上传分块处理
- 转换任务异步执行
- 缓存转换结果
- 定期清理临时文件

## 扩展建议

- 添加用户认证系统
- 支持更多文档格式
- 集成第三方转换API
- 添加文件存储服务（如S3）
- 实现批量转换功能
- 添加转换历史记录

## 许可证

MIT License