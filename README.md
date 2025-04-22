# Excel文件过滤工具

这是一个基于Web的Excel文件过滤工具，可以帮助用户快速筛选Excel文件中的数据。

## 功能特点

- 支持上传Excel文件（.xlsx和.xls格式）
- 自动识别并显示Excel文件的列标题
- 支持选择要过滤的列
- 支持输入多个关键词（用逗号分隔）进行筛选
- 生成过滤后的新Excel文件并提供下载

## 安装步骤

1. 克隆仓库：
```bash
git clone https://github.com/你的用户名/excel-filter-tool.git
cd excel-filter-tool
```

2. 安装依赖：
```bash
pip install -r requirements.txt
```

3. 运行应用：
```bash
python app.py
```

4. 在浏览器中访问：
```
http://localhost:5000
```

## 使用说明

1. 点击"选择Excel文件"按钮上传Excel文件
2. 从下拉菜单中选择要过滤的列
3. 在文本框中输入关键词（多个关键词用逗号分隔）
4. 点击"开始过滤"按钮
5. 等待处理完成后，过滤后的文件将自动下载

## 技术栈

- Python 3.x
- Flask
- Pandas
- OpenPyXL
- HTML/CSS/JavaScript

## 注意事项

- 上传文件大小限制为16MB
- 仅支持.xlsx和.xls格式的Excel文件
- 确保Excel文件的第一行包含列标题 