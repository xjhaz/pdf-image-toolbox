# 🧰 PDF Image Toolbox

A desktop tool built with **PyQt5** and **PyMuPDF (fitz)** for extracting embedded images from PDF files and generating automated configuration files, as well as batch inserting images into PDFs based on configuration or custom parameters.

一个基于 **PyQt5 + PyMuPDF (fitz)** 的桌面工具，用于：

- 从 PDF 文件中**提取嵌入图片并生成自动化配置文件**；
- 根据配置文件或自定义参数**批量将图片插入 PDF**。

---

## ✨ 功能概述

| 功能                     | 说明                                                         |
| ------------------------ | ------------------------------------------------------------ |
| 📤 **从 PDF 提取图片**    | 自动扫描 PDF 内所有嵌入图像，导出为 PNG 并生成 `<PDF名>_config.json` |
| 📥 **批量插入图片**       | 按配置文件中定义的坐标、缩放比例、页码等信息批量插入图片     |
| 🧩 **智能配置生成**       | 自动生成 `<PDF文件名>_config.json`，可复用以恢复版式         |
| 📂 **批量处理目录**       | 支持遍历目录，对多个 PDF 进行批量插图操作                    |
| 🎨 **可视化界面**         | 采用 PyQt5 构建，操作直观，所有路径可视化可修改              |
| 🧱 **透明图标与独立打包** | 内置 `pdf_image_toolbox_icon_transparent.ico`，支持 PyInstaller 打包为 `.exe` |

---

## ⚙️ 环境依赖

请确保 **Python ≥ 3.8**，并安装以下依赖：

```bash
pip install -r requirements.txt
```

或手动安装：

```bash
pip install PyQt5 PyMuPDF Pillow
```

---

## 🚀 启动方式

### 1) 源码运行

```bash
python pdf_image_toolbox.py
```

### 2) 使用已打包的exe文件


在release中下载

---

## 🧭 使用说明

### 模式一：从 PDF 提取图片

1. 选择要处理的 PDF 文件；
2. 软件自动在相同目录下创建 `pic/` 文件夹；
3. 导出所有嵌入图片并生成 `<PDF名>_config.json`；
4. 可修改配置后用于插入模式。

### 模式二：批量插入图片

1. 选择处理目录（可含多个 PDF）；
2. 可加载提取时生成的配置文件；
3. 点击「开始处理」批量插入图片；
4. 输出到自动生成的 `output/` 目录。

---

## 🪶 文件结构

```
📁 pdf-image-toolbox/
├── pdf_image_toolbox.py                     # 主程序
├── pdf_toolbox.ico   # 应用图标
├── requirements.txt                         # 依赖列表
└── README.md
```

---

## 🧩 技术栈

- **PyQt5**：图形界面与交互逻辑  
- **PyMuPDF (fitz)**：PDF 图像提取与插入  
- **Pillow**：图标与图像透明度处理  
- **PyInstaller**：独立 EXE 打包  

---

## 🧰 版本要求

| 依赖包  | 最低版本 |
| ------- | -------- |
| PyQt5   | 5.15.9   |
| PyMuPDF | 1.24.3   |
| Pillow  | 10.0.0   |

---

## 📜 License

This project is licensed under the **MIT License**.  
See the [LICENSE](./LICENSE) file for details.
