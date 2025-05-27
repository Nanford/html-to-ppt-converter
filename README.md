# HTML to PPT Converter


一个强大的在线工具，可以将16:9比例的HTML页面精确转换为PPT文件，支持多种转换模式，保持原始布局和样式。

🚀 **[在线演示](https://your-domain.vercel.app)**

## ✨ 特性

- 🎯 **三种转换模式**
  - **混合模式**：背景使用截图，文字保持可编辑（推荐）
  - **截图模式**：100%视觉还原，但内容不可编辑
  - **元素解析模式**：所有内容转换为PPT元素，完全可编辑

- 🎨 **完美还原**
  - 支持复杂CSS样式（渐变、阴影、动画）
  - 保持16:9画面比例
  - 精确的坐标转换系统

- 🛠️ **易于使用**
  - 纯前端实现，无需服务器
  - 零配置，开箱即用
  - 支持实时预览

- 🔒 **隐私安全**
  - 所有转换在本地完成
  - 不上传任何数据到服务器
  - 开源透明

## 🚀 快速开始

### 在线使用

访问 [https://your-domain.vercel.app](https://your-domain.vercel.app) 即可直接使用。

### 本地部署

1. **克隆仓库**
```bash
git clone https://github.com/Nanford/html-to-ppt-converter.git
cd html-to-ppt-converter
```

2. **安装依赖**（可选，用于开发）
```bash
npm install
```

3. **本地运行**
```bash
# 使用任何静态服务器
npx serve public

# 或使用Python
python -m http.server 8000 --directory public
```

4. **访问应用**
打开浏览器访问 `http://localhost:8000`

## 📦 部署到 Vercel

### 方法一：使用 Vercel CLI

1. **安装 Vercel CLI**
```bash
npm i -g vercel
```

2. **部署**
```bash
vercel
```

3. **按提示操作**
- 选择项目名称
- 确认项目设置
- 等待部署完成

### 方法二：通过 GitHub

1. **Fork 本仓库**

2. **访问 [Vercel Dashboard](https://vercel.com/dashboard)**

3. **导入 GitHub 项目**
- 点击 "New Project"
- 选择你的 GitHub 仓库
- 点击 "Deploy"

4. **自动部署完成**

## 🏗️ 项目结构

```
html-to-ppt-converter/
├── .git/                   # Git版本控制
├── public/                 # 静态文件目录
│   ├── index.html         # 主页面
│   ├── favicon.ico        # 网站图标
│   ├── robots.txt         # 搜索引擎配置
│   └── src/               # 源代码目录
│       ├── css/
│       │   └── styles.css # 样式文件
│       └── js/
│           ├── converter.js # 转换核心逻辑
│           ├── main.js    # 主应用逻辑
│           └── utils.js   # 工具函数
├── node_modules/          # 依赖包目录
├── vercel.json            # Vercel配置
├── package.json           # 项目配置
├── package-lock.json      # 依赖锁定文件
├── README.md              # 项目说明
├── LICENSE                # 许可证
└── .gitignore            # Git忽略文件
```

## 🔧 技术栈

- **前端框架**：原生 JavaScript (ES6+)
- **样式**：CSS3 with CSS Variables
- **转换库**：
  - [html2canvas](https://html2canvas.hertzen.com/) - HTML截图
  - [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) - PPT生成
- **部署**：Vercel

## 📖 使用指南

### 基本使用

1. **粘贴HTML代码**
   - 将完整的HTML代码粘贴到输入框
   - 支持包含CSS和内联样式的HTML

2. **选择转换模式**
   - **混合模式**：适合大多数场景
   - **截图模式**：适合复杂样式
   - **元素模式**：适合需要大量编辑的场景

3. **设置选项**
   - 输入PPT文件名
   - 选择是否保留背景
   - 选择是否提取文本

4. **开始转换**
   - 点击"开始转换"按钮
   - 等待转换完成
   - PPT文件将自动下载

### 高级功能

- **预览功能**：转换前预览HTML效果
- **高质量输出**：勾选"高质量输出"获得更清晰的图片
- **快捷键**：
  - `Ctrl/Cmd + Enter`：开始转换
  - `Ctrl/Cmd + P`：预览HTML

## 🎯 最佳实践

### HTML要求

1. **固定尺寸**
   - 建议使用16:9比例（如1920x1080）
   - 使用绝对定位时注意坐标转换

2. **字体选择**
   - 使用常见字体确保兼容性
   - 避免使用Web字体

3. **图片处理**
   - 使用Base64编码的图片
   - 或确保图片URL可访问

### 样式建议

```css
/* 推荐的容器样式 */
.container {
    width: 1920px;
    height: 1080px;
    position: relative;
}

/* 使用相对单位 */
.element {
    width: 50%;
    font-size: 2em;
}
```

## 🤝 贡献指南

欢迎贡献代码！请遵循以下步骤：

1. Fork 本仓库
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

## 📄 许可证

本项目基于 MIT 许可证开源 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 🙏 致谢

- [html2canvas](https://github.com/niklasvh/html2canvas)
- [PptxGenJS](https://github.com/gitbrent/PptxGenJS)
- 所有贡献者和用户

## 📮 联系方式

- 作者：Nanford
- Email：your.email@example.com
- GitHub：[@yourusername](https://github.com/yourusername)

## 🗺️ 路线图

- [x] 基础转换功能
- [x] 三种转换模式
- [x] 实时预览
- [ ] 批量转换
- [ ] 模板系统
- [ ] API接口
- [ ] 浏览器扩展
- [ ] 更多PPT模板

---

如果这个项目对你有帮助，请给一个 ⭐️ Star！