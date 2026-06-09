# 📚 作业批改系统

基于 **智谱AI视觉模型** 的实验报告智能批改系统。上传学生Word文档，自动提取学号姓名、分析截图、AI综合评分，导出Excel成绩表。

---

## 🚀 快速启动（无需安装Python）

### 方法一：下载独立 exe（推荐）

项目已配置 **GitHub Actions 自动构建**，每发布一个新版本自动生成 Windows/Mac 的可执行文件：

**下载步骤：**
1. 打开本项目的 GitHub 页面
2. 点击上方的 **Actions** 标签
3. 在左侧点击 **"构建exe安装包"**
4. 在列表中找到最新的运行记录，点击进入
5. 在底部 **Artifacts** 区域下载：
   - `作业批改系统-Windows.zip` → 解压得 `.exe`
   - `作业批改系统-Mac.zip` → 解压得 Mac 可执行文件

**使用方式：**
- 🪟 **Windows**：双击 `作业批改系统.exe`
- 🍎 **Mac**：右键 → 打开（首次会提示未验证，点"仍要打开"）

> 首次启动后会自动打开浏览器 http://localhost:8900

### 方法二：Python 源码运行（开发者）

需要安装 Python 3.9+：

**Mac：**
```bash
# 双击 start.command，或在终端：
cd 作业批改系统
chmod +x start.command && ./start.command
```

**Windows：**
```bash
# 双击 start.bat，或在终端：
cd 作业批改系统
start.bat
```

---

## 📖 使用说明

### 1. 上传文档
- 支持 **选择文件** 或 **选择文件夹** 上传学生实验报告（.docx格式）
- 支持批量上传

### 2. 开始批改
- 点击 **「⚡ 开始批改所有文档」**
- 系统自动解析每份文档，AI逐张分析截图+文字内容
- 评分完成后展示详细结果（每张截图的分项得分 + 文字评价）

### 3. 调整与导出
- 支持手动修改分数和评语
- 点击 **「📊 导出 Excel」** 下载成绩汇总表
- 点击 **「📄 导出 HTML 报告」** 生成可打印的评分报告

---

## 🤖 评分机制

| 维度 | 说明 |
|------|------|
| **截图评分** (0-60分) | AI分析实验截图中命令执行、界面加载、错误报错等情况 |
| **文字评分** (0-40分) | AI分析实验目的、步骤、总结、技术细节的完整度 |
| **总分** (0-100分) | 截图+文字综合评分 |
| **无截图** | AI单独对文字严格评分，并标注缺少截图 |

> 评分模型：智谱AI GLM-4V-Flash（在线免费视觉模型）
>
> ⚠️ 评分需要网络连接（调用智谱AI在线API）
> ⚠️ API Key已内置在代码中，开箱即用

---

## ⚙️ 自定义配置

如需更换API Key或模型，编辑项目中的 `.env` 文件：
```
ZHIPU_API_KEY=你的API密钥
ZHIPU_MODEL=glm-4v-flash    # 免费，也可改为 glm-4v-plus（付费更强）
```

---

## 📁 项目结构

```
作业批改系统/
├── app.py                    # 主程序
├── requirements.txt          # Python依赖
├── .env                      # API配置
├── start.command             # Mac一键启动脚本
├── start.bat                 # Windows一键启动脚本
├── grading_system.spec       # PyInstaller打包配置
├── .github/workflows/build.yml  # GitHub Actions自动构建
├── templates/                # 网页模板
│   ├── index.html
│   ├── grade.html
│   └── report.html
└── static/
    └── style.css
```

---

## 🔧 GitHub Actions 手动触发构建

即使不打标签，也可以手动触发构建：

1. 打开 GitHub 仓库页面
2. 点击 **Actions** 标签
3. 左侧点击 **"构建exe安装包"**
4. 点击右侧 **"Run workflow"** 按钮
5. 选择分支，点击 **"Run workflow"**
6. 等待几分钟，构建完成后在运行记录底部下载

---

## ❓ 常见问题

**Q: 启动后浏览器无法访问？**
A: 尝试 `http://127.0.0.1:8900` 代替 `localhost`。或者检查端口是否被占用。

**Q: 评分时说"API不可用"？**
A: 检查电脑是否联网，或 `.env` 文件中的 API Key 是否正确。

**Q: exe 被 Windows 报毒？**
A: 独立打包的 exe 容易被误报，添加信任即可。也可用源码方式运行。

**Q: 可以换其他AI模型吗？**
A: 可以，修改 `.env` 中的 `ZHIPU_MODEL` 为 `glm-4v-plus` 等其他智谱模型。
