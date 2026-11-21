# 安装指南 (详细版)

本指南为不熟悉 Python 或命令行操作的用户设计。我们强烈建议使用 **Conda** 来管理您的运行环境，这可以有效防止与您电脑上的其他软件产生冲突。

## 第一步：安装 Miniconda
Conda 是一个能够为该计算器创建一个“独立运行空间”（环境）的工具。

1.  **下载 Miniconda**: 
    -   [Windows 安装包](https://repo.anaconda.com/miniconda/Miniconda3-latest-Windows-x86_64.exe)
    -   [macOS 安装包](https://repo.anaconda.com/miniconda/Miniconda3-latest-MacOSX-x86_64.pkg)
    -   [Linux 安装包](https://repo.anaconda.com/miniconda/Miniconda3-latest-Linux-x86_64.sh)
2.  **按照向导安装**: 使用默认设置即可。在 Windows 上安装完成后，请在开始菜单中搜索并打开 **"Anaconda Prompt"** 或 **"Miniconda Prompt"**。

## 第二步：下载项目
1.  从 GitHub 下载项目的 ZIP 压缩包。
2.  将 ZIP 文件解压到一个您可以轻松找到的文件夹（例如桌面或文档文件夹）。

## 第三步：配置运行环境
打开您的终端（Windows 用户请务必使用 **Miniconda Prompt**，macOS/Linux 用户请打开终端），然后进入项目所在的文件夹：

```bash
# Windows 示例 (请将路径替换为您解压文件的实际位置)
cd C:\Users\YourName\Desktop\trickcal-team-optimizer
```

进入文件夹后，运行以下**一键安装指令**，它会自动为您创建环境并安装所有依赖：

```bash
conda create -n petdispatch python=3.10 -y && conda activate petdispatch && pip install -r requirements.txt
```

## 第四步：运行计算器
以后每次您想要使用计算器时，只需打开 Prompt/终端并运行：

```bash
# 1. 激活环境 (每次打开窗口只需执行一次)
conda activate petdispatch

# 2. 运行网页界面
streamlit run src/ui/web_gui.py
```

*执行后会自动在您的默认浏览器中打开一个新的标签页，您可以在其中操作计算器。*

---

## 备选方案：命令行版
如果您更喜欢使用纯文字版本，请运行：
```bash
python main.py
```

## 常见问题排查
-   **“找不到命令”**: 请确保您在 Windows 上使用的是 **Miniconda Prompt**，或者在安装 Miniconda 后重新启动了终端。
-   **文件夹路径错误**: 在运行安装指令前，请务必使用 `cd` 命令进入包含 `main.py` 和 `requirements.txt` 的实际文件夹。
-   **国内下载慢**: 如果安装过程中非常缓慢，可以尝试在运行命令前切换 pip 源（如使用清华源）。
    ```bash
    pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple
    ```
