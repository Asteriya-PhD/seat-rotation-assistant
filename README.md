# 座位轮转助手（Seat Rotation Assistant）

一个**完全离线**的本地小工具：导入现有 Excel 座位表（支持合并单元格、走道/讲台/门窗、复杂布局），
按规则轮转座位，再导出保留原格式的 Excel。给班主任个人使用。

## 下载 Windows 免安装版

到本仓库的 **Releases** 页面下载：

- `座位轮转助手_Windows免安装.zip`

解压后双击 **`启动座位轮转助手.bat`** 即可，**不需要安装 Python**。
（首次运行如遇 SmartScreen 提示，点「更多信息 → 仍要运行」。）

## 功能

- 识别任意格式座位表：座位 / 讲台 / 门窗 / 走道 / 排号 / 大组标签 / 课代表区自动识别，可手工修正；
- 轮转规则：前后、左右、前后+左右、整组互换、整组+组内、全班整体、自定义；步长 1~3；大组可固定；
- 组长标记（单元格底色）跟随学生移动，可关闭；
- 表头「座位周期」自动更新；轮换周期可选每 1/2/3/4 周（默认 2 周）；
- 导出文件名智能递进（第 N 轮 → 第 N+1 轮；带日期则替换日期）；
- 导出保留合并单元格、边框、颜色、字体、行高列宽、打印设置、图片等原格式（XML 原位替换）。

## 在 macOS 上使用（源码运行）

需要 Python 3（macOS 一般自带）：

```bash
python3 app/server.py
```

或双击仓库里的 `启动.command`。

## 自行构建 Windows 免安装版

Windows 机器上：

```bat
py -3 -m pip install --upgrade pyinstaller
pyinstaller --noconfirm --clean --name SeatRotator --console --paths app --add-data "app/web;app/web" windows_entry.py
```

GitHub Actions 会自动构建：每次 push 到 `main` 或打 `v*` tag 时，
在 Actions 里可下载 `座位轮转助手_Windows免安装` artifact；打 tag 会自动创建 Release。

## 说明

- 仓库内示例座位表已**匿名化**（学生01、学生02…），不含真实学生信息；
- 数据只保存在程序所在文件夹（`数据/`、`导出结果/`），不联网；
- 运行需要 Windows 10/11（64 位）或 macOS；无第三方 Python 依赖。
