# ppt time 1.0.8

ppt time 是一款给讲台电脑使用的会议放映软件。默认在放映屏幕顶部显示倒计时和页码，如 `04:59 │ 2/18`。操作员可在本机或局域网浏览器中管理议程、投放大会主画面、切换串场图片、接收文稿及控制翻页。

## 首次使用

1. **Windows x64 安装版**。保存文稿，退出正在使用的 WPS、PowerPoint 和旧版 ppt time。双击 `ppt-time-Setup.exe`。安装窗口分别提供「启用 WPS 演示页码联动」和「启用 PowerPoint 页码联动」，勾选你使用的演示软件，然后从开始菜单打开 **ppt time**。安装位置为当前用户的 `%LOCALAPPDATA%\Programs\ppt time`。
2. **Windows x64 便携版**：完整解压 `ppt-time-Windows-x64.zip`，双击 `ppt time.exe`。同目录的 `_internal` 包含随程序分发的运行组件。
3. 右键浮窗或系统托盘图标，打开 **计时器设置 → 流程配置**。初始第一环节为 5 分钟，第二环节为 2 分钟。选择 **预览放映计时条（8 秒）** 可先确认演示屏和字号。预览示例在放映开始后会换成现场数据。
4. 双击 PPT 文件，在 PowerPoint 中按 `F5`。ppt time 可以先开，也可以在放映开始后打开。它会连接当前放映窗口并读取页码。也可在托盘菜单选 **打开放映文件…**。默认识别到放映后自动开始计时，按 `Esc` 结束放映后自动暂停。

## 演示格式与页码

Windows PowerPoint 的页码显示由 **计时器设置 → 放映显示 → 显示页码** 控制。双击文稿按 `F5` 后，ppt time 会从已打开的放映读取页码。安装时启用联动可让 PowerPoint 直接报告当前页和总数。便携版或曾手动移除联动的用户，可右键浮窗选择 **PowerPoint 页码联动 → 启用页码加载项**，保存文稿并完全退出、重新打开 PowerPoint 后生效。菜单里的 **查看联动状态** 可以确认 PowerPoint 已连接。直接打开文稿后仍只见倒计时，可选择 **复制页码诊断**，查看每个候选窗口的匹配结果。

**用 WPS 演示放映**时，分别打开 ppt time 和 WPS 中的 `.pptx`、`.ppt` 或 `.dps` 文稿，打开 **显示页码**，在 WPS 里按 `F5`。即使 WPS 返回的放映窗口句柄为零，ppt time 也会检查实际的放映窗口。后台进程和普通编辑窗口可以继续打开。多个真实放映同时存在时，程序会核对文稿归属，避免在投影上显示另一份文稿的页码。

如果按 `F5` 后仍只见倒计时，右键 ppt time 浮窗，选择 **WPS 演示页码联动 → 启用页码加载项**，保存文稿并完全退出 WPS。重新打开 WPS 和文稿后按 `F5`，加载项会从 WPS 内部读取当前幻灯片和总页数。可在同一菜单的 **查看联动状态** 中确认是否已连接。此方式需要 WPS 客户端支持演示 JS 加载项。放映结束后仍可右键选择 **复制页码诊断**，查看最近一次外部读取结果。

也可以选择 **打开放映文件…**，由 ppt time 启动所选文稿并跟踪其放映窗口。读数完整时显示如 `3/20`；只取得文稿中的实际页码时显示“第 3 页”。自定义放映中，若 PowerPoint 仅返回出场顺序，会标注“放映第 3 页”。

| 场景 | 演讲者可用的功能 |
| --- | --- |
| Windows PowerPoint / WPS 演示 | 全屏检测、开始和结束联动、远程翻页。自动读取当前幻灯片在文稿中的实际编号和总页数。WPS 页码取决于安装版本开放的演示接口。 |
| 内置 PDF 放映 | 通过 **打开放映文件…** 打开 PDF，程序自行渲染，支持前后翻页、真实页码和退出。 |
| Windows 外部 PDF 阅读器 | 可检测 WPS PDF 和 Adobe Acrobat / Reader 放映。WPS PDF 可选装随程序提供的加载项来报告页码。能否取得外部阅读器的页码取决于其版本和界面模式。 |
| 内置视频播放 | 从控制台或文件菜单打开 MP4、MOV、WebM 等本地视频。`空格` 切换播放与暂停，前后翻页键分别跳转 10 秒，可跳转范围取决于媒体文件。 |
| Mac | 提供源码启动器及构建流程。PowerPoint 联动通过 Apple 自动化接口读取放映状态。首次访问按系统提示授权。Mac 版 WPS 和 Keynote 可使用 **手动放映计时**。Mac 的真实 Office 放映与多显示器行为仍待实机验证。 |

托盘菜单第一行显示现场状态。出现 **放映中 · 未读取到页码** 时，倒计时仍继续，观众端会收起页码。仅页码模式等待取得页码后再显示计时条。外部 PowerPoint 和 WPS 的实际页码同步仍需在会场安装版本上验证，可使用同一份文稿试一次顺序翻页、跳页和退出放映。

**在 WPS 内放映 PDF** 时，可从 Windows 托盘进入 **WPS PDF 联动 → 安装加载项**，完全退出并重新打开 WPS，再打开 PDF。选择 **查看联动状态** 可检查连接状态及放映中的页码。停用时选择 **移除本软件的加载项** 并重启 WPS。通过 ppt time 自带的 PDF 播放器放映可以直接读取页码。

## 计时、放映条与结束提示

**调整放映条大小**：右键初始浮窗或托盘图标，进入 **放映条外观… → 大小与位置 → 放映条大小**。自动模式按初始浮窗的时间字号和投影屏尺寸计算比例，时间、页码和图标一起缩放。在默认设置下，1080p 屏幕的放映时间字号从 20 px 提升到 32 px。也可选固定比例，其中 100% 对应 1.0.2 的尺寸。旧版已单独调过放映字号的配置会保留原大小，可在此切换为自动。确认后，右键选择 **预览放映计时条（8 秒）** 即可查看。

- **计时器设置**可调整环节名称、时长、预警颜色及声音、自动切换环节、显示屏和整体缩放。Windows 放映时，右键计时条可打开菜单，左键单击仍可翻页。默认在放映期间挂起全局快捷键，遥控翻页笔可继续向演示程序送键。
- **节点预警音效**和**环节结束音效**可各自试听并选择六种原创合成提示音，也可选择本机的 MP3 或 WAV 文件。
- **放映条外观…**可设置位置、偏移、文字及背景外观。放映可选经典刘海、优雅刘海和时间胶囊。画面输出可跟随放映屏或选择所有屏幕。Windows 的智能闪避从放映窗口读取被覆盖区域，画面兼容时可自动淡出。
- **逐页计时与结束提示…**可设置整份文稿或单页倒计时。选择文件并将页码设为 `*` 可设置该文稿起始时长。指定页码后，每次进入该页会使用对应时长。未设置的页面延续当前计时。可给同页不同文稿分别设置规则。
- 可选择仅显示页码、超时后继续正计时、自定义结束文字，或在指定屏幕和全部屏幕显示结束黑屏。`Esc`、空格或回车关闭黑屏。勾选允许关闭时，文字提示也支持这些按键；操作员还可用托盘菜单 **关闭结束提示**。开启自动切换环节时，时间到后先进入下一环节。

## 会场控制台

右键托盘后打开 **会议控制台…**。

| 标签页 | 操作 |
| --- | --- |
| 大会主画面 | 指定背景图片和输出屏幕，可叠加日期时间。在预览确认后投放。按需开启“放映结束后自动显示主画面”和“收到末页结束信号时提前切回主画面”。 |
| 串场图片 | 添加候场、茶歇和颁奖等图片，调整顺序，选择一张投放，使用上一张和下一张切换。 |
| 会议议程 | 设置讲者、时长和本地附件，选定议程后开始。点击 **日程管理与导入…** 可维护多场会议和会场，登记日期、开始时间、讲者单位与场次状态，筛选后查看表格、卡片或时间线，以及状态和文件关联数量。支持多选批量修改、删除，也可导入 CSV 或 XLSX 日程表。 |
| 文件看板 | 登记会场文件，找到后直接打开放映。接收的会议文件也可加入看板。 |

## 局域网遥控与收件

1. 讲台电脑和操作手机或台下电脑接入同一网络。打开 **局域网遥控与收件…**，填写会场名称，选择可由其他设备访问的本机网卡地址，并指定收件位置。默认端口为 `48731`。
2. 开启服务，为对应角色复制连接或出示二维码。**投稿**可提交会议文件或整份文件夹，目录层级会保留。**监看**可查看会场状态、计时和日程。**控制**可打开会场文件、翻页、停止、调整倒计时及切换大会画面。
3. 若需向监看端和控制端提供屏幕预览，勾选“向监看端和控制端共享所选屏幕画面”，再选择演示用的屏幕。控制端可以逐级浏览目录，使用列表或图标视图，按名称、数字顺序、修改日期升降序排列。勾选“整理文件”可从收件箱移除单个文件，在“已移除”中点相应条目即可还原。原位置存在同名文件时，恢复的文件会附加序号。
4. 多会场同步时，在局域网工作台填写目标会场提供的对应角色连接，可发送文件或拉取、推送议程。议程附件在本机按已收文件匹配。每个会场在自己的讲台电脑上运行服务。

从工作台选择 **打开日程看板**，可按会议、会场、日期、上下午、讲者、单位、文稿和状态组合筛选，查看列表、卡片或时间线和对应的场次数量。文稿收件状态按议程文件名与本场文件看板匹配。监看链接只读，控制链接可新增、编辑、调整议程顺序或删除条目。其他操作员同时修改时，页面会提示重新核对最新议程。控制端还可切换放映时的仅页码模式和超时正计时。

在控制页调整计时条的字体、颜色、字号或位置，讲台的放映条会同步更新。

如需在 OBS 中叠加倒计时，从讲台复制 **OBS 透明计时条** 链接，添加为 OBS 的浏览器源。另一条 **OBS 投影画面** 链接显示讲台选中的屏幕，需要先开启屏幕共享。投影画面约每秒刷新一次，适合监看静态幻灯片。

Windows 托盘中的 **讲台状态与桌面快捷操作…** 显示 CPU、内存、图形设备和供电情况。勾选 **本次运行允许遥控端操作 Windows 桌面** 后，控制链接持有者还可显示桌面、打开资源管理器或任务视图、切换或关闭前台窗口、打开系统投影设置、锁屏。重新启动 ppt time 后需重新勾选。

## 跨网手机投稿

**云端收件与手机扫码…** 需要会务方提供可访问的 ppt time 中转服务。桌面包提供 `relay_server.py`，会务方负责运行服务和配置域名。该服务仅需要 Python 3.11 标准库。

在服务器目录运行下列命令。第一次会创建收件目录和数据库。带 `--show-receiver-key` 的命令用于在服务端显示讲台接收密钥。

```sh
python3 relay_server.py --directory ./meeting-inbox --show-receiver-key
python3 relay_server.py --directory ./meeting-inbox --bind 127.0.0.1 --port 8766
```

此方式可在同机用 `http://127.0.0.1:8766` 调试。要给异地手机和讲台使用，可将 HTTPS 反向代理转发到本机端口，或给服务直接配置与会场域名匹配的证书和私钥。直接启用 TLS 的示例为

```sh
python3 relay_server.py --directory ./meeting-inbox --bind 0.0.0.0 --port 8766 \
  --cert /path/to/fullchain.pem --key /path/to/privkey.pem
```

在讲台电脑 **云端收件与手机扫码…** 中填入完整服务地址，例如 `https://收件域名:8766`、接收密钥和本地收件目录，点 **连接并收件**。给投稿者使用手机投递链接或二维码。支持 PPT、PPTX、PPS、PPSX、DPS 和 PDF，单文件上限 256 MiB。讲台验证大小和 SHA-256、保存文件并向服务确认后，中转服务删除对应文件。

如需让讲者将文稿交给指定场次，勾选 **在手机投递页显示会议议程，允许讲者选择对应场次**。打开投递链接的讲者可按日期、时间、会场和议题选定场次，讲台收件后会将文件关联到对应议程。关闭该选项后，讲者仍可直接投稿。

## 从源码运行

使用 Python 3.11。Windows PowerShell 示例

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py
```

macOS 14 或更高版本可解压源码后双击 `Start-ppt-time.command`。也可在终端运行 `bash Start-ppt-time.command`。首次运行需要联网安装 Qt、PDF、媒体和二维码组件。启动器在用户的 `~/Library/Application Support/ppt time` 下保存隔离环境。CI 提供对应 Apple 芯片和 Intel 的 `ppt time.app` 构建流程。

## 测试与生成发布包

在已安装运行依赖的 Python 3.11 环境中运行

```sh
python -m unittest discover -s tests -v
```

在 Windows x64 PowerShell 中激活此环境后，以下命令将生成目录版程序与安装器。构建安装器使用 Windows x64 CPython 3.11 的 `msilib` 和系统 .NET Framework 4.x 的 C# 编译器。

```powershell
python -m pip install pyinstaller==6.22.3
$buildPython = (Get-Command python).Source
& $buildPython powerpoint_addin/build.py
& $buildPython -c "from comtypes.client import GetModule; GetModule('UIAutomationCore.dll'); import comtypes.gen.UIAutomationClient"
$env:PATH = "$env:SystemRoot\System32;$env:SystemRoot;$env:SystemRoot\System32\Wbem"
& $buildPython -m PyInstaller --onedir --windowed --noconfirm --name 'ppt time' `
  --icon icon.ico --version-file version.txt --add-data 'icon.ico:.' --add-data 'icon.svg:.' `
  --add-data 'wps_pdf_addin:wps_pdf_addin' --add-data 'wps_presentation_addin:wps_presentation_addin' `
  --hidden-import comtypes.gen.UIAutomationClient `
  --add-data 'powerpoint_addin/PptTime.PowerPointBridge.dll:powerpoint_addin' `
  --hidden-import PySide6.QtPdf --hidden-import PySide6.QtMultimedia `
  --hidden-import PySide6.QtMultimediaWidgets --hidden-import win32timezone `
  --hidden-import win32com.client --hidden-import win32ui --collect-submodules segno main.py
Copy-Item -Path README.md,LICENSE,relay_server.py -Destination 'dist/ppt time'
& $buildPython packaging/build_windows_installer.py 'dist/ppt time' dist/ppt-time-Setup.exe --version 1.0.8 --icon icon.ico
Compress-Archive -Path 'dist/ppt time/*' -DestinationPath dist/ppt-time-Windows-x64.zip
```

项目中的 `.github/workflows/build.yml` 可手动运行，或使用与程序版本一致的 Git 标签（例如 `v1.0.0`）触发发布。它分别生成 Windows 安装版和便携版、两种 Mac 架构的应用包、对应源码包及 SHA-256 校验文件。桌面分发包内附源码、许可证及服务脚本。

## 来源

项目基于 QwejayHuang 的开源 [QTimer](https://github.com/Qwejay/QTimer) v1.6.0（提交 `6c966c1`）。本版在该基础上实现会议计时、放映控制及文件传输，使用 GPL-3.0 许可证，见 [LICENSE](LICENSE)。
