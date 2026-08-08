# 项目记忆（MEMORY.md）

> 供后续开发/维护快速恢复上下文。最后更新：2026-08-08

## 一、项目概况

- 名称：自动 M/PBO 报告工具（PyQt5 + openpyxl 桌面工具，macOS 上开发，Windows 发布）
- 入口：`main.py`；依赖：`requirements.txt`（PyQt5、openpyxl、Pillow 等）
- Windows 打包：GitHub Actions（`.github/workflows/build.yml`），push 到 main 自动构建，产物 `dist/auto_report.zip`（PyInstaller 目录包 + 图标/splash）
- 远程仓库：`origin git@github.com:Longtianhong88888/auto_report_ABU.git`，分支 `main`
- 最新提交：`ac3986f 解除图片像素上限：关闭Pillow DecompressionBomb限制，支持超高分辨率图片`
- 近三次提交：`ac3986f`（解除图片像素上限）、`5577f1f`（关闭MKLDNN修复Windows CPU崩溃）、`ab4504f`（修复CI模型预下载）
- 本地未提交内容：`manual_shots/`（说明书截图）、使用说明书 PPT、`问题点.txt`（用户问题清单）

## 二、模块结构

| 文件 | 职责 |
|---|---|
| `main.py` | 主窗口、文件加载、映射管理、输出报告、程序入口、登录/授权对话框（UserIdDialog、AuthorizeIdDialog） |
| `dialogs.py` | 对话框：ImageSetupDialog、VersionFinderDialog、ArchiveConfigDialog、JMPConfigDialog、SourceSelectDialog、InternalImageSelectDialog、BatchImageDialog、TransformDialog |
| `mapping_operations.py` | `MappingOperations` 混入类：数据/图片/归档/JMP 映射执行 |
| `safe_eval.py` | 自定义表达式受限安全求值 |
| `version_finder.py` | 从 PDF/Excel 档案提取版本号 |
| `user_auth.py` | 授权工号加载/保存/批量解析 |
| `table_zoom.py` | `TableZoomMixin`：表格缩放（Ctrl+滚轮/Mac 捏合）与 Ctrl+Shift+方向键快速选区 |
| `utils.py` | 公共函数：`column_width_chars`（区间列宽）、`apply_uniform_sizes`（统一尺寸）、`normalize_mappings`（配置标准化） |
| `ocr_engine.py` | OCR 引擎：PaddleOCR 懒加载单例（兼容 ocr()/predict()）、预处理、数值提取 |
| `ocr_worker.py` | OCR 后台线程（对话框预览用） |
| `ocr_dialog.py` | OCR 切片量测配置对话框（ROI 框选、预览、表达式） |
| `ui_theme.py` | Apple 风格 UI 设计体系（QSS/颜色/间距/窗口尺寸，移植自 MC_LogAnalysis） |
| `user_guide.py` | 应用内「使用说明」对话框（HTML 内容） |
| `constants.py` | 尺寸常量、默认授权工号、管理员与密码 |

## 三、核心业务流程

### 启动与授权
1. splash 欢迎画面 → 主界面先显示（作为背景）→ `UserIdDialog` 工号验证（模态）
2. 授权工号（默认 `G1655895`，大小写不敏感）→ 直接进入主界面；未授权 → 红字“您当前ID未授权！”
3. 管理员 `G1659304` → 授权密码输入框（打码）→ 密码 `Zy1659304`（在 constants.py）→ `AuthorizeIdDialog` 批量授权新工号（空格/逗号/分号/斜杠/顿号分隔）
4. 授权结果持久化：开发环境写程序目录 `authorized_ids.json`；打包后（`sys.frozen`）写 `~/.auto_report_authorized_ids.json`；读取失败回退默认列表

### 输出报告执行顺序（固定）
`归档映射 → 数据/图片映射 → 单独单元格更新（cell_edits 重放）→ JMP（最后）`

- 单独单元格更新来源：右键“修改单元格内容”、版本号查找写入；重放时若编辑位于归档区块内，列号随区块右移（多个归档区块链式右移）
- 输出时自动处理 Summary：
  - Date 表头右侧填当天日期（`datetime.now()`，保留单元格日期格式）
  - Build Phase / Configuration / Event：从“IPQC 数据所在文件夹名 → 上一级文件夹名”依次解析并合并（每项取第一个命中）
  - 解析失败字段跳过更新，弹窗提示“请在报告中手动修改”，弹窗会列出已查找的文件夹
- 默认输出文件名：沿用模板文件名格式，替换 Configuration（`C\d{4}`）与 Event（`MBO|PBO`）段，不含 Build Phase；保存对话框预填，其余由用户自行修改
- 所有报表预览表格支持缩放：Ctrl+滚轮（Windows）与 Mac 触控板双指捏合（`TableZoomMixin`，table_zoom.py），同步缩放列宽/行高/字体，事件 50ms 合并防卡顿；应用于主预览表与 SourceSelectDialog
- 表格支持 Ctrl+Shift+方向键按 Excel 习惯快速扩展选区（连续非空/空段边界）；活动格用 `selectionModel().setCurrentIndex(..., NoUpdate)` 设置，避免塌缩选区

### 配置体系（2026-08 重构后）
- 打开模板**不再自动加载**配置；`self.mappings` 初始为空
- “保存配置”：`QFileDialog` 自定义文件名（默认“模板名_config.json”），内容只保留映射路径，剥离 `image_bytes`
- “导入我的配置”：选择 .json；来源模板与当前不一致时确认；导入后刷新映射列表
- 图片映射记录来源路径 `image_src_sheet` + `image_src_pos`（源 Sheet + 图片锚点位置），输出时从当前 IPQC 数据源按位置重新读取；旧配置 `image_ref=[sheet, idx]` 兼容回退
- 配置导入时经 `utils.normalize_mappings` 标准化：JMP `header_cols` 统一为字符串列表（旧配置字符串兼容移到导入时处理）、`block_range`/`target_range` 列表转元组
- 图片缓存失败会统计并弹窗提示（全部失败提示文件可能损坏），不再静默忽略
- OCR 切片量测（右键 → OCR识别）：单值/按标签/全部数值/自定义模式，ROI 框选、5 种预处理；映射写入 target_cells；数值提取优先取“= 后面的值”避免标签数字误判；`all_numbers` 多值时逗号拼接写入单格

## 四、业务提取规则

- Process control rev.：匹配 `Table : Process Control Rev x.xx`（取最大版本，如 BUF ERS → 3.55）；匹配不到 → 跳过 + 提示手动修改
- ERS / VSR rev.：文件名 `RevN` 优先，正文 RevN 兜底（取最大）
- MCO rev.：文件名末尾 `-NN`（如 `056-25088-02.pdf` → 2），再兜底 RevN
- Build Phase：`C\d+\.\d+`（如 C6.0）
- Configuration：固定 5 位 = `C` + 4 位数字（如 C6004/C6081），内部规则
- Event：`MBO|PBO`（大小写不敏感，输出大写）
- 版本号对话框自动推荐：模板名以 CLO 开头 → 推荐 099-55402 系 ERS；其他专案 → 优先 “BUF Module ERS”

## 五、渲染“所见即所得”（易踩坑）

- 无填充单元格不渲染黑色；表格强制白色底色（防系统深色模式）
- **预览不允许出现黑色单元格**：深/黑色填充（主题 dk1=#000000、indexed 8 等，亮度 < 0.5）在预览中一律替换为浅灰 `#D9D9D9`（`_cell_fill_color`，阈值与替换色为类常量 `PREVIEW_DARK_FILL_THRESHOLD` / `PREVIEW_DARK_FILL_REPLACEMENT`），自动字色随浅底变黑字，保证可读；仅预览层替换，输出文件不受影响
- “自动色”字体（theme 0/1、indexed 8/9/64/65、未设置颜色）按背景亮度自适应：亮度 < 0.5 → 白字，否则黑字；显式 rgb 原样
- 主题色 tint 按 OOXML 公式计算（负值向黑、正值向白）
- **openpyxl 不展开 `<col min max>` 区间列**（如 `<col min="3" max="13" width="33"/>`）→ `_column_width_chars()` 遍历 `column_dimensions` 按 min/max 匹配；未定义列回落 `defaultColWidth`；行高同理回落 `defaultRowHeight`（main.py 与 dialogs.py 各有一份）
- 正则 `\b` 会把下划线当单词字符 → 文件名/文件夹名匹配用 `(?<![A-Za-z0-9])…(?![A-Za-z0-9])`
- **混入类 MRO 坑**：PyQt 类的 MRO 中 Qt 类排在普通混入类之前，`TableZoomMixin.eventFilter` 里不能调 `super().eventFilter`（会落到 `object` 报 AttributeError，且异常被 Qt 吞掉后事件空转像卡死）→ 直接 `return False` 放行即可

## 六、已知问题与待办

- 【已修复 2026-08-08】Mac 双指捏合缩放报错：`'QPinchGesture' object has no attribute 'scaleDelta'`。`QPinchGesture` 没有 `scaleDelta()`，`table_zoom.py` 改为 `scaleFactor() / lastScaleFactor()` 计算每次手势事件的相对缩放倍数，并防御 `lastScaleFactor=0`
- 【已修复 2026-08-08】模板预览/图片输出提示“图片像素超过限制”载入失败：根因是 Pillow DecompressionBomb 默认上限（178,956,970 像素）；已在 `mapping_operations._process_image_data` 与 `ocr_engine.ocr_image` 打开图片前置 `Image.MAX_IMAGE_PIXELS = None` 解除限制（本地受信任的高分辨率显微镜图片），并移除从未启用的 `OCR_MAX_IMAGE_DIMENSION_FOR_PREVIEW` 残留常量
- 【已修复 2026-08-07】JMP 预览上限 500 行的问题：`display_sheet` / `SourceSelectDialog` 行上限改为 `MAX_PREVIEW_ROWS`（constants.py，默认 10000，按内容实际范围计算）；渲染循环改为稀疏遍历 `ws._cells`（不再物化空单元格，8000×30 渲染约 2.9s）
- ACF 模板约 109MB，加载慢（非代码问题）；测试优先用 BUF27 模板（35MB）
- 归档读取公式缓存值；文件未经 Excel 保存过时公式列为空 → 输出告警
- openpyxl 保存后图片 `ref` 会被关闭 → 用 `_snapshot_template_images` / `_refresh_image_refs` 重建
- `generate_acf_report.py` 已从磁盘删除（README 已更新，代码无引用）
- 本地 `authorized_ids.json` 当前含 G1653332/G1655895/G1659304（不入库，gitignore 已忽略）
- OCR 功能：本机 venv 已装好 paddlepaddle 3.3.1 / paddleocr 3.7.0 / paddlex 3.7.2（`pip install -r requirements.txt`），OCR 引擎实测可用（内置 PP-OCRv6 模型，5s 识别 155.76/2.30 全对）；requirements 锁 `paddlepaddle/paddleocr>=3.0,<4.0`
- Windows 打包为“解压即用”：build.yml 用 `--collect-all paddle/paddleocr/paddlex/cv2` 收集全部依赖 + `--add-data paddleocr_models` 内置模型，zip 解压后无需安装任何东西
- OCR 已知限制：按标签独立 ROI（label_rois / ocr_batch_with_rois）尚未在对话框/输出中接入；输出阶段 OCR 同步执行（已加 processEvents 刷新进度，未做线程化）

## 七、测试与交付

- 离屏测试：`QT_QPA_PLATFORM=offscreen`；对话框/主窗口可直接 `grab()` 出真实截图；模态 QDialog 用 stub（monkeypatch `exec_`/`get_results`）
- 测试用小模板（如 `/tmp/small_tpl.xlsx`）避免大文件卡顿
- 说明书：`使用说明书_自动MPBO报告工具.pptx`（15 页，截图 + 文字；原始截图在 `manual_shots/`）
- Windows 版：GitHub Actions 构建，Actions 页面下载 `auto_report_package`

## 八、界面（Apple 风格，2026-08-08 移植自 MC_LogAnalysis）

- QSS 卡片式：背景 `#F5F5F7`、卡片白、主色 `#007AFF`（`ui_theme.py` 的 `APPLE_QSS`）；`main.py` 里 `app.setStyleSheet(APPLE_QSS)` 全局生效（主窗口 + 所有对话框统一），主窗口再设一份同款样式
- 窗口尺寸改为按屏幕可用区域 **80%** 动态计算（`window_target_size`，最小 900×620），替换原固定 `setGeometry(1400, 850)`
- 三个面板（报告Sheet列表 / 预览表 / 映射列表）均为白色圆角卡片（`setProperty("card", True)` + `CARD_PAD` 内边距）；顶部按钮行右侧新增「使用说明」链接按钮（`user_guide.py` 对话框，`show_user_guide`）
- 主按钮「打开报告文件」为实心蓝（`primary`）、「打开IPQC数据源文件」为浅灰次按钮（`secondary`）
- 启动画面：与主窗口同尺寸（splash.png **cover 裁切**填满），最短展示 1000ms，`FADE_MS=300` 淡出衔接主窗口；登录对话框在淡出结束后弹出（`_crossfade(on_finished=start_login)`），`_ANIMS` 持有动画引用防提前回收
- `AA_EnableHighDpiScaling` / `AA_UseHighDpiPixmaps` 在 QApplication 创建前设置（原有，保留）；版权仍在状态栏右下角（灰色 11px，`Copyright © 2026 ABU NPD EOL`）
- 注意：QSS 的 `QWidget{font-size:13px}` 与表格缩放功能不冲突（缩放用显式 item 字体）；离屏测试时透明度动画提示 "plugin does not support setting window opacity" 属环境限制，真机正常
