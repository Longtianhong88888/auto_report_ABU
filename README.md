# 自动 M/PBO 报告工具

基于 PyQt5 + openpyxl 的桌面工具，用于从 Excel 报告模板生成 M/PBO 报告。

## 功能

- 打开 Excel 报告模板（.xlsx），在表格中预览并勾选需要输出的 Sheet；
- 打开 IPQC 数据源文件，读取数据与内嵌图片；
- 通过右键菜单为模板区域添加四类映射：
  - **数据填充**：把数据源区域复制到模板区域，支持数据转换（除以 1000、去末尾字母、安全自定义表达式）；
  - **图片**：单张或批量，可从数据源内嵌图片或外部文件选择，支持旋转、缩放；
  - **归档（右移）**：把锚定列起的数据块整体右移 1 列，新数据取自同一 Sheet 的"来源列"（默认 J 列），归档后自动补齐框线（细框线 + 粗外框）；
  - **JMP 数据区**：按表头列数写入 JMP 数据。
- 输出顺序固定：归档 → 数据/图片 → 单独单元格更新 → JMP；
- Summary 自动处理：
  - Date 自动填充报告当天日期；
  - Build Phase / Configuration / Event 从文件夹名提取（从 IPQC 数据所在文件夹向上两级查找）；
  - 版本号（Process control / ERS / VSR / MCO）从指定 PDF/Excel 档案自动提取；
- 输出文件名沿用模板格式并更新 Configuration / Event 段；
- 映射配置只记录映射路径，可自定义名称保存/导入复用（`保存配置` / `导入我的配置`）。

## 用户验证与授权

- 启动后需输入已授权工号才能使用；
- 管理员工号（默认 `G1659304`）输入后需验证授权密码（默认 `Zy1659304`），密码正确后可批量授权新工号（空格、逗号、分号、斜杠等分隔）；
- 授权结果保存在程序目录（打包后保存在用户目录）的 `authorized_ids.json`。

## 使用步骤

1. 输入授权工号进入主界面；
2. `打开报告文件` 选择模板；
3. `打开IPQC数据源文件` 选择数据源；
4. 在模板表格中选中区域，右键选择映射类型并配置；
5. 勾选需要输出的 Sheet，点击 `输出报告` 保存。

## 环境与打包

```bash
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

Windows 可执行文件由 `.github/workflows/build.yml` 在 GitHub Actions 上自动构建（目录包 + zip 工件）。

## 模块结构

```text
main.py                 主窗口、文件加载、映射管理、报告输出、程序入口
dialogs.py              所有对话框（图片设置/转换/归档/JMP/源选择/图片选择/批量/版本号）
mapping_operations.py   数据/图片/归档/JMP 映射执行逻辑（MappingOperations 混入类）
safe_eval.py            自定义表达式的受限安全求值
version_finder.py       版本号从 PDF/Excel 档案中提取
user_auth.py            用户ID授权加载/保存与批量解析
constants.py            统一尺寸常量、授权ID与管理员配置
```

## 注意事项

- 归档读取的是文件中的**公式缓存值**。若模板/数据源未经 Excel 打开保存过（例如刚由程序生成），公式列可能读为空值，输出时工具会给出警告。
- 数据源文件若与模板存在同名 Sheet，映射按各自配置的 `source_sheet` 读取，不再按同名自动匹配。
- 自定义表达式在受限的安全子集内求值（仅 `x` 与四则运算、比较、`abs/round/min/max/int/float/str/len/sum`），不支持任意 Python 代码。
- 图片选择对话框支持按位置/Sheet 搜索、缩略图浏览与大图预览，方便在数百张图片中定位。
