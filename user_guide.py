"""应用内「使用说明」对话框。"""
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QTextBrowser, QDialogButtonBox,
)

USER_GUIDE_HTML = """
<h2>自动 M/PBO 报告制作软件 · 使用说明</h2>

<h3>一、基本流程</h3>
<ol>
  <li><b>打开报告文件</b>：选择报告模板（模板内嵌图片会自动快照，输出时原样保留）。</li>
  <li><b>打开 IPQC 数据源文件</b>：程序读取数据源中的数据、图片与归档内容供映射使用。</li>
  <li>勾选需要输出的 Sheet，在预览表中选中区域，按右键添加映射：
      <b>数据填充 / 图片 / 归档 / JMP / OCR 切片量测</b>。</li>
  <li>点击<b>确认映射</b>，再点击<b>输出报告</b>。
      执行顺序固定：归档映射 → 数据/图片映射 → 单独单元格更新 → JMP。</li>
</ol>

<h3>二、常用功能</h3>
<ul>
  <li><b>版本号查找</b>：从档案自动识别 Process control / ERS / VSR / MCO 版本号并写入 Summary。</li>
  <li><b>修改单元格内容</b>：右键单元格可直接修改/新增值，输出时自动重放（归档区块内会自动右移列）。</li>
  <li><b>OCR 切片量测</b>：右键 OCR 识别，支持 ROI 框选、5 种预处理、
      单值/多值标注/全部数值/自定义表达式模式，结果写入目标单元格。</li>
  <li><b>保存配置 / 导入我的配置</b>：映射可跨会话复用（配置只存映射路径，不含图片字节）。</li>
</ul>

<h3>三、操作提示</h3>
<ul>
  <li>预览表格支持 <b>Ctrl+滚轮</b> 缩放（Mac 触控板双指捏合）；
      <b>Ctrl+Shift+方向键</b> 按 Excel 习惯快速扩展选区。</li>
  <li>首次使用需输入<b>已授权工号</b>；管理员可批量授权新工号。</li>
  <li>输出文件名默认沿用模板命名格式（替换 Configuration / Event 段），保存对话框可自行修改。</li>
  <li>Summary 自动填充当天日期与 Build Phase / Configuration / Event（从 IPQC 数据源文件夹名解析）。</li>
</ul>
"""


def show_user_guide(parent=None):
    """显示使用说明对话框。"""
    dlg = QDialog(parent)
    dlg.setWindowTitle("使用说明")
    dlg.resize(720, 560)
    layout = QVBoxLayout(dlg)
    browser = QTextBrowser()
    browser.setOpenExternalLinks(False)
    browser.setHtml(USER_GUIDE_HTML)
    layout.addWidget(browser, 1)
    buttons = QDialogButtonBox(QDialogButtonBox.Ok)
    buttons.accepted.connect(dlg.accept)
    layout.addWidget(buttons)
    dlg.exec_()
