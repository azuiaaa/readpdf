from PyQt5 import QtCore
from PyQt5.QtWidgets import *
from testcase_homekit_iteration_tool import main
from messageDialog import UI_Message
from pathlib import Path
import os


class UIDialog(QDialog):
    _flag_signal = QtCore.pyqtSignal(dict, bool)

    def __init__(self):
        super(UIDialog, self).__init__()
        self.current_data = dict()
        self.exit_flag = False

    def setupUi(self):
        """
        生成主窗口的主要布局
        :param childWindow: 子窗口对象
        :return:
        """
        # 引入样式文件
        # self.stylefile = "style.qss"
        # self.qssStyple = CommonHelper.readQSS(self.stylefile)
        self.qssStyle = "QPushButton{border-radius: 16px;background-color: #ffffff;min-width: 67px;min-height: 24px;}" \
                        "QPushButton#Cancel, QPushButton#selectExcel, QPushButton#selectPdf {border: 1px solid #b3b2b3;border-radius: 16px}" \
                        "QPushButton#Play {background-color: rgb(240, 230, 140);border: 1px solid rgb(240, 230, 140);}" \
                        "QPushButton#OK, QPushButton#selectExcel, QPushButton#selectPdf {background-color: rgb(0, 136, 249);border: 1px solid rgb(0, 136, 249);color: rgb(255, 255, 255);border-radius: 16px;}" \
                        "QPushButton#OK:pressed, QPushButton#Cancel:pressed, QPushButton#selectExcel:pressed, QPushButton#selectPdf:pressed {background-color: rgb(240, 239, 240);border-radius: 16px;}" \
                        "QTextBrowser {background-color: transparent;}QTextEditor#focus {border: 3px solid #FA8072;}"
        self.setStyleSheet(self.qssStyle)

        self.setWindowTitle("手动操作提示")  # 窗口标题
        self.setGeometry(400, 400, 400, 300)  # 窗口位置与大小

        # 信息提示框
        # self.textBrowser = QTextBrowser()
        # self.textBrowser.setText("")

        # 信息提示框
        self.label = QLabel()
        self.label.setText("<html><head/><body>用例位置<br>excel</body></html>")
        # self.label.setAlignment(Qt.AlignCenter)

        self.label_2 = QLabel()
        self.label_2.setText("<html><head/><body>苹果用例<br>PDF</body></html>")

        self.label_3 = QLabel()
        self.label_3.setText("<html><head/><body>更新用例版本<br>如R11.2</body></html>")

        # fail原因文本框
        self.editor_excel = QLineEdit()
        self.editor_pdf = QLineEdit()
        self.editor_sheet = QLineEdit()

        # 打开文件资源管理器按钮
        self.btn_excel = QPushButton()
        self.btn_excel.setObjectName("selectExcel")
        self.btn_excel.setText("选择")
        self.add_shadow(self.btn_excel)
        self.btn_excel.clicked.connect(lambda: self.handle_select(True))

        # 打开文件资源管理器按钮
        self.btn_pdf = QPushButton()
        self.btn_pdf.setObjectName("selectPdf")
        self.btn_pdf.setText("选择")
        self.add_shadow(self.btn_pdf)
        self.btn_pdf.clicked.connect(lambda: self.handle_select(False))

        # 确认按钮
        self.btn_ok = QPushButton()
        self.btn_ok.setObjectName("OK")
        self.btn_ok.setText("更新")
        self.add_shadow(self.btn_ok)
        self.btn_ok.clicked.connect(self.handle_ok)

        # 关闭按钮
        self.btn_cancel = QPushButton()
        self.btn_cancel.setObjectName("Cancel")
        self.btn_cancel.setText("关闭")
        self.add_shadow(self.btn_cancel)
        self.btn_cancel.clicked.connect(self.handle_cancel)

        self.glayout = QGridLayout()
        self.glayout.setObjectName("main")

        label_y, text_y, btn_y, btn_w, btn_h = 0, 1, 4, 3, 1
        text_h = btn_h

        self.glayout.addWidget(self.label, 0, label_y)
        self.glayout.addWidget(self.editor_excel, 0, text_y, btn_h, btn_w)  # row, col, rowspan, colspan
        self.glayout.addWidget(self.btn_excel, 0, btn_y)
        self.glayout.addWidget(self.label_2, 1, label_y)
        self.glayout.addWidget(self.editor_pdf, 1, text_y, btn_h, btn_w)
        self.glayout.addWidget(self.btn_pdf, 1, btn_y)
        self.glayout.addWidget(self.label_3, 2, label_y)
        self.glayout.addWidget(self.editor_sheet, 2, text_y, text_h, 4)
        self.glayout.addWidget(self.btn_ok, 3, 3)
        self.glayout.addWidget(self.btn_cancel, 3, 4)
        self.setLayout(self.glayout)

    def handle_ok(self):
        """

        :return:
        """
        try:
            path_excel = self.editor_excel.text()
            path_pdf = self.editor_pdf.text()
            new_sheet_name = self.editor_sheet.text()

            if path_excel == "":
                UI_Message.MessageWarning(self, text="用例库excel路径不可为空！")
                self.editor_excel.setFocus()
                return
            if path_pdf == "":
                UI_Message.MessageWarning(self, text="苹果用例PDF路径不可为空！")
                self.editor_pdf.setFocus()
                return
            if new_sheet_name == "":
                UI_Message.MessageWarning(self, text="本次用例更新版本不可为空！")
                self.editor_sheet.setFocus()
                return
            if not Path(path_excel).exists() or not Path(path_pdf).exists():
                UI_Message.MessageWarning(self, text="文件不存在，请检查路径及文件名称是否正确！")
                return
            if os.path.splitext(path_excel)[-1] not in [".xlsx", ".xls"]:
                UI_Message.MessageWarning(self, text="用例库文件类型错误，请检查是否为.xlsx或者.xls！")
                return
            if os.path.splitext(path_pdf)[-1] != ".pdf":
                UI_Message.MessageWarning(self, text="苹果用例文件类型错误，请检查是否为pdf！")
                return
            path_result = main(path_excel, path_pdf, new_sheet_name)
            if path_result:
                UI_Message.MessageInformation(self, text="用例更新完成\n文件路径：{}".format(path_result))
        except FileNotFoundError:
            UI_Message.MessageWarning(self, text="未找到对应文件，请检查路径及文件名称是否正确！")
            return
        except TypeError as e:
            print(e)

    def handle_cancel(self):
        """

        :return:
        """
        self.close()

    def handle_select(self, flag=True):
        """
        选择文件
        :param flag: true表示为链接excel文件按钮，false表示链接pdf按钮
        :return:
        """
        if flag:
            self.editor_excel.setFocus()
            self.editor_excel.setText(QFileDialog.getOpenFileName()[0])
        else:
            self.editor_pdf.setFocus()
            self.editor_pdf.setText(QFileDialog.getOpenFileName()[0])

    def add_shadow(self, button):
        """
        设置按钮阴影，暂未找到方法直接在qss文件中定义
        :param button: 需设置的按钮对象
        :return:
        """
        self.effect_shadow = QGraphicsDropShadowEffect(self)
        self.effect_shadow.setOffset(0, 0)  # 偏移
        self.effect_shadow.setBlurRadius(10)  # 阴影半径
        self.effect_shadow.setColor(QtCore.Qt.gray)  # 阴影颜色
        button.setGraphicsEffect(self.effect_shadow)  # 将设置套用到button窗口中
