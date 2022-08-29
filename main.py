from PyQt5.QtWidgets import QApplication, QMainWindow
import sys

from uiDialog import UIDialog

if __name__ == '__main__':
    # 创建QApplication类的实例
    app = QApplication(sys.argv)
    # 创建一个主窗口
    mainWindow = QMainWindow()  # 使用重写过的MainWindow类
    # 创建Ui_MainWindow的实例
    ui = UIDialog()
    # 调用setupUi在指定窗口(主窗口)中添加控件
    ui.setupUi()
    # 显示窗口
    ui.show()

    # 进入程序的主循环，并通过exit函数确保主循环安全结束
    sys.exit(app.exec_())