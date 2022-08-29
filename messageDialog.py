from PyQt5.QtWidgets import QWidget, QMessageBox


class UI_Message(QWidget):
    def MessageInformation(dialog, title="Info", text="消息提示"):
        # 弹出消息对话框
        reply = QMessageBox.information(dialog, title, text, QMessageBox.Ok)
        return reply

    def MessageQuestion(dialog, title="Question", text="消息询问"):
        # 弹出消息对话框
        reply = QMessageBox.question(dialog, title, text, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        return reply

    def MessageWarning(dialog, title="Warning", text="消息警告"):
        # 弹出消息对话框
        reply = QMessageBox.warning(dialog, title, text, QMessageBox.Ok)
        return reply

    def MessageCritical(dialog, title="Critical", text="严重错误警告"):
        # 弹出消息对话框
        reply = QMessageBox.critical(dialog, title, text, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        return reply