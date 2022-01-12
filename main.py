# 这是一个示例 Python 脚本。
import pdfplumber
# 按 ⌃R 执行或将其替换为您的代码。
# 按 双击 ⇧ 在所有地方搜索类、文件、工具窗口、操作和设置。


def print_hi(path):
    # 在下面的代码行中使用断点来调试脚本。
    with pdfplumber.open(path) as pdf:
        content = ''
        for i in range(len(pdf.pages)):
            page = pdf.pages[i]
            page_content = '\n'.join(page.extract_text().split('\n')[:-1])
            content = content + page_content
        print(content)

# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':
    print_hi('/Users/zaochuan/Downloads/HomeKit Certification Test Cases R11.1.pdf')

# 访问 https://www.jetbrains.com/help/pycharm/ 获取 PyCharm 帮助
