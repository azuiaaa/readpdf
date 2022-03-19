import pandas as pd
import numpy as np
# 构造数据框

if __name__ == '__main__':
    np.random.seed(24)
    df = pd.DataFrame(np.random.randn(6, 4), columns=list('ABCD'))
    df.iloc[3, 3] = np.nan
    df.iloc[0, 2] = np.nan


    def style_func(x) -> str:
        # 元素颜色
        color = 'red' if x > 0 else ''
        # 元素字体
        weight = 'bold' if x > 0 else 'normal'
        # 返回样式
        return ';'.join([f'color:{color}', f'font-weight:{weight}'])


    df.style.bar(subset=['A'], align='mid', color=['#d65f5f', '#5fba7d']). \
        background_gradient(subset=['B']). \
        highlight_null(subset=['C']). \
        applymap(style_func, subset=['D']). \
        format('{:.2f}')
    df.style.highlight_max()
    print(df.style.render())