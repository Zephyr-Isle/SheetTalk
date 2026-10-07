"""生成 assets/icon.ico(与窗口/托盘/侧边栏按钮同一设计,实现见 app/icon_design.py)。

用法: cd 项目根 && python assets/make_icon.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app.icon_design import write_ico

if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    out = write_ico(os.path.join(here, "icon.ico"))
    print("已生成:", out, "(%d 字节)" % os.path.getsize(out))
