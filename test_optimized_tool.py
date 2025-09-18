#!/usr/bin/env python3
"""
测试优化后的文件重命名工具界面
"""

import tkinter as tk
from file_rename_tool_optimized import FileRenameTool

if __name__ == "__main__":
    print("启动优化后的文件重命名工具...")
    print("\n主要优化内容：")
    print("1. 减少了各个组件之间的内边距（padding从10改为5，行间距从5改为2）")
    print("2. 将规则配置改为单行布局，节省垂直空间")
    print("3. 将安全提示移到按钮行，节省一行空间")
    print("4. 减小了按钮宽度，使用更紧凑的布局")
    print("5. 调整PanedWindow权重：预览区域权重从3增加到5，日志区域从1保持不变")
    print("6. 设置Treeview初始高度为20行，日志区域初始高度为5行")
    print("7. 优化了标签文字，使用更简洁的描述")
    print("\n结果：预览区域占据了更大比例的界面空间，整体布局更加紧凑高效")
    
    try:
        root = tk.Tk()
        app = FileRenameTool(root)
        root.mainloop()
    except Exception as e:
        print(f"\n错误：{e}")
        input("\n按Enter键退出...")