import io, base64, os
from jinja2 import Environment, FileSystemLoader
import tkinter as tk
from tkinter import filedialog

def fig_to_base64(fig):
    """matplotlib 图 → base64 字符串"""
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=300, bbox_inches='tight')
    buf.seek(0)
    return base64.b64encode(buf.read()).decode()

def export_html(data, fig_map: dict):
    """
    data      : get_monthly_data() 返回的 dict
    fig_map   : dict(key=模板里变量名, value=matplotlib.figure.Figure)
    保存路径由用户选择
    """
    # ---- 让用户选保存路径 ----
    root = tk.Tk()
    root.withdraw()  # 不显示主窗口
    out_html = filedialog.asksaveasfilename(
        title='保存FAST月度使用报告',
        defaultextension='.html',
        initialfile='FAST_usage_report.html',
        filetypes=[('HTML 文件', '*.html'), ('所有文件', '*.*')]
    )
    if not out_html:          # 用户点了“取消”
        return

    template_dir = r'J:\Engineering\ShareFolder\new_ABB_Production_Tools\Pd\html_template'
    env = Environment(loader=FileSystemLoader(template_dir))
    tpl = env.get_template('report_tpl.html')

    # 把图片转成 base64
    img_ctx = {k: fig_to_base64(fig) for k, fig in fig_map.items()}

    # 把字符串列表转成真正的 list
    for key in ['terminal_top3_projects', 'hole_top3_projects', 'size_top3_projects',
                'wire_top3_projects', 'bom_top3_projects', 'bom_compare_eplan_top3']:
        data[key] = [eval(lst) for lst in data[key]]

    html = tpl.render(**data, **img_ctx)
    with open(out_html, 'w', encoding='utf-8') as f:
        f.write(html)

    tk.messagebox.showinfo('完成', f'用户使用情况报告已保存至\n{os.path.abspath(out_html)}')