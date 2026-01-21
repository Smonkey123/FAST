import PyPDF2
import fitz
import re
import numpy as np

def extract_terminal_diagram_from_pdf(pdf_path):
    td_page, td_name, td_str, td_list = [], [], [], []

    with open(pdf_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        outlines = reader.outline
        # 注意：outline 是一维列表，原脚本 outlines[3] 会抛 IndexError
        if len(outlines) > 3:
            for idx, item in enumerate(outlines[3]):
                title = str(item.get("/Title", ""))
                if "&T/" in title and ("端子图表" in title or "TERMINAL DIAGRAM" in title.upper()):
                    td_page.append(idx)
                    td_name.append(title)
                    td_str.append(title.split(" ")[-1])
                    td_list.append((idx, title, title.split(" ")[-1]))

    return td_list

def check_separator_plate(pdf_path, td_list, out_path):
    short_list = []    # 短接片跨端子
    miss_list = []    # 隔板缺失

    with fitz.open(pdf_path) as pdf:
        for idx, _ , string in td_list:
            page = pdf.load_page(idx)
            # 当前页端子排名称
            typical_name = string.split("+-")[-2].split("=")[-1]
            strip_name = string.split("+-")[-1]
            w = page.rect.width
            scale = w / 1207.98
            # print(f"第{idx + 1}页 宽度={w:.2f} pt  缩放比例={scale:.3f}")

            # ===== 动态参数 =====
            DOT_MIN_W = 5 * scale  # 最小外接框宽（pt）
            DOT_MAX_W = 8 * scale  # 最大外接框宽（pt）
            Y_TOL = 2.0 * scale  # 同一行 y 差阈值（pt）
            MIN_DIST = 20 * scale  # 连心线最小长度
            EPS = 0.5  # 点到线段距离容忍
            ASPECT_TOL = 1.01  # 长宽比容忍度
            x_height = 20 * scale  # 检查高度

            # 提前拿全页文字
            words = page.get_text("words")

            def find_terminal_number(cx, cy):
                TERM_PAT = re.compile(r'^[A-Z0-9]{1,}$')  # 仅大写字母+数字，长度≥1

                """返回离(cx,cy)最近的端子号（数字/字母/混合）"""
                r = DOT_MAX_W * 1.2
                best = None, 1e9
                for w in words:
                    txt = w[4].strip()
                    if TERM_PAT.fullmatch(txt):  # 符合端子号格式
                        wx, wy = (w[0] + w[2]) / 2, (w[1] + w[3]) / 2
                        d = abs(wx - cx) + abs(wy - cy)
                        if d < best[1]:
                            best = txt, d
                return best[0] if best[0] else None

            # ------ 1. 先收当前页全部实心黑点 ------
            dots = []  # [(cx, cy), ...]
            for d in page.get_drawings():
                if d.get("fill") != (0, 0, 0):
                    continue
                if d.get("stroke") and d.get("width", 0) > 0.3:
                    continue
                r = d["rect"]
                w, h = r.width, r.height
                if not (DOT_MIN_W <= w <= DOT_MAX_W and DOT_MIN_W <= h <= DOT_MAX_W):
                    continue
                if max(w, h) / min(w, h) > ASPECT_TOL:
                    continue
                cx = (r.x0 + r.x1) / 2
                cy = (r.y0 + r.y1) / 2
                dots.append((cx, cy))
                # 可选：标黄点
                page.draw_circle((cx, cy), (r.x1-r.x0)/2, color=(0, 1, 0), width=1)

            if not dots:
                main_dots = []
            else:
                y_vals = np.array([cy for _, cy in dots])
                uni = np.unique(y_vals)
                counts = np.array([np.count_nonzero(y_vals == v) for v in uni])

                peak_i = int(counts.argmax())
                main_y = uni[peak_i]

                # 安全计算半宽容差
                if uni.size == 1:
                    y_tol_h = 1.0
                else:
                    y_tol_h = 0.5 * float(np.min(np.diff(uni)))

                main_dots = [(cx, cy) for cx, cy in dots
                             if abs(cy - main_y) < y_tol_h + Y_TOL]
                main_dots.sort(key=lambda p: p[0])
                # print(f"第{idx + 1}页 主轴线 y={main_y:.1f}  保留黑点 {len(main_dots)} 个")

            # 按 x 升序，保证从左往右
            dots.sort(key=lambda p: p[0])

            EPS = 0.5  # 距离容忍 pt

            def dist_pt_segment(p, a, b):
                """点到线段距离（纯数值版）"""
                ab_x, ab_y = b.x - a.x, b.y - a.y
                ap_x, ap_y = p.x - a.x, p.y - a.y
                # 投影参数 t
                ab2 = ab_x * ab_x + ab_y * ab_y
                if ab2 == 0:
                    return (ap_x * ap_x + ap_y * ap_y) ** 0.5
                t = max(0., min(1., (ab_x * ap_x + ab_y * ap_y) / ab2))
                proj_x = a.x + t * ab_x
                proj_y = a.y + t * ab_y
                return ((p.x - proj_x) ** 2 + (p.y - proj_y) ** 2) ** 0.5

            # ------ 2. 统计每两个黑点之间存在的实际线段长度 ------
            # 先收集当前页所有黑色线段
            black_segments = []
            for d in page.get_drawings():
                if d.get("color") != (0, 0, 0):
                    continue
                for cmd in d["items"]:
                    if cmd[0] != "l":
                        continue
                    p0, p1 = fitz.Point(cmd[1]), fitz.Point(cmd[2])
                    black_segments.append((p0, p1))

            # 对每条黑线段，找出落在它上的所有黑点
            for p0, p1 in black_segments:
                on_line = []
                for cx, cy in dots:
                    if dist_pt_segment(fitz.Point(cx, cy), p0, p1) < EPS:
                        on_line.append((cx, cy))
                # 按 x 排序，相邻即有效对
                on_line.sort()
                for (x1, y1), (x2, y2) in zip(on_line, on_line[1:]):
                    seg_len = x2 - x1
                    t1 = find_terminal_number(x1, y1) or "?"
                    t2 = find_terminal_number(x2, y2) or "?"

                    # print(f"黑点间有线段: {typical_name}——{strip_name}:{t1}-{t2}  长度={seg_len:.2f} pt")
                    # 可选：把实际截断区间画成红色

                    if seg_len > MIN_DIST:
                        page.draw_line(fitz.Point(x1, y1), fitz.Point(x2, y2), color=(1, 0, 0), width=1.2)
                        page.draw_circle(fitz.Point(x1, y1), 1, color=(1, 0, 0), width=0.8)
                        page.draw_circle(fitz.Point(x2, y2), 1, color=(1, 0, 0), width=0.8)

                        # print(f"{typical_name}——{strip_name}:{t1}-{t2}【短接片跨端子】")
                        # print(f"第{idx + 1}页{strip_name}:{t1}-{t2}【短接片跨端子】")
                        # print(f"黑点({x1:.2f},{y1}) -> ({x2:.2f},{y2})间出现了短接片跨端子")

                        short_list.append({
                            "page": idx + 1,
                            "typical": typical_name,
                            "strip": strip_name,
                            "term_pair": f"{t1}-{t2}",
                            "coord": (x1, y1, x2, y2),  # 如外部还要坐标
                        })

            # ------ 3. 快速看出「哪两段黑点之间缺线」 ------
            # 先建集合：已有线段的两端坐标对（保留 1 位小数防浮点误差）
            exist_pairs = set()
            for p0, p1 in black_segments:
                on_line = sorted([(cx, cy) for cx, cy in dots
                                  if dist_pt_segment(fitz.Point(cx, cy), p0, p1) < EPS])
                for a, b in zip(on_line, on_line[1:]):
                    exist_pairs.add((round(a[0], 1), round(a[1], 1), round(b[0], 1), round(b[1], 1)))
            # print(exist_pairs)
            # 与理论相邻黑点对比
            for (x1, y1), (x2, y2) in zip(main_dots, main_dots[1:]):
                if (round(x1, 1), round(y1, 1), round(x2, 1), round(y2, 1)) not in exist_pairs:
                    t1 = find_terminal_number(x1, y1) or "?"
                    t2 = find_terminal_number(x2, y2) or "?"
                    # print(f"第{idx + 1}页{strip_name}:{t1}-{t2}【端子间无黑线段】")
                    # print(f"第{idx+1}页【缺线】黑点间无黑线段: ({x1:.2f},{y1}) -> ({x2:.2f},{y2})")

            # ------ 4. 在无黑线段上方区域检查字符 X ------
            # 先拿当前页所有单词（含坐标）
            words = page.get_text("words")   # [(x0,y0,x1,y1,"text"), ...]
            # print(main_dots)
            for (x1, y1), (x2, y2) in zip(main_dots, main_dots[1:]):
                if (round(x1,1), round(y1,1), round(x2,1), round(y2,1)) in exist_pairs:
                    continue          # 有线段就跳过
                # 检查矩形：左右 x1~x1+MIN_DIST，底 y1，高 x_height
                check_rect = fitz.Rect(x1, y1 - x_height, x1+MIN_DIST, y1)

                x_list = [w for w in words
                          if w[4] == "X" and fitz.Rect(w[:4]) in check_rect]
                if x_list:
                    # print(f"第{idx+1}页【缺线】端子 ({x1:.2f},{y1})->({x2:.2f},{y2}) 上方发现 X 标记")
                    # 可选：把 X 框出来
                    for w in x_list:
                        page.draw_rect(fitz.Rect(w[:4]), color=(0,1,1), width=1.2)
                else:
                    page.draw_circle(fitz.Point(x1, y1), 1, color=(1, 0, 0), width=0.8)
                    page.draw_circle(fitz.Point(x2, y2), 1, color=(1, 0, 0), width=0.8)
                    page.draw_rect(check_rect, color=(1, 0, 1), width=1.2)
                    t1 = find_terminal_number(x1, y1) or "?"
                    t2 = find_terminal_number(x2, y2) or "?"

                    # print(f"{typical_name}——{strip_name}:{t1}【隔板缺失】")
                    # print(f"第{idx + 1}页{strip_name}:{t1}-{t2}【隔板缺失】")
                    # print(f"第{idx+1}页【缺线】端子 ({x1:.2f},{y1})->({x2:.2f},{y2}) 上方无 X")

                    miss_list.append({
                        "page": idx + 1,
                        "typical": typical_name,
                        "strip": strip_name,
                        "term_pair": t1,
                        "coord": (x1, y1, x2, y2),
                    })

            if main_dots:
                x, y = main_dots[-1]  # 末端
                check_rect = fitz.Rect(x, y - x_height, x + MIN_DIST, y)
                x_list = [w for w in words if w[4] == "X" and fitz.Rect(w[:4]) in check_rect]
                if x_list:
                    for w in x_list:
                        page.draw_rect(fitz.Rect(w[:4]), color=(0, 1, 1), width=1.2)
                else:
                    page.draw_circle(fitz.Point(x, y), 1, color=(1, 0, 0), width=0.8)
                    page.draw_rect(check_rect, color=(1, 0, 1), width=1.2)
                    t = find_terminal_number(x, y) or "?"

                    # 1. 必须落在某条黑线段上
                    # 2. 该点到线段起点距离 < 到终点距离
                    near_left = False
                    for p0, p1 in black_segments:
                        if dist_pt_segment(fitz.Point(x, y), p0, p1) < EPS:
                            # 新增：终点必须在黑点右侧，且水平距离 ≥ 1/3*MIN_DIST
                            if (p1.x - x) >= MIN_DIST / 3:
                                near_left = True
                                break

                    if near_left:  # 靠近左侧 → 跨页，不报缺失
                        continue

                        # 下方仍是原逻辑
                    # print(f"{typical_name}——{strip_name}:{t}【隔板缺失】")
                    miss_list.append({
                        "page": idx + 1,
                        "typical": typical_name,
                        "strip": strip_name,
                        "term_pair": t,
                        "coord": (x, y, x, y),
                    })

        pdf.save(out_path)
    return {"short": short_list, "missing": miss_list}




if __name__ == "__main__":
    pdf_path = r'C:\Users\CNXIGAO13\Desktop\Terminal_SP\504556032.pdf'
    out_path = r"C:\Users\CNXIGAO13\Desktop\Terminal_SP\504556032_annotated.pdf"

    td_list = extract_terminal_diagram_from_pdf(pdf_path)
    # print(td_list)
    check_separator_plate(pdf_path, td_list, out_path)