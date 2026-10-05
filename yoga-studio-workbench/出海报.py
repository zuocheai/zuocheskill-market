# -*- coding: utf-8 -*-
"""把「底图」＋「中文文案」合成成成品海报。

为什么不让 AI 直接把中文画进图里：生成模型画中文字符常出错、变形、拼错。
正法＝**底图无字（AI 出）＋ 文字排版（本地精确合成）**。

用法：
    python 出海报.py --底图 底图.png --数据 海报数据.json --出 输出目录
"""
import argparse
import json
import os
import sys

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    sys.exit(
        "缺「Pillow」这个包，出不了海报图。\n"
        "补法：在命令行里跑这一条 ——  python -m pip install Pillow\n"
        "（装完重跑即可；拿不准就把这句话发给你正在用的 AI 助手。注意只装 Pillow，别去装那个叫 PIL 的老包。）"
    )

FONTS = {
    "regular": "C:/Windows/Fonts/msyh.ttc",
    "bold": "C:/Windows/Fonts/msyhbd.ttc",
    "black": "C:/Windows/Fonts/simhei.ttf",
}

# 安全区：文字只排在这些边距内，避免压到画面主体
MARGIN_RATIO = 0.075


def load_font(kind, size):
    path = FONTS.get(kind, FONTS["regular"])
    if not os.path.exists(path):
        path = FONTS["regular"]
    return ImageFont.truetype(path, size)


def wrap_cn(text, font, draw, max_w):
    """按像素宽度给中文断行（逐字试宽）。"""
    if not text:
        return []
    lines, cur = [], ""
    for ch in text:
        if ch == "\n":
            lines.append(cur)
            cur = ""
            continue
        trial = cur + ch
        w = draw.textlength(trial, font=font)
        if w > max_w and cur:
            lines.append(cur)
            cur = ch
        else:
            cur = trial
    if cur:
        lines.append(cur)
    return lines


def build(base_img, data, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    W, H = base_img.size
    m = int(min(W, H) * MARGIN_RATIO)
    max_w = W - m * 2

    # 半透明文字底板，让白字在亮底图上也读得清
    panel_top = int(data.get("面板起始比", 0.52) * H)
    panel = Image.new("RGBA", (W, H - panel_top), (255, 255, 255, 0))
    panel_draw = ImageDraw.Draw(panel)
    alpha = int(data.get("面板不透明度", 0.88) * 255)
    panel_draw.rectangle([0, 0, W, H - panel_top],
                         fill=data.get("面板颜色", "#FFFFFF") + f"{alpha:02X}")
    base_img = base_img.convert("RGBA")
    base_img.alpha_composite(panel, (0, panel_top))

    draw = ImageDraw.Draw(base_img)
    y = panel_top + m

    def put(text, kind, size_ratio, color, spacing=1.5, bold=False):
        nonlocal y
        if not text:
            return
        fsize = int(H * size_ratio)
        font = load_font("bold" if bold else kind, fsize)
        for line in wrap_cn(text, font, draw, max_w):
            tw = draw.textlength(line, font=font)
            draw.text(((W - tw) / 2, y), line, font=font, fill=color)
            y += fsize * spacing
        y += fsize * 0.35

    put(data.get("主标题"), "bold", 0.055, data.get("主色", "#1A1A1A"), 1.35, True)
    put(data.get("副标题"), "regular", 0.032, data.get("副色", "#444444"))
    put(data.get("卖点"), "regular", 0.030, data.get("正文色", "#333333"), 1.7)
    put(data.get("价格"), "bold", 0.046, data.get("强调色", "#B8574A"), 1.4, True)
    put(data.get("时间地点"), "regular", 0.028, data.get("正文色", "#333333"), 1.7)
    put(data.get("行动指令"), "bold", 0.034, data.get("行动色", "#1A1A1A"), 1.5, True)

    out = os.path.join(out_dir, data.get("文件名", "海报.png"))
    base_img.convert("RGB").save(out, quality=95)
    return out


DEFAULT_DATA = {
    "主标题": "9.9 元 体验课",
    "副标题": "新学员专属 · 含体测",
    "卖点": "教龄 8 年｜一对一私教｜不限次数",
    "价格": "￥9.9",
    "时间地点": "本周六上午 10:00 · 学院路店",
    "行动指令": "私我「体验」两个字，我给你排时间",
    "主色": "#1A1A1A",
    "副色": "#555555",
    "正文色": "#333333",
    "强调色": "#B8574A",
    "行动色": "#1A1A1A",
    "面板颜色": "#FFFFFF",
    "面板不透明度": 0.88,
    "面板起始比": 0.52,
    "文件名": "海报_成品.png",
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--底图", required=True)
    ap.add_argument("--数据", help="海报数据 JSON；不给就用内置示例")
    ap.add_argument("--出", default=".")
    a = ap.parse_args()

    if not os.path.exists(a.底图):
        print("底图不存在:", a.底图)
        sys.exit(1)

    data = dict(DEFAULT_DATA)
    if a.数据 and os.path.exists(a.数据):
        with open(a.数据, "r", encoding="utf-8") as f:
            data.update(json.load(f))

    img = Image.open(a.底图)
    out = build(img, data, a.出)
    print("已生成:", out)
    print("大小:", round(os.path.getsize(out) / 1024, 1), "KB")


if __name__ == "__main__":
    main()
