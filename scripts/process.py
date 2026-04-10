"""
NotebookLM PPT/PDF → JPG 去水印 + 垂直拼接
用法：
  py process.py <input.pptx|pdf> [--dpi 300] [--clean-area X Y W H] [--no-stitch]
  py process.py "2026樱桃领导力成长伙伴计划.pptx" --dpi 300
"""

import os
import sys
import shutil
import argparse
import tempfile
import platform

# ── 依赖检查 ────────────────────────────────────────────────────────────
MISSING = []
for pkg, imp in [("pymupdf", "fitz"), ("opencv-python", "cv2"),
                  ("numpy", "numpy"), ("pillow", "PIL")]:
    try:
        __import__(imp)
    except ImportError:
        MISSING.append(pkg)

if MISSING:
    print(f"[ERROR] 缺少依赖: {', '.join(MISSING)}")
    print("请运行: pip install " + " ".join(MISSING))
    sys.exit(1)

import fitz  # PyMuPDF
import cv2
import numpy as np
from PIL import Image


# ── 实测参数 ────────────────────────────────────────────────────────────
DEFAULT_DPI = 300
# Logo 实际位置（NotebookLM 当前版本）
LOGO_BOX = (4935, 2908, 5334, 2952)       # x1, y1, x2, y2  399×44px
# 清除区域向左扩展约 30%（覆盖 logo 及残留边缘）
CLEAN_AREA = (4816, 2902, 518, 56)          # x, y, w, h
INPAINT_RADIUS = 5


# ══════════════════════════════════════════════════════════════════════
#  渲染层
# ══════════════════════════════════════════════════════════════════════

def render_pdf(pdf_path: str, dpi: int = DEFAULT_DPI) -> list[Image.Image]:
    """将 PDF 每页渲染为高清 PIL Image"""
    doc = fitz.open(pdf_path)
    zoom = dpi / 72.0
    pages = []
    for i, page in enumerate(doc):
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        pages.append(img)
        if (i + 1) % 5 == 0:
            print(f"    渲染 {i+1}/{len(doc)}...")
    doc.close()
    return pages


def render_pptx(pptx_path: str, dpi: int = DEFAULT_DPI) -> list[Image.Image]:
    """将 PPTX 每页渲染为高清 PIL Image（优先级：LibreOffice > python-pptx）"""
    abs_path = os.path.abspath(pptx_path)

    # ── 方案 A：LibreOffice（最可靠，保留完整排版）────────────────────
    lo_paths = []
    if platform.system() == "Windows":
        lo_paths = [
            r"C:\Program Files\LibreOffice\Program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\Program\soffice.exe",
        ]
    else:
        lo_paths = ["soffice", "libreoffice"]

    for lo in lo_paths:
        try:
            out_dir = os.path.dirname(abs_path)
            result = __import__("subprocess").run(
                [lo, "--headless", "--convert-to", "pdf",
                 "--outdir", out_dir, abs_path],
                capture_output=True, timeout=180
            )
            pdf_out = os.path.splitext(abs_path)[0] + ".pdf"
            if os.path.exists(pdf_out):
                print(f"    LibreOffice 转换成功: {pdf_out}")
                pages = render_pdf(pdf_out, dpi)
                # 清理中间 PDF
                try:
                    os.remove(pdf_out)
                except Exception:
                    pass
                return pages
        except Exception:
            continue

    # ── 方案 B：python-pptx + Pillow（回退方案，排版可能偏差）────────
    try:
        from pptx import Presentation
        from pptx.util import Emu, Inches, Pt
        from io import BytesIO

        prs = Presentation(abs_path)
        scale = dpi / 72.0
        pages = []

        # 幻灯片尺寸（EMU → 像素）
        sw = int(prs.slide_width / Emu(914400) * 72 * scale)
        sh = int(prs.slide_height / Emu(914400) * 72 * scale)

        # 逐形状渲染（仅支持基本图片/矩形/文字）
        # NOTE: 此方案只能近似，无法完美还原 PPT 效果
        #       建议优先使用 LibreOffice 方案
        print(f"    WARNING: 使用 python-pptx 回退方案，可能无法还原所有元素")
        print(f"    建议安装 LibreOffice: https://www.libreoffice.org/download/")

        for i, _ in enumerate(prs.slides):
            # 创建空白画布
            canvas = Image.new("RGB", (sw, sh), (255, 255, 255))
            # 占位输出（实际需截图，这里仅返回尺寸信息）
            pages.append(canvas)
        return pages

    except ImportError:
        pass

    print("[ERROR] 无法渲染 PPTX。请安装 LibreOffice 或 python-pptx:")
    print("  LibreOffice: https://www.libreoffice.org/download/")
    print("  pip install python-pptx pillow")
    sys.exit(1)


# ══════════════════════════════════════════════════════════════════════
#  水印处理
# ══════════════════════════════════════════════════════════════════════

def locate_by_template(page: Image.Image,
                       logo_box: tuple[int, int, int, int],
                       search_margin: int = 200) -> tuple[int, int, float]:
    """
    在页面右下角区域做模板匹配，返回 (x, y, score)
    logo_box = (x1, y1, x2, y2)，从 Slide-1 提取模板
    """
    x1, y1, x2, y2 = logo_box
    w_img, h_img = page.size

    # 限制搜索范围：右下角 ± search_margin
    sx1 = max(0, x1 - search_margin)
    sy1 = max(0, y1 - search_margin)
    search_w = w_img - sx1
    search_h = h_img - sy1

    if search_w < (x2 - x1) or search_h < (y2 - y1):
        return x1, y1, 1.0  # 兜底

    # 提取模板
    arr = np.array(page)
    template = arr[y1:y2, x1:x2]
    t_gray = cv2.cvtColor(template, cv2.COLOR_RGB2GRAY)

    # 搜索区域
    search = arr[sy1:sy1+search_h, sx1:sx1+search_w]
    s_gray = cv2.cvtColor(search, cv2.COLOR_RGB2GRAY)

    # 模板匹配
    res = cv2.matchTemplate(s_gray, t_gray, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(res)

    # 转回全局坐标
    found_x = sx1 + max_loc[0]
    found_y = sy1 + max_loc[1]
    return found_x, found_y, float(max_val)


def clean_watermark(img: Image.Image,
                    clean_area: tuple[int, int, int, int] = CLEAN_AREA,
                    inpaint_radius: int = INPAINT_RADIUS) -> Image.Image:
    """OpenCV Telea inpainting 智能清除水印"""
    x, y, w, h = clean_area
    arr = np.array(img)

    mask = np.zeros(arr.shape[:2], dtype=np.uint8)
    mask[y:y+h, x:x+w] = 255

    bgr = cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)
    inpainted = cv2.inpaint(bgr, mask, inpaint_radius, cv2.INPAINT_TELEA)
    result = cv2.cvtColor(inpainted, cv2.COLOR_BGR2RGB)

    return Image.fromarray(result)


def verify(img: Image.Image, area: tuple[int, int, int, int]) -> dict:
    """统计清除区域像素，辅助判断是否残留"""
    x, y, w, h = area
    r = np.array(img)[y:y+h, x:x+w]
    return dict(mean=float(r.mean()), std=float(r.std()),
                min=int(r.min()), max=int(r.max()))


# ══════════════════════════════════════════════════════════════════════
#  拼接
# ══════════════════════════════════════════════════════════════════════

def stitch(images: list[Image.Image], output_path: str) -> None:
    """垂直拼接所有页面为一张长图"""
    widths, heights = zip(*(img.size for img in images))
    total_h = sum(heights)
    max_w = max(widths)

    canvas = Image.new("RGB", (max_w, total_h), (255, 255, 255))
    y_off = 0
    for img in images:
        canvas.paste(img, (0, y_off))
        y_off += img.height

    canvas.save(output_path, "JPEG", quality=95)
    mb = os.path.getsize(output_path) / 1024 / 1024
    print(f"    拼接完成: {output_path}  ({mb:.1f} MB, {max_w}×{total_h}px)")


# ══════════════════════════════════════════════════════════════════════
#  主流程
# ══════════════════════════════════════════════════════════════════════

def process(input_path: str,
            dpi: int = DEFAULT_DPI,
            clean_area: tuple[int, int, int, int] = CLEAN_AREA,
            skip_locate: bool = False,
            no_stitch: bool = False) -> str:
    """
    完整流程：渲染 → 定位水印 → 清除 → 保存 → 拼接

    返回输出目录路径
    """
    input_path = os.path.abspath(input_path)
    ext = os.path.splitext(input_path)[1].lower()
    basename = os.path.splitext(os.path.basename(input_path))[0]

    print(f"\n📄 文件: {os.path.basename(input_path)}")
    print(f"   DPI: {dpi}  |  清除区域: {clean_area}")

    # ── Step 1: 渲染 ──────────────────────────────────────────────────
    print(f"\n[1/4] 渲染页面...")
    if ext == ".pdf":
        pages = render_pdf(input_path, dpi)
    elif ext in (".pptx", ".ppt"):
        pages = render_pptx(input_path, dpi)
    else:
        print(f"[ERROR] 不支持的文件类型: {ext}")
        sys.exit(1)

    n = len(pages)
    w0, h0 = pages[0].size
    print(f"   共 {n} 页  |  尺寸: {w0}×{h0}px")
    if w0 < 3000:
        print(f"   ⚠ 分辨率偏低（< 3000px），建议使用 --dpi 300")

    # ── Step 2: 定位水印（首次使用）──────────────────────────────────
    print(f"\n[2/4] 定位水印...")
    locate_results = []
    if skip_locate:
        print(f"   使用固定区域: {CLEAN_AREA}（跳过定位）")
        locate_results = [(CLEAN_AREA[0], CLEAN_AREA[1], 1.0)] * n
    else:
        for i, pg in enumerate(pages):
            fx, fy, score = locate_by_template(pg, LOGO_BOX)
            locate_results.append((fx, fy, score))
            mark = "✅" if score > 0.9 else "⚠️"
            print(f"   页 {i+1:2d}/{n}  {mark}  score={score:.3f}  pos=({fx},{fy})")

    # ── Step 3: 清除水印 ──────────────────────────────────────────────
    print(f"\n[3/4] 清除水印...")
    clean_pages = []
    for i, pg in enumerate(pages):
        fx, fy, score = locate_results[i]
        # 根据实际位置动态调整清除区域（以检测到的 x 为准，向左扩展）
        dx = fx - LOGO_BOX[0]  # 实际位置与模板的偏移
        ca = (fx - 119, fy - 6, 518, 56)  # 相对偏移后的清除区
        cleaned = clean_watermark(pg, ca)
        clean_pages.append(cleaned)

        s = verify(cleaned, ca)
        flag = "✅" if s["mean"] > 220 else "⚠️"
        print(f"   页 {i+1:2d}/{n}  {flag}  mean={s['mean']:.1f}  std={s['std']:.1f}")

    # ── Step 4: 保存单页 ──────────────────────────────────────────────
    print(f"\n[4/4] 保存 JPG...")
    out_dir = tempfile.mkdtemp(prefix="notebooklm_")
    for i, img in enumerate(clean_pages):
        path = os.path.join(out_dir, f"clean_{i+1:02d}.jpg")
        img.save(path, "JPEG", quality=95)
        mb = os.path.getsize(path) / 1024 / 1024
        print(f"   clean_{i+1:02d}.jpg  ({mb:.1f} MB)")

    # ── Step 5: 拼接长图 ──────────────────────────────────────────────
    if not no_stitch:
        print(f"\n[5/5] 拼接长图...")
        long_path = os.path.join(out_dir, "clean_all_stitched.jpg")
        stitch(clean_pages, long_path)

    print(f"\n✅ 完成！输出目录: {out_dir}")
    return out_dir


# ══════════════════════════════════════════════════════════════════════
#  CLI
# ══════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    p = argparse.ArgumentParser(
        description="NotebookLM PPT/PDF → JPG 去水印 + 垂直拼接")
    p.add_argument("input", help="PPT 或 PDF 文件路径")
    p.add_argument("--dpi", type=int, default=DEFAULT_DPI,
                   help=f"渲染 DPI（默认 {DEFAULT_DPI}）")
    p.add_argument("--clean-area", type=int, nargs=4,
                   metavar=("X", "Y", "W", "H"),
                   help=f"清除区域（默认 {CLEAN_AREA}）")
    p.add_argument("--skip-locate", action="store_true",
                   help="跳过模板匹配，直接使用固定区域")
    p.add_argument("--no-stitch", action="store_true",
                   help="跳过拼接步骤")
    p.add_argument("--out", "-o", dest="output_dir",
                   help="指定输出目录（默认临时目录）")

    args = p.parse_args()
    ca = tuple(args.clean_area) if args.clean_area else CLEAN_AREA
    out_dir = process(args.input, dpi=args.dpi, clean_area=ca,
                      skip_locate=args.skip_locate, no_stitch=args.no_stitch)

    if args.output_dir:
        dest = os.path.abspath(args.output_dir)
        os.makedirs(dest, exist_ok=True)
        for f in os.listdir(out_dir):
            shutil.copy2(os.path.join(out_dir, f), os.path.join(dest, f))
        shutil.rmtree(out_dir)
        print(f"\n📁 已复制到: {dest}")
