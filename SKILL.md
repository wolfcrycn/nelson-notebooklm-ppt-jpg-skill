# Nelson — NotebookLM PPT/PDF 转 JPG 去水印拼接技能

## 功能说明

将 NotebookLM 导出的 PPT（.pptx）或 PDF（.pdf）转换为高清 JPG 图片（300 DPI），自动检测并清除右下角 NotebookLM 水印（logo），然后将所有页面垂直拼接为一张长图。

## 触发场景

- 用户上传 `.pptx` 或 `.pdf` 文件，并提到：
  - 「去水印」「清除 logo」「去掉右下角水印」
  - 「转成 JPG」「导出图片」
  - 「拼接成长图」「合成一张图」
  - 「处理这个 PPT/PDF」
- 用户提到「NotebookLM」相关内容

## 已知参数（NotebookLM 当前版本实测）

| 参数 | 值 | 说明 |
|------|-----|------|
| 导出 DPI | 300 | 高清导出（等效约 210 DPI / 5334×3000px） |
| Logo 实际位置 | (4935, 2908) – (5334, 2952) | 399×44px，右下角 |
| 清除区域（+30%扩展） | (4816, 2902) 518×56px | 覆盖 logo 及边缘残留 |
| 修复算法 | OpenCV Telea inpaint | 半径 5px，智能填充 |
| Logo 位置一致性 | 11/12 页一致 | 第 12 页需单独处理（匹配分 < 0.9 时扩大区域） |

⚠️ **NotebookLM 更新后 logo 位置可能变化，首次处理新文件时需重新定位。**

## 工作流程

```
用户上传 PPT/PDF
       ↓
  检测文件类型
   ↙           ↘
 PPT (.pptx)     PDF (.pdf)
   ↓                ↓
 LibreOffice      PyMuPDF
 转为 PDF           渲染
   ↓                ↓
   ─── 统一进入 ───
   ↓
 300 DPI JPG 渲染
   ↓
 模板匹配定位水印
 （跳过 if --skip-locate）
   ↓
 动态清除区域
 OpenCV Telea inpaint
   ↓
 质量验证
 (mean > 220 ✅)
   ↓
 保存 clean_01.jpg ~ clean_NN.jpg
   ↓
 垂直拼接 → clean_all_stitched.jpg
```

## 详细步骤

### Step 1 — 安装依赖

首次使用时，确保以下 Python 包已安装：

```bash
pip install pymupdf opencv-python numpy pillow
```

LibreOffice（PPTX 渲染推荐）：
- 下载：https://www.libreoffice.org/download/
- Windows 安装后自动注册 `soffice.exe` 路径

### Step 2 — 确定水印位置（首次）

用 Slide-1 右下角区域做模板匹配：

```python
import fitz, cv2, numpy as np
from PIL import Image

# Step 1: 渲染 PDF
doc = fitz.open("slides.pdf")
zoom = 300 / 72.0
pix = doc[0].get_pixmap(matrix=fitz.Matrix(zoom, zoom))
page = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

# Step 2: 提取 logo 模板（已知位置）
x1, y1, x2, y2 = 4935, 2908, 5334, 2952
template = np.array(page)[y1:y2, x1:x2]
t_gray = cv2.cvtColor(template, cv2.COLOR_RGB2GRAY)

# Step 3: 对所有页面做模板匹配
for i, pg in enumerate(pages):
    arr = np.array(pg)
    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    res = cv2.matchTemplate(gray, t_gray, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(res)
    print(f"Slide {i+1}: score={max_val:.3f} at {max_loc}")
```

- **阈值判断**：score > 0.9 → 位置可信，直接使用
- **分数低（< 0.9）**：该页面 logo 位置偏移，扩大搜索区域或手动指定清除区域

### Step 3 — 清除水印

```python
def clean_watermark(img_pil, clean_area=(4816, 2902, 518, 56)):
    x, y, w, h = clean_area
    arr = np.array(img_pil)
    mask = np.zeros(arr.shape[:2], dtype=np.uint8)
    mask[y:y+h, x:x+w] = 255
    bgr = cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)
    inpainted = cv2.inpaint(bgr, mask, 5, cv2.INPAINT_TELEA)
    return Image.fromarray(cv2.cvtColor(inpainted, cv2.COLOR_BGR2RGB))
```

### Step 4 — 质量验证

```python
x, y, w, h = 4816, 2902, 518, 56
region = np.array(cleaned)[y:y+h, x:x+w]
print(f"mean={region.mean():.1f}  std={region.std():.1f}")
# mean > 220 ✅（接近周围背景 ~242）
# mean < 180 ⚠️（仍有残留，扩展清除区域）
```

### Step 5 — 垂直拼接

```python
from PIL import Image

widths, heights = zip(*(img.size for img in images))
total_h = sum(heights)
canvas = Image.new("RGB", (max(widths), total_h), (255, 255, 255))
y = 0
for img in images:
    canvas.paste(img, (0, y)); y += img.height
canvas.save("clean_all_stitched.jpg", quality=95)
```

## 使用方式

### 通过 skill 直接执行

```
用户说「处理这个 PPT，去掉右下角水印」
→ 加载本 SKILL.md
→ 找到输入文件路径
→ 执行 process.py
→ 发送结果给用户
```

### 命令行执行

```bash
# 基本用法
py scripts/process.py "2026樱桃领导力成长伙伴计划.pptx"

# 指定清除区域
py scripts/process.py "slides.pdf" --clean-area 4816 2902 518 56

# 跳过模板匹配（已知位置，直接处理）
py scripts/process.py "slides.pdf" --skip-locate

# 指定输出目录
py scripts/process.py "slides.pdf" --out ./output
```

## 输出文件

| 文件 | 说明 |
|------|------|
| `clean_01.jpg` ~ `clean_NN.jpg` | 去水印后的单页高清图 |
| `clean_all_stitched.jpg` | 所有页面垂直拼接长图 |

## 常见问题

**Q: 模板匹配分数低（< 0.9）**
A: 该页面 logo 位置偏移，扩大搜索区域，或手动指定 `--clean-area 4700 2900 600 80`

**Q: 水印残留（mean < 180）**
A: 逐步扩大清除区域：518×56 → 600×80 → 700×100，直到前后对比干净

**Q: PPTX 渲染效果不好**
A: 安装 LibreOffice，脚本会自动优先使用 LibreOffice 方案（效果最好）

**Q: 长图发不出去（文件过大 > 25MB）**
A: 分段发送（每 4-6 页一段），或 quality 降到 85

**Q: 只处理部分页面**
A: 修改 process.py 中的 pages 切片：`pages = pages[0:6]`（前6页）

## 脚本结构

```
nelson-notebooklm-ppt-jpg-skill/
├── SKILL.md            ← 本文件
└── scripts/
    └── process.py      ← 主处理脚本（含渲染/去水印/拼接）
```
