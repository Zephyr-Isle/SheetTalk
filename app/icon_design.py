"""统一图标设计的光栅化实现:窗口/托盘/快捷方式/exe 内嵌/侧边栏按钮共用。

设计(64 单位坐标系,四周留 2 单位边距):
  - 圆角方块,填充对角线性渐变 #107c41 → #0b5c30
  - 白色网格(α235):三条横线、两条竖线,圆头端点
  - 方块外全透明

纯标准库。历史教训:渲染时必须把像素坐标换算回 64 单位设计空间
(u = (i+0.5) * 64 / n),否则大尺寸下内容只画在左上角 60×60 的区域里。
"""
import struct
import zlib

GRAD1 = (0x10, 0x7C, 0x41)
GRAD2 = (0x0B, 0x5C, 0x30)

# 网格线段 (x0, y0, x1, y1),64 单位坐标
SEGMENTS = [
    (16, 20, 48, 20), (16, 32, 48, 32), (16, 44, 48, 44),
    (26, 16, 26, 48), (38, 16, 38, 48),
]
PEN_HALF = 1.5            # 线宽 3 的一半
BOX_HALF = 28.0           # 圆角方块半边长 (60-4)/2
CORNER_R = 14.0
WHITE_A = 235 / 255.0
FILL = 60.0 / 56.0        # 画布占比:整体放大,四周只剩 2/64 边距


def _clamp(v, lo, hi):
    return lo if v < lo else (hi if v > hi else v)


def _seg_dist(px, py, seg):
    x0, y0, x1, y1 = seg
    dx, dy = x1 - x0, y1 - y0
    l2 = dx * dx + dy * dy
    t = 0.0 if l2 == 0 else _clamp(((px - x0) * dx + (py - y0) * dy) / l2, 0.0, 1.0)
    ex, ey = x0 + t * dx - px, y0 + t * dy - py
    return (ex * ex + ey * ey) ** 0.5


def _pixel(u, v):
    """返回 (r, g, b, a)。u/v 必须是 64 单位设计空间中的浮点坐标。

    缩放方向注意:设计内容画在 [4,60](56 单位),要放大到 [2,62](60 单位)呈现,
    采样坐标需「压缩」:d = 32 + (u-32)/FILL。写成乘号会把图形越改越小。
    """
    u = 32.0 + (u - 32.0) / FILL
    v = 32.0 + (v - 32.0) / FILL
    # 圆角方块 SDF(对称,取第一象限再折算)
    qx, qy = abs(u - 32.0), abs(v - 32.0)
    cx, cy = max(qx - (BOX_HALF - CORNER_R), 0.0), max(qy - (BOX_HALF - CORNER_R), 0.0)
    dist = (cx * cx + cy * cy) ** 0.5 + CORNER_R - BOX_HALF  # <0 在内部
    a_box = _clamp(0.5 - dist, 0.0, 1.0)
    if a_box <= 0.0:
        return (0, 0, 0, 0.0)
    t = _clamp((u + v) / 128.0, 0.0, 1.0)
    r = GRAD1[0] + (GRAD2[0] - GRAD1[0]) * t
    g = GRAD1[1] + (GRAD2[1] - GRAD1[1]) * t
    b = GRAD1[2] + (GRAD2[2] - GRAD1[2]) * t
    # 网格:任一线段的距离 ≤ 半线宽即白,带 1px 抗锯齿边
    d = min(_seg_dist(u, v, s) for s in SEGMENTS)
    a_grid = _clamp(PEN_HALF + 0.5 - d, 0.0, 1.0) * WHITE_A
    if a_grid > 0.0:
        r = r * (1 - a_grid) + 255 * a_grid
        g = g * (1 - a_grid) + 255 * a_grid
        b = b * (1 - a_grid) + 255 * a_grid
    return (int(r + 0.5), int(g + 0.5), int(b + 0.5), a_box)


def render(size, ss=4):
    """渲染 size×size 像素(行主序,自上而下)。ss 为超采样倍数。"""
    n = size * ss
    grid = [[0.0] * 4 for _ in range(size * size)]
    scale = 64.0 / n  # 像素坐标 → 64 单位设计空间(关键:别直接用像素当坐标)
    for j in range(n):
        v = (j + 0.5) * scale
        for i in range(n):
            u = (i + 0.5) * scale
            r, g, b, a = _pixel(u, v)
            k = (j // ss) * size + (i // ss)
            cell = grid[k]
            cell[0] += r
            cell[1] += g
            cell[2] += b
            cell[3] += a
    s2 = ss * ss
    out = []
    for cell in grid:
        out.append((int(cell[0] / s2 + 0.5), int(cell[1] / s2 + 0.5),
                    int(cell[2] / s2 + 0.5), cell[3] / s2))
    return out, size


def bmp_entry(pixels, size):
    """32bpp BMP DIB 条目(BITMAPINFOHEADER + 自下而上 BGRA + 全零 AND 掩码)。"""
    header = struct.pack("<IiiHHIIiiII", 40, size, size * 2, 1, 32, 0,
                         size * size * 4, 0, 0, 0, 0)
    rows = []
    for y in range(size - 1, -1, -1):
        row = bytearray()
        for x in range(size):
            r, g, b, a = pixels[y * size + x]
            row += bytes((b, g, r, int(a * 255 + 0.5)))
        rows.append(bytes(row))
    mask_row = b"\x00" * (((size + 31) // 32) * 4)
    return header + b"".join(rows) + mask_row * size


def png_bytes(pixels, size):
    """把像素矩阵编码成 PNG(用于 128/256 的 ICO 条目与 /icon-*.png 接口)。"""
    raw = b"".join(
        b"\x00" + bytes(v for px in pixels[y * size:(y + 1) * size]
                        for v in (px[2], px[1], px[0], int(px[3] * 255 + 0.5)))
        for y in range(size))

    def chunk(typ, data):
        return struct.pack(">I", len(data)) + typ + data + \
            struct.pack(">I", zlib.crc32(typ + data) & 0xFFFFFFFF)
    ihdr = struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b""))


def write_ico(path, sizes=(16, 24, 32, 48, 64, 128, 256)):
    """生成多尺寸 .ico(≤64 用 BMP 条目,更大用 PNG 条目)。"""
    images = []
    for s in sizes:
        px, _ = render(s, ss=4 if s <= 64 else 2)
        images.append((s, bmp_entry(px, s) if s <= 64 else png_bytes(px, s)))
    with open(path, "wb") as f:
        # ICONDIR:(保留=0, 类型=1(图标), 数量=N) —— 类型必须为 1
        f.write(struct.pack("<HHH", 0, 1, len(images)))
        offset = 6 + 16 * len(images)
        for s, data in images:
            f.write(struct.pack("<BBBBHHII", s % 256, s % 256, 0, 0, 1, 32,
                                len(data), offset))
            offset += len(data)
        for _, data in images:
            f.write(data)
    return path
