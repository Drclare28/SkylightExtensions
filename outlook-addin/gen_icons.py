#!/usr/bin/env python3
"""Generate the add-in icon PNGs using only the Python standard library.

Outputs assets/icon-{16,32,64,80}.png - a rounded square with a cyan gradient
and a simple white paper-plane glyph (a nod to "Skylight").
"""
import os
import struct
import zlib


def clamp(v, lo=0, hi=255):
    return max(lo, min(hi, int(round(v))))


def write_png(path, size, rgba):
    """Write a size x size RGBA image as a PNG file."""
    raw = b""
    for y in range(size):
        raw += b"\x00"  # filter: none
        for x in range(size):
            raw += bytes(rgba[y * size + x])
    def chunk(tag, data):
        c = struct.pack(">I", len(data)) + tag + data
        c += struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        return c
    png = b"\x89PNG\r\n\x1a\n"
    png += chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0))
    png += chunk(b"IDAT", zlib.compress(raw, 9))
    png += chunk(b"IEND", b"")
    with open(path, "wb") as f:
        f.write(png)


def rounded_rect(x, y, size, radius):
    """True when (x,y) is inside a size x size rounded square at origin."""
    r = radius
    x0, x1 = r, size - r
    y0, y1 = r, size - r
    if x < x0:
        if y < y0:
            return (x - x0) ** 2 + (y - y0) ** 2 <= r * r
        if y > y1:
            return (x - x0) ** 2 + (y - y1) ** 2 <= r * r
    if x > x1:
        if y < y0:
            return (x - x1) ** 2 + (y - y0) ** 2 <= r * r
        if y > y1:
            return (x - x1) ** 2 + (y - y1) ** 2 <= r * r
    return True


def plane_px(x, y, size):
    """True when (x,y) inside the paper-plane glyph, in unit space."""
    # Paper plane: three triangles. Geometry in normalized [-1,1] space.
    u = (x + 0.5) / size * 2 - 1
    v = (y + 0.5) / size * 2 - 1
    # Main body upward triangle (nose at top-right).
    a = (0.85, -0.35)   # nose
    b = (-0.8, 0.6)     # left wing tip
    c = (0.2, 0.75)     # right tail base
    if in_tri(u, v, a, b, c):
        return True
    # Nose-line fold back to center.
    d = (-0.05, -0.05)
    if in_tri(u, v, a, b, d):
        return True
    # Bottom fold triangle.
    e = (0.35, 0.55)
    f = (0.05, -0.05)
    if in_tri(u, v, a, e, f):
        return True
    return False


def in_tri(px, py, a, b, c):
    def sign(p1, p2, p3):
        return (p1[0] - p3[0]) * (p2[1] - p3[1]) - (p2[0] - p3[0]) * (p1[1] - p3[1])
    d1 = sign((px, py), a, b)
    d2 = sign((px, py), b, c)
    d3 = sign((px, py), c, a)
    has_neg = (d1 < 0) or (d2 < 0) or (d3 < 0)
    has_pos = (d1 > 0) or (d2 > 0) or (d3 > 0)
    return not (has_neg and has_pos)


def make_icon(size):
    margin = max(1, size // 16)
    radius = size // 5
    px = []
    for y in range(size):
        for x in range(size):
            if rounded_rect(x, y, size, radius):
                # vertical cyan-blue gradient
                t = y / size
                r = clamp(18 + 40 * t)
                g = clamp(108 + 80 * t)
                b = clamp(230 + 20 * t)
                if plane_px(x - margin, y, size - 2 * margin):
                    r, g, b = 255, 255, 255
                # anti-alias edge: lighten boundary pixels
                if not rounded_rect(x, y, size, radius + 1):
                    r, g, b = clamp(r + 60), clamp(g + 60), clamp(b + 60)
                px.append((r, g, b, 255))
            else:
                px.append((0, 0, 0, 0))
    return px


def main():
    base = os.path.dirname(os.path.abspath(__file__))
    assets = os.path.join(base, "assets")
    os.makedirs(assets, exist_ok=True)
    for size in (16, 32, 64, 80):
        path = os.path.join(assets, "icon-%d.png" % size)
        write_png(path, size, make_icon(size))
        print("wrote %s" % path)


if __name__ == "__main__":
    main()