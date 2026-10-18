#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Taqdimot generatori: "Fleshkalar (USB) qanday qilib ma'lumot saqlaydi?"

Ushbu skript tashqi kutubxonalarsiz (faqat Python standart kutubxonasi)
to'liq yaroqli .pptx (OpenXML / OOXML) faylini yaratadi.

Sabab: sandbox muhitida PyPI (python-pptx, Pillow) ga ulanish cheklangan.
.pptx fayli aslida XML qismlardan iborat ZIP arxiv, PNG esa zlib bilan
kodlanadigan oddiy format — shu bois ham taqdimot, ham unga JOYLANADIGAN
HAQIQIY RASMLAR (diagrammalar) sof Python bilan chiziladi.

Talablar:
  * Qat'iy 10 ta slayd.
  * Faqat "Times New Roman" shrifti.
  * Eng kichik shrift o'lchami 12 pt.
  * 4 qismli tuzilma (3 asosiy reja + 1 xulosa).
  * HAR BIR SLAYDDA mavzuga mos, sof Python bilan chizilgan PNG diagramma
    (ppt/media/ ichiga joylanadi) + Unsplash URL manzil izohi.
  * Matn sof, akademik o'zbek (lotin) tilida.
"""

import math
import os
import struct
import zipfile
import zlib
from datetime import datetime
from xml.sax.saxutils import escape

# ---------------------------------------------------------------------------
# EMU (English Metric Units) hisoblari. 1 pt = 12700 EMU. Slayd 16:9.
# ---------------------------------------------------------------------------
EMU_PER_PT = 12700
SLIDE_W = 12192000  # 13.333 inch (16:9)
SLIDE_H = 6858000   # 7.5 inch

# Mavzuga mos ranglar palitrasi (raqamli texnologiya / mikroelektronika)
C_BG_DARK = "0B1E3B"      # chuqur ko'k fon
C_PANEL = "12294A"        # panel ko'ki
C_ACCENT = "27E0C8"       # neon turkuaz urg'u
C_ACCENT2 = "5AA9FF"      # yorqin ko'k
C_TEXT = "EAF2FF"         # oqimtir matn
C_MUTED = "9FB8D8"        # kulrang-ko'k ikkilamchi matn
C_TITLE_BG = "061329"     # titul foni

FONT = "Times New Roman"


def rgb(hexstr):
    """'RRGGBB' -> (r, g, b)"""
    return tuple(int(hexstr[i:i + 2], 16) for i in (0, 2, 4))


def mix(c1, c2, t):
    return tuple(int(round(c1[i] + (c2[i] - c1[i]) * t)) for i in range(3))


# RGB ranglar (diagrammalar uchun)
BG = rgb(C_BG_DARK)
PANEL = rgb(C_PANEL)
ACCENT = rgb(C_ACCENT)
ACC2 = rgb(C_ACCENT2)
TEXT = rgb(C_TEXT)
MUTED = rgb(C_MUTED)
INK = rgb(C_TITLE_BG)
SILVER = (202, 212, 224)
SILVER_D = (118, 130, 148)
GREEN = (95, 220, 150)
GREEN_D = (24, 96, 60)
ORANGE = (255, 180, 90)
RED = (255, 110, 110)
PURPLE = (175, 135, 255)
YELLOW = (255, 230, 130)
CHIP = (8, 18, 36)
PIN = (150, 162, 178)
OXIDE = (70, 85, 110)
SUB = (16, 40, 70)
FG_FILL = (24, 160, 145)
TRACE = (32, 110, 122)
STREAM = (52, 72, 102)   # fonda sekin ko'rinadigan ikkilik oqim
GRID_D = (10, 24, 46)


# ===========================================================================
# 5x7 BITMAP SHRIFT (faqat ASCII: diagramma ichidagi qisqa yozuvlar uchun)
# ===========================================================================
def _g(*rows):
    assert len(rows) == 7 and all(len(r) == 5 for r in rows)
    return rows


FONT5 = {
    " ": _g(".....", ".....", ".....", ".....", ".....", ".....", "....."),
    "A": _g(".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#"),
    "B": _g("####.", "#...#", "#...#", "####.", "#...#", "#...#", "####."),
    "C": _g(".###.", "#...#", "#....", "#....", "#....", "#...#", ".###."),
    "D": _g("####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####."),
    "E": _g("#####", "#....", "#....", "####.", "#....", "#....", "#####"),
    "F": _g("#####", "#....", "#....", "####.", "#....", "#....", "#...."),
    "G": _g(".###.", "#...#", "#....", "#.###", "#...#", "#...#", ".###."),
    "H": _g("#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#"),
    "I": _g("#####", "..#..", "..#..", "..#..", "..#..", "..#..", "#####"),
    "J": _g("..###", "...#.", "...#.", "...#.", "...#.", "#..#.", ".##.."),
    "K": _g("#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#"),
    "L": _g("#....", "#....", "#....", "#....", "#....", "#....", "#####"),
    "M": _g("#...#", "##.##", "#.#.#", "#.#.#", "#...#", "#...#", "#...#"),
    "N": _g("#...#", "##..#", "##..#", "#.#.#", "#..##", "#..##", "#...#"),
    "O": _g(".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###."),
    "P": _g("####.", "#...#", "#...#", "####.", "#....", "#....", "#...."),
    "Q": _g(".###.", "#...#", "#...#", "#...#", "#.#.#", "#..#.", ".##.#"),
    "R": _g("####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#"),
    "S": _g(".####", "#....", "#....", ".###.", "....#", "....#", "####."),
    "T": _g("#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#.."),
    "U": _g("#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###."),
    "V": _g("#...#", "#...#", "#...#", "#...#", "#...#", ".#.#.", "..#.."),
    "W": _g("#...#", "#...#", "#...#", "#.#.#", "#.#.#", "##.##", "#...#"),
    "X": _g("#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#"),
    "Y": _g("#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#.."),
    "Z": _g("#####", "....#", "...#.", "..#..", ".#...", "#....", "#####"),
    "0": _g(".###.", "#...#", "#..##", "#.#.#", "##..#", "#...#", ".###."),
    "1": _g("..#..", ".##..", "..#..", "..#..", "..#..", "..#..", ".###."),
    "2": _g(".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####"),
    "3": _g("#####", "...#.", "..#..", "...#.", "....#", "#...#", ".###."),
    "4": _g("...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#."),
    "5": _g("#####", "#....", "####.", "....#", "....#", "#...#", ".###."),
    "6": _g("..##.", ".#...", "#....", "####.", "#...#", "#...#", ".###."),
    "7": _g("#####", "....#", "...#.", "..#..", ".#...", ".#...", ".#..."),
    "8": _g(".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###."),
    "9": _g(".###.", "#...#", "#...#", ".####", "....#", "...#.", ".##.."),
    "-": _g(".....", ".....", ".....", "#####", ".....", ".....", "....."),
    "+": _g(".....", "..#..", "..#..", "#####", "..#..", "..#..", "....."),
    "=": _g(".....", ".....", "#####", ".....", "#####", ".....", "....."),
    ".": _g(".....", ".....", ".....", ".....", ".....", ".##..", ".##.."),
    ":": _g(".....", ".##..", ".##..", ".....", ".##..", ".##..", "....."),
    "/": _g("....#", "....#", "...#.", "..#..", ".#...", "#....", "#...."),
    "?": _g(".###.", "#...#", "....#", "..##.", "..#..", ".....", "..#.."),
}


# ===========================================================================
# Sof Python mini-grafika: Canvas + PNG koder
# ===========================================================================
class Canvas:
    """RGB buferga chizish uchun oddiy tuval (faqat standart kutubxona)."""

    def __init__(self, w, h, bg=BG):
        self.w = w
        self.h = h
        self.px = bytearray(bytes(bg) * (w * h))

    # --- past daraja ---
    def set(self, x, y, c):
        if 0 <= x < self.w and 0 <= y < self.h:
            i = (y * self.w + x) * 3
            self.px[i] = c[0]
            self.px[i + 1] = c[1]
            self.px[i + 2] = c[2]

    def rect(self, x0, y0, x1, y1, c, fill=True, width=1):
        """[x0,x1) x [y0,y1) yarim-ochiq to'rtburchak."""
        x0, x1 = sorted((int(x0), int(x1)))
        y0, y1 = sorted((int(y0), int(y1)))
        x0 = max(0, x0); y0 = max(0, y0)
        x1 = min(self.w, x1); y1 = min(self.h, y1)
        if x0 >= x1 or y0 >= y1:
            return
        if fill:
            row = bytes(c) * (x1 - x0)
            for y in range(y0, y1):
                i = (y * self.w + x0) * 3
                self.px[i:i + len(row)] = row
        else:
            for k in range(width):
                self.line(x0, y0 + k, x1 - 1, y0 + k, c)
                self.line(x0, y1 - 1 - k, x1 - 1, y1 - 1 - k, c)
                self.line(x0 + k, y0, x0 + k, y1 - 1, c)
                self.line(x1 - 1 - k, y0, x1 - 1 - k, y1 - 1, c)

    def vgrad(self, x0, y0, x1, y1, c1, c2):
        h = max(y1 - y0, 1)
        for y in range(int(y0), int(y1)):
            self.rect(x0, y, x1, y + 1, mix(c1, c2, (y - y0) / h))

    def hgrad(self, x0, y0, x1, y1, c1, c2):
        w = max(x1 - x0, 1)
        for x in range(int(x0), int(x1)):
            self.rect(x, y0, x + 1, y1, mix(c1, c2, (x - x0) / w))

    def line(self, x0, y0, x1, y1, c, width=1):
        x0, y0, x1, y1 = int(x0), int(y0), int(x1), int(y1)
        dx, dy = x1 - x0, y1 - y0
        steps = max(abs(dx), abs(dy), 1)
        r = max(width // 2, 0)
        for i in range(steps + 1):
            t = i / steps
            x = x0 + dx * t
            y = y0 + dy * t
            for ox in range(-r, r + 1):
                for oy in range(-r, r + 1):
                    self.set(int(x) + ox, int(y) + oy, c)

    def circle(self, cx, cy, r, c, fill=True, width=1):
        cx, cy, r = int(cx), int(cy), int(r)
        for yy in range(cy - r - 1, cy + r + 2):
            for xx in range(cx - r - 1, cx + r + 2):
                d = math.hypot(xx - cx, yy - cy)
                if fill and d <= r:
                    self.set(xx, yy, c)
                elif not fill and r - width < d <= r + 0.5:
                    self.set(xx, yy, c)

    def arrow(self, x0, y0, x1, y1, c, width=2, head=13):
        self.line(x0, y0, x1, y1, c, width)
        ang = math.atan2(y1 - y0, x1 - x0)
        for da in (2.65, -2.65):
            self.line(x1, y1,
                      int(x1 + head * math.cos(ang + da)),
                      int(y1 + head * math.sin(ang + da)), c, width)

    # --- matn ---
    def text(self, x, y, s, c, scale=2):
        cx = int(x)
        for ch in s:
            g = FONT5.get(ch.upper(), FONT5[" "])
            for ry, row in enumerate(g):
                for rx, v in enumerate(row):
                    if v == "#":
                        self.rect(cx + rx * scale, y + ry * scale,
                                  cx + (rx + 1) * scale, y + (ry + 1) * scale, c)
            cx += 6 * scale
        return cx

    @staticmethod
    def textw(s, scale=2):
        return len(s) * 6 * scale

    def ctext(self, cx, y, s, c, scale=2):
        self.text(cx - self.textw(s, scale) // 2, y, s, c, scale)

    # --- PNG chiqarish ---
    def png(self):
        raw = bytearray()
        stride = self.w * 3
        for y in range(self.h):
            raw.append(0)  # filter: None
            raw += self.px[y * stride:(y + 1) * stride]

        def chunk(tag, data):
            c = tag + data
            return (struct.pack(">I", len(data)) + c
                    + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF))

        ihdr = struct.pack(">IIBBBBB", self.w, self.h, 8, 2, 0, 0, 0)
        return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
                + chunk(b"IDAT", zlib.compress(bytes(raw), 9))
                + chunk(b"IEND", b""))


# ===========================================================================
# Diagramma chizish funksiyalari
# ===========================================================================
def _drive(c, cx, cy, scale=1.0, flip=False):
    """Stilizatsiya qilingan USB fleshka (yondan ko'rinish)."""
    bw, bh = int(300 * scale), int(56 * scale)   # korpus
    cw = int(90 * scale)                          # ulagich
    x0 = cx - (bw + cw) // 2
    y0 = cy - bh // 2
    if flip:
        # ulagich chapda
        c.hgrad(x0, y0 + int(8 * scale), x0 + cw, y0 + bh - int(8 * scale),
                SILVER, SILVER_D)
        for i in range(2):
            yy = y0 + int((18 + i * 16) * scale)
            c.rect(x0 + int(10 * scale), yy, x0 + cw - int(12 * scale),
                   yy + int(8 * scale), (70, 78, 92))
        bx0 = x0 + cw
    else:
        bx0 = x0
        ux0 = x0 + bw
        c.hgrad(ux0, y0 + int(8 * scale), ux0 + cw, y0 + bh - int(8 * scale),
                SILVER_D, SILVER)
        for i in range(2):
            yy = y0 + int((18 + i * 16) * scale)
            c.rect(ux0 + int(12 * scale), yy, ux0 + cw - int(10 * scale),
                   yy + int(8 * scale), (70, 78, 92))
    # korpus
    c.vgrad(bx0, y0, bx0 + bw, y0 + bh, mix(ACC2, PANEL, 0.15), INK)
    c.rect(bx0, y0, bx0 + bw, y0 + bh, ACCENT, fill=False, width=2)
    c.rect(bx0 + int(10 * scale), y0 + int(7 * scale),
           bx0 + bw - int(10 * scale), y0 + int(12 * scale), ACCENT)
    # ichki NAND oynachasi
    c.rect(bx0 + int(150 * scale), y0 + int(18 * scale),
           bx0 + int(225 * scale), y0 + int(38 * scale), CHIP)
    c.rect(bx0 + int(150 * scale), y0 + int(18 * scale),
           bx0 + int(225 * scale), y0 + int(38 * scale), ACC2, fill=False)
    # LED
    c.circle(bx0 + int(35 * scale), cy, int(7 * scale), ACCENT)
    c.circle(bx0 + int(35 * scale), cy, int(11 * scale),
             mix(ACCENT, BG, 0.6), fill=False)


def img_banner(flip=False, w=1200, h=150, seed=0):
    """Titul/yakun slaydlari uchun dekorativ lenta."""
    c = Canvas(w, h, bg=INK)
    # ikkilik raqamlar oqimi (fon)
    for i in range(60):
        x = (i * 197 + seed * 31) % (w - 24)
        y = 8 + (i * 89 + seed * 17) % (h - 30)
        ch = "01"[(i * 7 + seed) % 2]
        c.text(x, y, ch, STREAM, scale=2)
    # markaziy nur
    c.circle(w // 2, h // 2, 46, mix(ACCENT, INK, 0.75))
    c.circle(w // 2, h // 2, 46, ACCENT, fill=False)
    _drive(c, w // 2, h // 2, scale=1.0, flip=flip)
    return c.png()


def img_anatomy(w=820, h=618):
    """Fleshka platasining anatomiyasi."""
    c = Canvas(w, h, bg=BG)
    c.rect(8, 8, w - 8, h - 8, (28, 56, 96), fill=False, width=2)
    # PCB plata
    c.rect(70, 150, 750, 540, (14, 34, 64))
    c.rect(70, 150, 750, 540, ACC2, fill=False, width=2)
    # montaj nuqtalari
    for mx, my in ((90, 170), (730, 170), (90, 520), (730, 520)):
        c.circle(mx, my, 7, SILVER_D)
    # USB ulagich
    c.hgrad(90, 230, 210, 410, SILVER, SILVER_D)
    c.rect(90, 230, 210, 410, (70, 80, 96), fill=False, width=2)
    for i in range(4):
        yy = 250 + i * 40
        c.rect(102, yy, 198, yy + 18, (70, 78, 92))
        c.rect(102, yy, 198, yy + 18, (40, 46, 58), fill=False)
    # kontroller chipi + oyoqchalari
    for i in range(6):
        yy = 252 + i * 24
        c.rect(284, yy, 300, yy + 10, PIN)
        c.rect(450, yy, 466, yy + 10, PIN)
    c.rect(300, 240, 450, 390, CHIP)
    c.rect(300, 240, 450, 390, ACCENT, fill=False, width=2)
    c.rect(330, 290, 420, 340, (12, 26, 50))
    c.rect(330, 290, 420, 340, mix(ACCENT, BG, 0.4), fill=False)
    # NAND chipi + oyoqchalari
    for i in range(8):
        xx = 545 + i * 20
        c.rect(xx, 184, xx + 10, 200, PIN)
        c.rect(xx, 440, xx + 10, 456, PIN)
    c.rect(530, 200, 710, 440, CHIP)
    c.rect(530, 200, 710, 440, ACC2, fill=False, width=2)
    c.rect(560, 260, 680, 380, (12, 26, 50))
    c.rect(560, 260, 680, 380, mix(ACC2, BG, 0.4), fill=False)
    # kvars rezonatori
    c.circle(330, 480, 22, SILVER)
    c.circle(330, 480, 22, SILVER_D, fill=False, width=2)
    c.circle(330, 480, 12, SILVER_D)
    # izlar (traces)
    for i in range(4):
        yy = 268 + i * 26
        c.line(466, yy, 530, yy, TRACE, width=2)
    for i in range(3):
        yy = 270 + i * 40
        c.line(210, yy, 300, yy, TRACE, width=2)
    c.line(352, 480, 530, 400, TRACE, width=2)
    # yozuvlar + ko'rsatkich chiziqlar
    c.ctext(150, 205, "USB", ACCENT, scale=2)
    c.ctext(375, 214, "CONTROLLER", ACCENT, scale=2)
    c.line(375, 232, 375, 240, ACCENT, width=2)
    c.ctext(620, 168, "NAND FLASH", ACC2, scale=2)
    c.line(620, 186, 620, 198, ACC2, width=2)
    c.ctext(330, 512, "CRYSTAL", MUTED, scale=2)
    c.text(700, 552, "PCB", MUTED, scale=2)
    c.ctext(410, 576, "FLESHKA PLATASI TUZILISHI", MUTED, scale=2)
    return c.png()


def img_matrix(w=820, h=618):
    """NAND xotira katakchalari matritsasi."""
    c = Canvas(w, h, bg=BG)
    c.rect(8, 8, w - 8, h - 8, (28, 56, 96), fill=False, width=2)
    x0, y0, cell, gap = 200, 120, 95, 25
    step = cell + gap
    # so'z va bit chiziqlari
    for r in range(4):
        yy = y0 + r * step + cell // 2
        c.line(90, yy, 725, yy, mix(MUTED, BG, 0.45), width=2)
        c.ctext(52, yy - 7, f"WL{r}", ACCENT, scale=2)
    for col in range(4):
        xx = x0 + col * step + cell // 2
        c.line(xx, 70, xx, 600, mix(MUTED, BG, 0.45), width=2)
        c.ctext(xx, 42, f"BL{col}", ACC2, scale=2)
    # katakchalar
    for r in range(4):
        for col in range(4):
            cx = x0 + col * step + cell // 2
            cy = y0 + r * step + cell // 2
            hl = (r == 1 and col == 2)
            c.rect(cx - 47, cy - 47, cx + 47, cy + 47,
                   mix(PANEL, ACCENT, 0.18) if hl else PANEL)
            c.rect(cx - 47, cy - 47, cx + 47, cy + 47,
                   ACCENT if hl else ACC2, fill=False, width=3 if hl else 2)
            # ichki "floating gate" belgisi
            charged = (r + col) % 3 == 0
            c.rect(cx - 26, cy - 13, cx + 26, cy + 13,
                   FG_FILL if charged else GRID_D)
            c.rect(cx - 26, cy - 13, cx + 26, cy + 13, ACCENT, fill=False)
            if charged:
                for dx, dy in ((-14, -4), (0, 4), (14, -4)):
                    c.circle(cx + dx, cy + dy, 3, INK)
    if True:
        hx = x0 + 2 * step + cell // 2
        hy = y0 + 1 * step + cell // 2
        c.circle(hx, hy, 58, mix(ACCENT, BG, 0.5), fill=False, width=2)
    c.ctext(410, 578, "WORD LINE X BIT LINE = XOTIRA KATAKCHASI", MUTED, scale=2)
    return c.png()


def img_floating_gate(w=820, h=618):
    """Suzuvchi zulfinli tranzistorning kesimi."""
    c = Canvas(w, h, bg=BG)
    c.rect(8, 8, w - 8, h - 8, (28, 56, 96), fill=False, width=2)
    # qatlamlar
    c.vgrad(120, 70, 480, 130, ACC2, (40, 80, 140))          # control gate
    c.rect(120, 70, 480, 130, mix(ACC2, TEXT, 0.3), fill=False, width=2)
    c.rect(120, 130, 480, 146, OXIDE)                        # yupqa oksid
    c.rect(160, 146, 440, 206, FG_FILL)                      # floating gate
    c.rect(160, 146, 440, 206, ACCENT, fill=False, width=2)
    c.rect(120, 206, 480, 240, OXIDE)                        # qalin oksid
    c.rect(60, 240, 560, 470, SUB)                           # substrat
    c.rect(60, 240, 560, 470, (40, 70, 110), fill=False, width=2)
    c.rect(100, 260, 230, 320, GREEN_D)                      # source
    c.rect(100, 260, 230, 320, GREEN, fill=False, width=2)
    c.rect(390, 260, 520, 320, GREEN_D)                      # drain
    c.rect(390, 260, 520, 320, GREEN, fill=False, width=2)
    # qamalgan elektronlar
    for ex, ey in ((200, 166), (260, 186), (320, 166), (380, 186), (350, 160),
                   (230, 190)):
        c.circle(ex, ey, 7, INK)
        c.line(ex - 3, ey, ex + 3, ey, ACCENT, width=2)
    c.ctext(300, 120 + 96, "E-", INK, scale=2)
    # kanal
    c.rect(230, 252, 390, 266, mix(GREEN, SUB, 0.6))
    # yozuvlar
    c.text(600, 93, "CONTROL", ACC2, scale=2)
    c.text(600, 111, "GATE", ACC2, scale=2)
    c.line(595, 100, 482, 100, ACC2, width=2)
    c.text(600, 169, "FLOATING", ACCENT, scale=2)
    c.text(600, 187, "GATE", ACCENT, scale=2)
    c.line(595, 176, 442, 176, ACCENT, width=2)
    c.text(600, 216, "OXIDE", MUTED, scale=2)
    c.line(595, 223, 482, 223, MUTED, width=2)
    c.ctext(165, 335, "SOURCE", GREEN, scale=2)
    c.ctext(455, 335, "DRAIN", GREEN, scale=2)
    c.ctext(310, 440, "SI SUBSTRATE", MUTED, scale=2)
    c.ctext(310, 500, "ELEKTRONLAR OXIDE ICHIDA QAMALADI", MUTED, scale=2)
    c.ctext(310, 578, "SUZUVCHI ZULFINLI TRANZISTOR", ACCENT, scale=2)
    return c.png()


def _mini_stack(c, cx, charged):
    """Kichik sxematik katakcha (holat ko'rsatish uchun)."""
    c.rect(cx - 60, 90, cx + 60, 118, mix(ACC2, BG, 0.35))
    c.rect(cx - 60, 90, cx + 60, 118, ACC2, fill=False)
    c.rect(cx - 60, 118, cx + 60, 130, OXIDE)
    c.rect(cx - 45, 130, cx + 45, 163, FG_FILL if charged else GRID_D)
    c.rect(cx - 45, 130, cx + 45, 163, ACCENT, fill=False)
    if charged:
        for dx, dy in ((-24, 140), (0, 150), (24, 140)):
            c.circle(cx + dx, dy, 5, INK)
    c.rect(cx - 60, 163, cx + 60, 183, OXIDE)
    c.rect(cx - 70, 183, cx + 70, 248, SUB)
    c.rect(cx - 70, 183, cx + 70, 248, (40, 70, 110), fill=False)


def _gauss(c, x0, ybase, mu, sigma, amp, color, xr):
    prev = None
    for x in range(xr[0], xr[1]):
        y = ybase - int(amp * math.exp(-((x - mu) ** 2) / (2 * sigma ** 2)))
        if prev:
            c.line(prev[0], prev[1], x, y, color, width=2)
        prev = (x, y)


def img_states(w=820, h=618):
    """Zaryad holatlari va Vth taqsimoti (SLC / MLC)."""
    c = Canvas(w, h, bg=BG)
    c.rect(8, 8, w - 8, h - 8, (28, 56, 96), fill=False, width=2)
    c.line(410, 30, 410, 560, (36, 64, 104), width=2)
    c.ctext(205, 40, "KATAKCHA HOLATLARI", ACCENT, scale=2)
    _mini_stack(c, 120, charged=False)
    _mini_stack(c, 290, charged=True)
    c.ctext(120, 272, "1", GREEN, scale=4)
    c.ctext(290, 272, "0", ORANGE, scale=4)
    c.ctext(120, 330, "ZARYAD YOQ", GREEN, scale=2)
    c.ctext(290, 330, "ZARYAD BOR", ORANGE, scale=2)
    c.ctext(205, 380, "BIT QIYMATI", MUTED, scale=2)
    # Vth taqsimotlari
    c.ctext(615, 40, "VTH TAQSIMOTI", ACC2, scale=2)
    # SLC
    c.line(460, 220, 790, 220, MUTED, width=2)
    c.line(460, 220, 460, 90, MUTED, width=2)
    _gauss(c, 460, 220, 530, 26, 110, GREEN, (460, 790))
    _gauss(c, 460, 220, 700, 26, 110, ORANGE, (460, 790))
    c.ctext(530, 230, "1", GREEN, scale=2)
    c.ctext(700, 230, "0", ORANGE, scale=2)
    c.text(760, 96, "V", MUTED, scale=2)
    c.ctext(615, 258, "SLC: 2 HOLAT = 1 BIT", MUTED, scale=2)
    # MLC
    c.line(460, 480, 790, 480, MUTED, width=2)
    c.line(460, 480, 460, 360, MUTED, width=2)
    for mu, lbl, col in ((500, "11", ACCENT), (570, "10", ACC2),
                         (640, "01", PURPLE), (710, "00", YELLOW)):
        _gauss(c, 460, 480, mu, 16, 90, col, (460, 790))
        c.ctext(mu, 490, lbl, col, scale=2)
    c.text(760, 366, "V", MUTED, scale=2)
    c.ctext(615, 528, "MLC: 4 HOLAT = 2 BIT", MUTED, scale=2)
    c.ctext(410, 578, "ZARYAD DARAJASI = BIT QIYMATI", ACCENT, scale=2)
    return c.png()


def img_wre(w=820, h=618):
    """Yozish / o'qish / o'chirish jarayonlari."""
    c = Canvas(w, h, bg=BG)
    c.rect(8, 8, w - 8, h - 8, (28, 56, 96), fill=False, width=2)
    panels = ((30, "WRITE", ACCENT), (290, "READ", ACC2), (550, "ERASE", ORANGE))
    for x0, title, col in panels:
        cx = x0 + 120
        c.rect(x0, 60, x0 + 240, 540, PANEL)
        c.rect(x0, 60, x0 + 240, 540, (40, 70, 110), fill=False, width=2)
        c.ctext(cx, 80, title, col, scale=2)
        # katakcha sxemasi
        c.rect(cx - 70, 140, cx + 70, 165, mix(ACC2, BG, 0.35))
        c.rect(cx - 70, 140, cx + 70, 165, ACC2, fill=False)
        c.rect(cx - 70, 165, cx + 70, 175, OXIDE)
        charged = title != "ERASE"
        c.rect(cx - 52, 175, cx + 52, 212, FG_FILL if charged else GRID_D)
        c.rect(cx - 52, 175, cx + 52, 212, ACCENT, fill=False)
        if charged:
            for dx in (-28, 0, 28):
                c.circle(cx + dx, 193, 5, INK)
        c.rect(cx - 70, 212, cx + 70, 232, OXIDE)
        c.rect(cx - 80, 232, cx + 80, 300, SUB)
        c.rect(cx - 80, 232, cx + 80, 300, (40, 70, 110), fill=False)
        if title == "WRITE":
            c.arrow(cx, 290, cx, 216, ACCENT, width=3)
            c.text(cx + 34, 246, "+V", ACCENT, scale=2)
            c.ctext(cx, 340, "E- KIRITILADI", MUTED, scale=2)
            c.ctext(cx, 364, "TUNNEL EFFECT", MUTED, scale=2)
        elif title == "READ":
            c.arrow(cx - 62, 266, cx + 56, 266, ACC2, width=3)
            c.circle(cx + 52, 236, 16, PANEL)
            c.circle(cx + 52, 236, 16, ACC2, fill=False, width=2)
            c.ctext(cx + 52, 229, "V", ACC2, scale=2)
            c.ctext(cx, 340, "TOK OLCHANADI", MUTED, scale=2)
            c.ctext(cx, 364, "0/1 ANIQLANADI", MUTED, scale=2)
        else:
            c.arrow(cx, 216, cx, 290, ORANGE, width=3)
            c.text(cx + 34, 246, "-V", ORANGE, scale=2)
            c.ctext(cx, 340, "E- CHIQARILADI", MUTED, scale=2)
            c.ctext(cx, 364, "BLOK BOYICHA", MUTED, scale=2)
    c.ctext(410, 578, "KUVVAT UZILSA HAM ZARYAD SAQLANADI", ACCENT, scale=2)
    return c.png()


def img_pipeline(w=820, h=618):
    """Rasm -> piksel -> bayt -> katakcha zanjiri."""
    c = Canvas(w, h, bg=BG)
    c.rect(8, 8, w - 8, h - 8, (28, 56, 96), fill=False, width=2)
    pal = [RED, ORANGE, YELLOW, GREEN, ACC2, PURPLE, ACCENT, (200, 120, 120)]
    centers = (110, 300, 490, 680)
    # 1) PHOTO — rangli mozaika
    for r in range(6):
        for col in range(6):
            colr = pal[(r * 7 + col * 13 + r * col) % len(pal)]
            c.rect(35 + col * 25, 180 + r * 25, 35 + (col + 1) * 25,
                   180 + (r + 1) * 25, colr)
    c.rect(35, 180, 185, 330, TEXT, fill=False)
    # 2) PIXELS — kattalashtirilgan piksellar
    shades = [mix(GREEN, BG, t / 24.0) for t in range(25)]
    for r in range(5):
        for col in range(5):
            c.rect(232 + col * 27, 193 + r * 27, 232 + (col + 1) * 27,
                   193 + (r + 1) * 27, shades[r * 5 + col])
            c.rect(232 + col * 27, 193 + r * 27, 232 + (col + 1) * 27,
                   193 + (r + 1) * 27, GRID_D, fill=False)
    c.rect(232, 193, 367, 328, TEXT, fill=False)
    # 3) BYTES — ikkilik satrlar
    bits = ("01101001", "11010010", "00101101", "10110100",
            "01100011", "10011100", "01011010", "11100101")
    for i, b in enumerate(bits):
        c.text(442, 180 + i * 20, b, ACCENT if i % 2 == 0 else ACC2, scale=2)
    # 4) CELLS — mini katakchalar
    for i in range(4):
        cx = 614 + i * 44
        c.rect(cx - 16, 200, cx + 16, 214, mix(ACC2, BG, 0.35))
        c.rect(cx - 16, 214, cx + 16, 220, OXIDE)
        c.rect(cx - 13, 220, cx + 13, 240, FG_FILL if i % 2 == 0 else GRID_D)
        c.rect(cx - 13, 220, cx + 13, 240, ACCENT, fill=False)
        if i % 2 == 0:
            c.circle(cx, 230, 3, INK)
        c.rect(cx - 13, 240, cx + 13, 250, OXIDE)
        c.rect(cx - 16, 250, cx + 16, 300, SUB)
    # strelkalar
    for x0, x1 in ((192, 226), (382, 416), (562, 596)):
        c.arrow(x0, 255, x1, 255, ACCENT, width=3)
    for cx, lbl in zip(centers, ("PHOTO", "PIXELS", "BYTES", "CELLS")):
        c.ctext(cx, 352, lbl, TEXT, scale=2)
    c.ctext(410, 430, "1 RASM = MILLIONLAB BIT = MILLIONLAB ZARYAD",
            MUTED, scale=2)
    c.ctext(410, 470, "64GB = 20000+ RASM SIGADI", ACCENT, scale=2)
    return c.png()


def img_3dnand(w=820, h=618):
    """3D NAND qatlamlari va sig'im o'sishi grafigi."""
    c = Canvas(w, h, bg=BG)
    c.rect(8, 8, w - 8, h - 8, (28, 56, 96), fill=False, width=2)
    c.ctext(200, 50, "3D NAND", ACCENT, scale=2)
    # qatlamlar (psevdo-3D plitalar)
    layers = []
    for i in range(4):
        y = 380 - i * 70
        layers.append(y)
        base = mix((20, 60, 110), ACC2, i * 0.12)
        for t in range(14):  # yuqori qiyasi
            c.line(60 + t * 1.4, y - t, 300 + t * 1.4, y - t,
                   mix(base, TEXT, 0.15))
        c.rect(60, y, 300, y + 26, base)
        c.rect(60, y, 300, y + 26, (40, 80, 130), fill=False)
    # vertikal ustunlar (bo'shliqlarda ko'rinadi)
    for j in range(6):
        xx = 92 + j * 36
        c.rect(xx, layers[3] - 14 - 22, xx + 10, layers[3] - 14, ACCENT)
        for i in range(3):
            y_lo = layers[i] - 14
            y_hi = layers[i + 1] + 26
            c.rect(xx, y_hi, xx + 10, y_lo, ACCENT)
        c.rect(xx, layers[0] + 26, xx + 10, layers[0] + 48, ACCENT)
    c.ctext(200, 470, "QATLAM + USTUN", MUTED, scale=2)
    # sig'im grafigi
    c.ctext(620, 50, "SIGIM OSISHI", ACC2, scale=2)
    c.line(470, 480, 790, 480, MUTED, width=2)
    c.line(470, 480, 470, 140, MUTED, width=2)
    data = (("2010", 8, MUTED), ("2015", 32, ACC2),
            ("2020", 128, ACCENT), ("2025", 512, GREEN))
    for i, (year, gb, col) in enumerate(data):
        x0 = 492 + i * 75
        hh = int(math.sqrt(gb) * 13)
        c.vgrad(x0, 480 - hh, x0 + 56, 480, mix(col, TEXT, 0.15),
                mix(col, BG, 0.45))
        c.rect(x0, 480 - hh, x0 + 56, 480, col, fill=False, width=2)
        c.ctext(x0 + 28, 480 - hh - 20, f"{gb}GB", col, scale=2)
        c.ctext(x0 + 28, 492, year, MUTED, scale=2)
    c.ctext(410, 578, "QATLAM QOSHILGANI SARI SIGIM OSHADI", ACCENT, scale=2)
    return c.png()


def build_images():
    """Barcha diagrammalarni chizib, {fayl_nomi: png_bytes} qaytaradi."""
    return {
        "image1.png": img_banner(flip=False, seed=1),          # slayd 1 lenta
        "image2.png": img_anatomy(),                           # slayd 3
        "image3.png": img_matrix(),                            # slayd 4
        "image4.png": img_floating_gate(),                     # slayd 5
        "image5.png": img_states(),                            # slayd 6
        "image6.png": img_wre(),                               # slayd 7
        "image7.png": img_pipeline(),                          # slayd 8
        "image8.png": img_3dnand(),                            # slayd 9
        "image9.png": img_banner(flip=True, w=1200, h=150, seed=7),   # slayd 10
        "image10.png": img_banner(flip=False, w=1200, h=120, seed=3),  # slayd 2
    }


# ===========================================================================
# Matnli "run" va "paragraph" yaratuvchi yordamchi funksiyalar
# ===========================================================================
def run(text, size_pt, color=C_TEXT, bold=False, italic=False):
    """Bitta matn 'run' obyektini qaytaradi (dict)."""
    return {
        "text": text,
        "size": size_pt,
        "color": color,
        "bold": bold,
        "italic": italic,
    }


def para(runs, align="l", bullet=False, space_before=600, line=110000):
    """Paragraf obyekti. `runs` - run dict'lari ro'yxati."""
    if isinstance(runs, dict):
        runs = [runs]
    return {
        "runs": runs,
        "align": align,
        "bullet": bullet,
        "space_before": space_before,
        "line": line,
    }


def run_xml(r):
    sz = int(round(r["size"])) * 100  # pt -> OOXML yuzdan birlar
    b = ' b="1"' if r["bold"] else ' b="0"'
    i = ' i="1"' if r["italic"] else ' i="0"'
    return (
        f'<a:r><a:rPr lang="uz-UZ" sz="{sz}"{b}{i} dirty="0">'
        f'<a:solidFill><a:srgbClr val="{r["color"]}"/></a:solidFill>'
        f'<a:latin typeface="{FONT}"/><a:cs typeface="{FONT}"/>'
        f'</a:rPr><a:t>{escape(r["text"])}</a:t></a:r>'
    )


def para_xml(p):
    if p["bullet"]:
        buchar = f'<a:buFont typeface="{FONT}"/><a:buChar char="&#8226;"/>'
    else:
        buchar = '<a:buNone/>'
    marL = ' marL="285750" indent="-285750"' if p["bullet"] else ''
    ppr = (
        f'<a:pPr{marL} algn="{p["align"]}">'
        f'<a:lnSpc><a:spcPct val="{p["line"]}"/></a:lnSpc>'
        f'<a:spcBef><a:spcPts val="{p["space_before"]}"/></a:spcBef>'
        f'{buchar}</a:pPr>'
    )
    runs = "".join(run_xml(r) for r in p["runs"])
    if not p["runs"]:
        runs = (
            f'<a:endParaRPr lang="uz-UZ"><a:latin typeface="{FONT}"/>'
            f'<a:cs typeface="{FONT}"/></a:endParaRPr>'
        )
    return f'<a:p>{ppr}{runs}</a:p>'


def textbox(sp_id, name, x, y, cx, cy, paragraphs, fill=None, anchor="t"):
    """Bitta matnli shakl (text box) uchun <p:sp> XML qaytaradi."""
    body = "".join(para_xml(p) for p in paragraphs)
    if fill:
        fill_xml = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
    else:
        fill_xml = '<a:noFill/>'
    return (
        f'<p:sp><p:nvSpPr>'
        f'<p:cNvPr id="{sp_id}" name="{escape(name)}"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr/></p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'{fill_xml}'
        f'</p:spPr>'
        f'<p:txBody>'
        f'<a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" '
        f'bIns="45720" anchor="{anchor}"><a:normAutofit/></a:bodyPr>'
        f'<a:lstStyle/>'
        f'{body}'
        f'</p:txBody></p:sp>'
    )


def rect(sp_id, name, x, y, cx, cy, color):
    """Dekorativ rang bo'lagi (fon panellari / urg'u chiziqlari)."""
    return (
        f'<p:sp><p:nvSpPr>'
        f'<p:cNvPr id="{sp_id}" name="{escape(name)}"/>'
        f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:ln><a:noFill/></a:ln>'
        f'</p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
    )


def pic(sp_id, name, x, y, cx, cy, rid="rId2"):
    """Joylangan rasm (<p:pic>) elementi."""
    return (
        f'<p:pic><p:nvPicPr>'
        f'<p:cNvPr id="{sp_id}" name="{escape(name)}"/>'
        f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>'
        f'<p:nvPr/></p:nvPicPr>'
        f'<p:blipFill><a:blip r:embed="{rid}"/>'
        f'<a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'</p:spPr></p:pic>'
    )


def slide_xml(shapes, bg=C_BG_DARK):
    sp_tree = "".join(shapes)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld>'
        f'<p:bg><p:bgPr><a:solidFill><a:srgbClr val="{bg}"/></a:solidFill>'
        '<a:effectLst/></p:bgPr></p:bg>'
        '<p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        f'<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f'{sp_tree}'
        '</p:spTree></p:cSld>'
        '<p:clrMapOvr><a:overrideClrMapping '
        'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" '
        'accent2="accent2" accent3="accent3" accent4="accent4" '
        'accent5="accent5" accent6="accent6" hlink="hlink" '
        'folHlink="folHlink"/></p:clrMapOvr>'
        '</p:sld>'
    )


# Umumiy joylashuv o'lchamlari (EMU)
MARGIN_X = 686000
TITLE_Y = 360000
CONTENT_Y = 1500000
CONTENT_W = SLIDE_W - 2 * MARGIN_X

# Kontent slaydlaridagi rasm joylashuvi (PNG 820x618 px, 150 dpi)
IMG_W_EMU = 820 * 6096          # 4 998 720
IMG_H_EMU = 618 * 6096          # 3 767 328
IMG_X = SLIDE_W - MARGIN_X - IMG_W_EMU   # 6 507 280
IMG_Y = 1550000
BODY_W = IMG_X - MARGIN_X - 250000       # matn ustuni eni: 5 571 280

# Lenta (banner) o'lchamlari
BANNER_W = CONTENT_W
BANNER_H = 1315000              # 1200x150 px (8:1)
BANNER2_H = 1052000             # 1200x120 px (10:1)


def header_shapes(next_id, title_text, part_label=None):
    """Har bir kontent slaydi uchun yuqori sarlavha bloki + urg'u chiziq."""
    shapes = []
    shapes.append(rect(next_id, "accentbar", MARGIN_X, TITLE_Y + 40000,
                       70000, 760000, C_ACCENT))
    next_id += 1
    paras = []
    if part_label:
        paras.append(para(run(part_label, 14, C_ACCENT, bold=True),
                          align="l", space_before=0))
    paras.append(para(run(title_text, 30, C_TEXT, bold=True),
                      align="l", space_before=200 if part_label else 0))
    shapes.append(textbox(next_id, "title", MARGIN_X + 170000, TITLE_Y,
                          CONTENT_W - 170000, 900000, paras, anchor="t"))
    next_id += 1
    return shapes, next_id


def caption_url(next_id, url):
    """Slayd pastidagi Unsplash URL manzili (ko'rinadigan izoh)."""
    paras = [
        para(run("Vizual manba (Unsplash): ", 12, C_ACCENT2, bold=True,
                 italic=True), align="l", space_before=0),
        para(run(url, 12, C_MUTED, italic=True), align="l", space_before=100),
    ]
    sp = textbox(next_id, "url_caption", MARGIN_X, SLIDE_H - 760000,
                 CONTENT_W, 640000, paras, fill=C_PANEL, anchor="ctr")
    return sp, next_id + 1


# ===========================================================================
# SLAYDLAR KONTENTI
# ===========================================================================
UNSPLASH = {
    1: "https://unsplash.com/photos/a-close-up-of-a-usb-flash-drive-on-a-table-9SoCnyQmkzI",
    2: "https://unsplash.com/photos/black-and-silver-usb-flash-drive-ZV_64LdGoao",
    3: "https://unsplash.com/photos/macro-shot-of-a-circuit-board-FO7JIlwjOtU",
    4: "https://unsplash.com/photos/green-computer-circuit-board-wtpVH-C58FA",
    5: "https://unsplash.com/photos/microchip-close-up-electronics-ZO02-30bFfs",
    6: "https://unsplash.com/photos/silicon-wafer-semiconductor-5fNmWej4tAA",
    7: "https://unsplash.com/photos/binary-code-digital-data-stream-xkBaqlcqeb4",
    8: "https://unsplash.com/photos/colorful-photo-gallery-thumbnails-on-screen-2JEWnQ6y1Ac",
    9: "https://unsplash.com/photos/modern-data-storage-technology-OqtafYT5kTw",
    10: "https://unsplash.com/photos/thank-you-note-on-digital-display-5QgIuuBxKwM",
}

# Slayd -> joylanadigan rasm fayli
SLIDE_IMG = {
    1: "image1.png",
    2: "image10.png",
    3: "image2.png",
    4: "image3.png",
    5: "image4.png",
    6: "image5.png",
    7: "image6.png",
    8: "image7.png",
    9: "image8.png",
    10: "image9.png",
}


def _content_slide(idx, title, part, bullets):
    """Rasm + matnli standart kontent slaydini quradi."""
    sid = 2
    shapes = []
    hdr, sid = header_shapes(sid, title, part)
    shapes += hdr
    body = [
        para([run(lead + " ", 17, C_ACCENT if idx != 9 else C_ACCENT2, bold=True),
              run(rest, 15, C_TEXT)], bullet=True)
        for lead, rest in bullets
    ]
    shapes.append(textbox(sid, "body", MARGIN_X, CONTENT_Y, BODY_W, 3900000,
                          body, anchor="t"))
    sid += 1
    shapes.append(pic(sid, "diagram", IMG_X, IMG_Y, IMG_W_EMU, IMG_H_EMU))
    sid += 1
    cap, sid = caption_url(sid, UNSPLASH[idx])
    shapes.append(cap)
    return slide_xml(shapes), SLIDE_IMG[idx]


def build_slides():
    """Har bir slayd uchun (slide_xml, rasm_fayli|None) juftligini qaytaradi."""
    slides = []

    # ------------------------------------------------------------------ 1
    # Titul varag'i
    sid = 2
    shapes = []
    shapes.append(rect(sid, "panel", 0, 0, SLIDE_W, SLIDE_H, C_TITLE_BG)); sid += 1
    shapes.append(rect(sid, "topbar", 0, 0, SLIDE_W, 90000, C_ACCENT)); sid += 1
    shapes.append(rect(sid, "botbar", 0, SLIDE_H - 90000, SLIDE_W, 90000, C_ACCENT2)); sid += 1

    inst = [
        para(run("Jizzax davlat pedagogika universiteti", 20, C_ACCENT2, bold=True),
             align="ctr", space_before=0),
        para(run("Axborot tizimlari fani", 16, C_MUTED, italic=True),
             align="ctr", space_before=200),
    ]
    shapes.append(textbox(sid, "univ", MARGIN_X, 520000, CONTENT_W, 900000,
                          inst, anchor="t")); sid += 1

    shapes.append(pic(sid, "banner", MARGIN_X, 1500000, BANNER_W, BANNER_H))
    sid += 1

    title = [
        para(run("Fleshkalar (USB) qanday qilib ma'lumot saqlaydi?", 36,
                 C_TEXT, bold=True), align="ctr", space_before=0),
        para(run("Harakatlanuvchi qismlarsiz minglab rasmlarning elektr "
                 "zaryadlari ko'rinishida saqlanishi", 18, C_ACCENT,
                 italic=True), align="ctr", space_before=400, line=120000),
    ]
    shapes.append(textbox(sid, "maintitle", MARGIN_X, 2950000, CONTENT_W,
                          1600000, title, anchor="ctr")); sid += 1

    presenter = [
        para(run("Taqdimotchi: Toshboʻriyeva Marjona", 18, C_TEXT, bold=True),
             align="ctr", space_before=0),
        para(run("242_26 - guruh", 16, C_MUTED), align="ctr", space_before=150),
    ]
    shapes.append(textbox(sid, "presenter", MARGIN_X, 4680000, CONTENT_W,
                          800000, presenter, anchor="ctr")); sid += 1

    cap, sid = caption_url(sid, UNSPLASH[1])
    shapes.append(cap)
    slides.append((slide_xml(shapes, bg=C_TITLE_BG), SLIDE_IMG[1]))

    # ------------------------------------------------------------------ 2
    # Reja
    sid = 2
    shapes = []
    hdr, sid = header_shapes(sid, "Taqdimot rejasi")
    shapes += hdr
    plan = [
        para([run("1-qism.  ", 20, C_ACCENT, bold=True),
              run("Fleshkalarning ishlash prinsipi va NAND xotira",
                  20, C_TEXT)], bullet=False, space_before=500),
        para([run("2-qism.  ", 20, C_ACCENT, bold=True),
              run("Ma'lumotlarni elektr zaryadi ko'rinishida saqlash",
                  20, C_TEXT)], bullet=False, space_before=500),
        para([run("3-qism.  ", 20, C_ACCENT, bold=True),
              run("Yozish, o'qish va o'chirish jarayonlari",
                  20, C_TEXT)], bullet=False, space_before=500),
        para([run("Xulosa.  ", 20, C_ACCENT2, bold=True),
              run("Fleshkalarning afzalliklari va texnologiya kelajagi",
                  20, C_TEXT)], bullet=False, space_before=500),
    ]
    shapes.append(textbox(sid, "plan", MARGIN_X, CONTENT_Y, CONTENT_W,
                          3200000, plan, anchor="t")); sid += 1
    shapes.append(pic(sid, "banner2", MARGIN_X, 4860000, BANNER_W, BANNER2_H))
    sid += 1
    cap, sid = caption_url(sid, UNSPLASH[2])
    shapes.append(cap)
    slides.append((slide_xml(shapes), SLIDE_IMG[2]))

    # ------------------------------------------------------------------ 3
    slides.append(_content_slide(
        3, "Fleshkaning ishlash prinsipi", "1-QISM — NAND XOTIRA",
        [
            ("Qattiq jism xotirasi (Solid-State).",
             "Fleshkalar ma'lumotni butunlay elektron tarzda, yarimo'tkazgichli "
             "mikrosxemalarda saqlaydi."),
            ("Harakatlanuvchi qismlar yo'q.",
             "Qattiq disklardan farqli ravishda fleshkalarda aylanuvchi plastinka "
             "yoki o'qiydigan boshcha mavjud emas."),
            ("NAND tuzilmasi.",
             "Xotira katakchalari NAND mantiqiy sxemasi asosida zich qatorlarga "
             "birlashtirilib, kichik hajmda katta sig'imni ta'minlaydi."),
            ("Energiyasiz saqlash.",
             "Ma'lumot elektr quvvati uzilgandan keyin ham saqlanib qoladi — bu "
             "energiyaga bog'liq bo'lmagan (non-volatile) xotira."),
        ]))

    # ------------------------------------------------------------------ 4
    slides.append(_content_slide(
        4, "NAND xotira katakchalari tuzilishi", "1-QISM — NAND XOTIRA",
        [
            ("Xotira katakchasi (cell).",
             "Ma'lumotning eng kichik birligi maxsus tranzistorda saqlanadi; "
             "har bir katakcha bir yoki bir nechta bitni ifodalaydi."),
            ("Qatorlar va ustunlar.",
             "Katakchalar o'zaro ulanib, so'z chiziqlari (word line) va bit "
             "chiziqlari (bit line) to'rini hosil qiladi."),
            ("Zichlik va sig'im.",
             "3D NAND texnologiyasida katakchalar vertikal qatlamlab joylashtirilib, "
             "sig'im sezilarli oshiriladi."),
            ("Ishonchlilik.",
             "Xatolarni tuzatuvchi kodlar (ECC) va kiyilishni taqsimlash (wear "
             "leveling) ma'lumot butunligini saqlaydi."),
        ]))

    # ------------------------------------------------------------------ 5
    slides.append(_content_slide(
        5, "Suzuvchi zulfin (Floating Gate) tranzistori",
        "2-QISM — ZARYAD KO'RINISHIDA SAQLASH",
        [
            ("Suzuvchi zulfin.",
             "Har bir katakchada ikki qavatli zulfinli tranzistor bo'lib, ulardan "
             "biri izolyator bilan to'liq o'ralgan suzuvchi zulfindir."),
            ("Elektronlarni qamab qo'yish.",
             "Elektronlar suzuvchi zulfinga kiritilib, oksid qatlam ularning "
             "chiqib ketishiga yo'l qo'ymaydi — zaryad shu yerda qamalib qoladi."),
            ("Doimiy saqlash fizikasi.",
             "Izolyatsiya tufayli zaryad quvvat uzilsa ham yillar davomida saqlanadi, "
             "bu esa ma'lumotning doimiyligini ta'minlaydi."),
        ]))

    # ------------------------------------------------------------------ 6
    slides.append(_content_slide(
        6, "Zaryad qanday qilib ma'lumotga aylanadi",
        "2-QISM — ZARYAD KO'RINISHIDA SAQLASH",
        [
            ("Zaryad miqdori — holat.",
             "Suzuvchi zulfindagi elektronlar mavjudligi yoki yo'qligi katakchaning "
             "mantiqiy holatini belgilaydi."),
            ("Chegaraviy kuchlanish.",
             "Qamalgan zaryad tranzistorning ochilish kuchlanishini o'zgartiradi; "
             "aynan shu farq o'qish paytida aniqlanadi."),
            ("Ko'p bitli katakchalar.",
             "MLC, TLC va QLC texnologiyalarida bitta katakcha zaryadning bir necha "
             "darajasini saqlab, 2, 3 yoki 4 bitni ifodalaydi."),
            ("Barqarorlik.",
             "Zaryad darajalari aniq o'lchanadigan oraliqlarga bo'linib, ma'lumot "
             "ishonchli tiklanadi."),
        ]))

    # ------------------------------------------------------------------ 7
    slides.append(_content_slide(
        7, "Yozish, o'qish va o'chirish jarayonlari", "3-QISM — JARAYONLAR",
        [
            ("Yozish.",
             "Yuqori kuchlanish ta'sirida elektronlar oksid qatlamdan o'tib "
             "(tunnel effekti) suzuvchi zulfinga joylashtiriladi."),
            ("O'qish.",
             "Katakchaga kichik kuchlanish berilib, tok o'tishi tekshiriladi va "
             "zaryad holati 0 yoki 1 sifatida aniqlanadi."),
            ("O'chirish.",
             "Teskari kuchlanish zaryadni suzuvchi zulfindan chiqarib yuboradi; "
             "o'chirish bloklar bo'yicha amalga oshiriladi."),
            ("Ikkilik sanoq sistemasi.",
             "Zaryad bor holat — 0, zaryad yo'q holat — 1 (yoki aksincha) sifatida "
             "talqin qilinadi."),
        ]))

    # ------------------------------------------------------------------ 8
    slides.append(_content_slide(
        8, "Minglab rasm elektron axborot sifatida", "3-QISM — JARAYONLAR",
        [
            ("Rasm — baytlar to'plami.",
             "Har bir rasm pikesllardan, piksellar esa 0 va 1 lardan iborat "
             "baytlar ketma-ketligi sifatida kodlanadi."),
            ("Baytlardan katakchalarga.",
             "Bu bitlar millionlab xotira katakchalaridagi zaryad holatlari orqali "
             "fizik tarzda saqlanadi."),
            ("Katta sig'im.",
             "Mikrosxemadagi katakchalar soni milliardlab bo'lgani uchun bitta "
             "fleshka minglab yuqori sifatli rasmni saqlay oladi."),
            ("Manzillash.",
             "Boshqaruvchi (controller) har bir faylni aniq manzillarga joylashtirib, "
             "keyinchalik xatosiz tiklaydi."),
        ]))

    # ------------------------------------------------------------------ 9
    slides.append(_content_slide(
        9, "Xulosa", "4-QISM — XULOSA",
        [
            ("Afzalliklari.",
             "Fleshkalar ixcham, yengil, tez va energiya tejamkor bo'lib, kundalik "
             "ma'lumot almashinuvini osonlashtiradi."),
            ("Chidamliligi.",
             "Harakatlanuvchi qismlar yo'qligi tufayli ular zarba va tebranishga "
             "chidamli, mexanik buzilishlarga kam uchraydi."),
            ("Texnologiya kelajagi.",
             "3D NAND va yangi yarimo'tkazgich materiallari sig'imni oshirib, narxni "
             "pasaytirmoqda — kelajakda yanada tezroq xotira kutilmoqda."),
            ("Umumiy xulosa.",
             "Fleshka — bu harakatsiz, lekin elektr zaryadlari asosida ishlaydigan "
             "ishonchli raqamli xotira yechimidir."),
        ]))

    # ------------------------------------------------------------------ 10
    sid = 2
    shapes = []
    shapes.append(rect(sid, "panel", 0, 0, SLIDE_W, SLIDE_H, C_TITLE_BG)); sid += 1
    shapes.append(rect(sid, "topbar", 0, 0, SLIDE_W, 90000, C_ACCENT)); sid += 1
    shapes.append(rect(sid, "botbar", 0, SLIDE_H - 90000, SLIDE_W, 90000, C_ACCENT2)); sid += 1
    shapes.append(pic(sid, "banner3", MARGIN_X, 880000, BANNER_W, BANNER_H))
    sid += 1
    body = [
        para(run("E'tiboringiz uchun rahmat!", 40, C_ACCENT, bold=True),
             align="ctr", space_before=0),
        para(run("Savollaringiz uchun ochiq maydon", 22, C_TEXT, italic=True),
             align="ctr", space_before=600),
        para(run("Toshboʻriyeva Marjona — 242_26-guruh", 16, C_MUTED),
             align="ctr", space_before=800),
    ]
    shapes.append(textbox(sid, "thanks", MARGIN_X, 2500000, CONTENT_W, 2300000,
                          body, anchor="ctr")); sid += 1
    cap, sid = caption_url(sid, UNSPLASH[10])
    shapes.append(cap)
    slides.append((slide_xml(shapes, bg=C_TITLE_BG), SLIDE_IMG[10]))

    return slides


# ===========================================================================
# OOXML paket qismlari (statik)
# ===========================================================================
def content_types(n):
    overrides = "".join(
        f'<Override PartName="/ppt/slides/slide{i}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, n + 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Override PartName="/ppt/presentation.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
        '<Override PartName="/ppt/theme/theme1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        f'{overrides}'
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        '</Types>'
    )


ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
    '</Relationships>'
)


def presentation_xml(n):
    sldids = "".join(
        f'<p:sldId id="{255 + i}" r:id="rId{i + 1}"/>' for i in range(1, n + 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        f'<p:sldIdLst>{sldids}</p:sldIdLst>'
        f'<p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}" type="screen16x9"/>'
        f'<p:notesSz cx="{SLIDE_H}" cy="{SLIDE_W}"/>'
        '</p:presentation>'
    )


def presentation_rels(n):
    rels = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>']
    for i in range(1, n + 1):
        rels.append(
            f'<Relationship Id="rId{i + 1}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
            f'Target="slides/slide{i}.xml"/>'
        )
    rels.append(
        f'<Relationship Id="rId{n + 2}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
        f'Target="theme/theme1.xml"/>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(rels) + '</Relationships>'
    )


THEME = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">'
    '<a:themeElements>'
    '<a:clrScheme name="Office">'
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
    '<a:dk2><a:srgbClr val="0B1E3B"/></a:dk2>'
    '<a:lt2><a:srgbClr val="EAF2FF"/></a:lt2>'
    '<a:accent1><a:srgbClr val="27E0C8"/></a:accent1>'
    '<a:accent2><a:srgbClr val="5AA9FF"/></a:accent2>'
    '<a:accent3><a:srgbClr val="12294A"/></a:accent3>'
    '<a:accent4><a:srgbClr val="9FB8D8"/></a:accent4>'
    '<a:accent5><a:srgbClr val="061329"/></a:accent5>'
    '<a:accent6><a:srgbClr val="2E86DE"/></a:accent6>'
    '<a:hlink><a:srgbClr val="5AA9FF"/></a:hlink>'
    '<a:folHlink><a:srgbClr val="9FB8D8"/></a:folHlink>'
    '</a:clrScheme>'
    '<a:fontScheme name="Office">'
    '<a:majorFont><a:latin typeface="Times New Roman"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
    '<a:minorFont><a:latin typeface="Times New Roman"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
    '</a:fontScheme>'
    '<a:fmtScheme name="Office">'
    '<a:fillStyleLst>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '</a:fillStyleLst>'
    '<a:lnStyleLst>'
    '<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
    '<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
    '<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
    '</a:lnStyleLst>'
    '<a:effectStyleLst>'
    '<a:effectStyle><a:effectLst/></a:effectStyle>'
    '<a:effectStyle><a:effectLst/></a:effectStyle>'
    '<a:effectStyle><a:effectLst/></a:effectStyle>'
    '</a:effectStyleLst>'
    '<a:bgFillStyleLst>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    '</a:bgFillStyleLst>'
    '</a:fmtScheme>'
    '</a:themeElements>'
    '</a:theme>'
)


SLIDE_MASTER = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="0B1E3B"/></a:solidFill>'
    '<a:effectLst/></p:bgPr></p:bg>'
    '<p:spTree>'
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    '</p:spTree></p:cSld>'
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" '
    'accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" '
    'accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
    '<p:txStyles>'
    '<p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"><a:latin typeface="Times New Roman"/></a:defRPr></a:lvl1pPr></p:titleStyle>'
    '<p:bodyStyle><a:lvl1pPr><a:defRPr sz="1800"><a:latin typeface="Times New Roman"/></a:defRPr></a:lvl1pPr></p:bodyStyle>'
    '<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"><a:latin typeface="Times New Roman"/></a:defRPr></a:lvl1pPr></p:otherStyle>'
    '</p:txStyles>'
    '</p:sldMaster>'
)

SLIDE_MASTER_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
    '</Relationships>'
)

SLIDE_LAYOUT = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    'type="blank" preserve="1">'
    '<p:cSld name="Blank">'
    '<p:spTree>'
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    '</p:spTree></p:cSld>'
    '<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" '
    'accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" '
    'accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>'
    '</p:sldLayout>'
)

SLIDE_LAYOUT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
    '</Relationships>'
)


def slide_rels(image_target=None):
    """Slayd munosabatlari: layout + (ixtiyoriy) joylangan rasm."""
    rels = [
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
        'Target="../slideLayouts/slideLayout1.xml"/>'
    ]
    if image_target:
        rels.append(
            '<Relationship Id="rId2" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            f'Target="../media/{image_target}"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(rels) + '</Relationships>'
    )


def core_xml():
    now = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<cp:coreProperties '
        'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>Fleshkalar (USB) qanday qilib ma&#39;lumot saqlaydi?</dc:title>'
        '<dc:creator>Toshbo\u02bbriyeva Marjona (242_26-guruh)</dc:creator>'
        '<cp:lastModifiedBy>Toshbo\u02bbriyeva Marjona</cp:lastModifiedBy>'
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
        '</cp:coreProperties>'
    )


def app_xml(n):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Application>Python OOXML Generator</Application>'
        f'<Slides>{n}</Slides>'
        '<Company>Jizzax davlat pedagogika universiteti</Company>'
        '</Properties>'
    )


# ===========================================================================
# Paketlash
# ===========================================================================
def build(output_path):
    slides = build_slides()
    n = len(slides)
    assert n == 10, f"Slaydlar soni 10 bo'lishi shart, hozir: {n}"
    images = build_images()

    parts = {
        "[Content_Types].xml": content_types(n),
        "_rels/.rels": ROOT_RELS,
        "ppt/presentation.xml": presentation_xml(n),
        "ppt/_rels/presentation.xml.rels": presentation_rels(n),
        "ppt/theme/theme1.xml": THEME,
        "ppt/slideMasters/slideMaster1.xml": SLIDE_MASTER,
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": SLIDE_MASTER_RELS,
        "ppt/slideLayouts/slideLayout1.xml": SLIDE_LAYOUT,
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels": SLIDE_LAYOUT_RELS,
        "docProps/core.xml": core_xml(),
        "docProps/app.xml": app_xml(n),
    }
    for i, (xml, img) in enumerate(slides, start=1):
        parts[f"ppt/slides/slide{i}.xml"] = xml
        parts[f"ppt/slides/_rels/slide{i}.xml.rels"] = slide_rels(img)
    for name, data in images.items():
        parts[f"ppt/media/{name}"] = data

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z:
        # [Content_Types].xml birinchi bo'lib yozilishi tavsiya etiladi
        z.writestr("[Content_Types].xml", parts.pop("[Content_Types].xml"))
        for name, data in parts.items():
            z.writestr(name, data)
    return output_path, n, len(images)


if __name__ == "__main__":
    out = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "Fleshkalar_USB_malumot_saqlashi.pptx",
    )
    path, n, imgs = build(out)
    print(f"Yaratildi: {path}")
    print(f"Slaydlar soni: {n}")
    print(f"Joylangan rasmlar soni: {imgs}")
