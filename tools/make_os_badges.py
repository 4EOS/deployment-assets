#!/usr/bin/env python3
"""Generate the OS/version badge icon set used on 4EOS infra diagrams.

libraries.excalidraw.com has no Windows Server edition icons, so we mint our own:
one flat 128x128 badge per OS generation, colour-graded old -> current, so a diagram
shows at a glance which servers are the old ones.

    python3 tools/make_os_badges.py [repo_root]
Writes icons/os-badges/<slug>.svg (+ .png when a converter is available) and README.md
"""
import os
import shutil
import subprocess
import sys

W, H, R = 128, 128, 22

# slug, top line, bottom line, colour, glyph
BADGES = [
    ("windows-server-2003", "2003", "WIN", "#5c1f1f", "win"),
    ("windows-server-2008", "2008", "WIN", "#8c2f2f", "win"),
    ("windows-server-2008-r2", "2008 R2", "WIN", "#a63d40", "win"),
    ("windows-server-2012", "2012", "WIN", "#d13b3f", "win"),
    ("windows-server-2012-r2", "2012 R2", "WIN", "#e5484d", "win"),
    ("windows-server-2016", "2016", "WIN", "#f5a524", "win"),
    ("windows-server-2019", "2019", "WIN", "#c2c14b", "win"),
    ("windows-server-2022", "2022", "WIN", "#30a46c", "win"),
    ("windows-server-2025", "2025", "WIN", "#1f9d8f", "win"),
    ("windows-10", "10", "DESK", "#4aa8ff", "win"),
    ("windows-11", "11", "DESK", "#6ec1ff", "win"),
    ("esxi-6.5", "6.5", "ESXi", "#a879ff", "stack"),
    ("esxi-7", "7", "ESXi", "#8f5cff", "stack"),
    ("esxi-8", "8", "ESXi", "#7442ff", "stack"),
    ("vcenter", "VC", "vCenter", "#5c7cff", "stack"),
    ("ubuntu", "LTS", "Ubuntu", "#e95420", "dots"),
    ("debian", "12", "Debian", "#d70a53", "dots"),
    ("appliance", "APP", "appliance", "#5b6470", "dots"),
    ("unknown", "?", "unknown", "#8b8b96", "dots"),
]


def glyph(kind, cx, cy, s=30):
    if kind == "win":                       # four panes
        g, gap = s, s * 0.12
        p = (g - gap) / 2
        return "".join(
            '<rect x="%.2f" y="%.2f" width="%.2f" height="%.2f" rx="1.5" fill="#fff" opacity="0.95"/>'
            % (cx - g / 2 + dx, cy - g / 2 + dy, p, p)
            for dx, dy in ((0, 0), (p + gap, 0), (0, p + gap), (p + gap, p + gap)))
    if kind == "stack":                     # three slabs (host / VM stack)
        return "".join(
            '<rect x="%.2f" y="%.2f" width="%.2f" height="%.2f" rx="2.5" fill="#fff" opacity="%.2f"/>'
            % (cx - s / 2, cy - s / 2 + i * (s * 0.36), s, s * 0.26, 0.95 - i * 0.22)
            for i in range(3))
    return "".join(                        # generic service dots
        '<circle cx="%.2f" cy="%.2f" r="%.2f" fill="#fff" opacity="%.2f"/>'
        % (cx, cy + (i - 1) * s * 0.36, s * 0.16, 0.95 - abs(i - 1) * 0.28) for i in range(3))


def svg(slug, top, bottom, colour, kind):
    return """<svg xmlns="http://www.w3.org/2000/svg" width="%d" height="%d" viewBox="0 0 %d %d">
  <rect x="2" y="2" width="%d" height="%d" rx="%d" fill="%s"/>
  <rect x="2" y="2" width="%d" height="%d" rx="%d" fill="none" stroke="#1e1e1e" stroke-opacity="0.25" stroke-width="3"/>
  %s
  <text x="%d" y="98" text-anchor="middle" font-family="Helvetica,Arial,sans-serif"
        font-size="%d" font-weight="700" fill="#ffffff">%s</text>
  <text x="%d" y="117" text-anchor="middle" font-family="Helvetica,Arial,sans-serif"
        font-size="13" font-weight="600" fill="#ffffff" fill-opacity="0.85">%s</text>
</svg>
""" % (W, H, W, H, W - 4, H - 4, R, colour, W - 4, H - 4, R, glyph(kind, W / 2, 58),
       W / 2, 24 if len(top) > 3 else 30, top, W / 2, bottom)


def converter():
    for exe, args in (("rsvg-convert", ["-w", "128", "-h", "128"]), ("magick", []), ("convert", [])):
        if shutil.which(exe):
            return exe, args
    return None, None


def main():
    root = sys.argv[1] if len(sys.argv) > 1 else "."
    out = os.path.join(root, "icons", "os-badges")
    os.makedirs(out, exist_ok=True)
    exe, cargs = converter()
    cargs = cargs or []
    made = []
    for slug, top, bottom, colour, kind in BADGES:
        p = os.path.join(out, slug + ".svg")
        open(p, "w").write(svg(slug, top, bottom, colour, kind))
        if exe:
            src = p
            cmd = ([exe] + cargs + [src, p[:-4] + ".png"]) if exe != "rsvg-convert" else \
                  ([exe] + cargs + [src, "-o", p[:-4] + ".png"])
            subprocess.run(cmd, capture_output=True)
        made.append(slug)
    with open(os.path.join(out, "README.md"), "w") as fh:
        fh.write("# OS badges\n\nOne flat badge per OS generation, colour-graded "
                 "old (dark red) -> current (green) so a diagram shows the old kit at a "
                 "glance. Regenerate with `python3 tools/make_os_badges.py .`\n\n")
        fh.write("PNG renderer detected: **%s**.\n\n" % (exe or "none - SVGs only, "
                 "install librsvg (`brew install librsvg`) for PNGs"))
        fh.write("| Badge | Means |\n|---|---|\n")
        for slug, top, bottom, colour, kind in BADGES:
            fh.write("| `%s` | %s %s |\n" % (slug, bottom, top))
        fh.write("\nColour scale: 2003/2008 `#5c1f1f`->`#a63d40`, 2012 R2 `#e5484d`, "
                 "2016 `#f5a524`, 2022 `#30a46c`. ESXi badges are purple, vCenter blue-purple, "
                 "Linux/appliance grey or brand.\n")
    print("wrote %d badges to %s (png converter: %s)" % (len(made), out, exe))


if __name__ == "__main__":
    main()
