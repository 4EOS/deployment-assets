#!/usr/bin/env python3
"""Download the Excalidraw icon libraries 4EOS wants for infra diagrams.

Writes them to icons/excalidraw-libraries/<author>__<name>.excalidrawlib and emits
icons/excalidraw-libraries/INDEX.md describing every library and the items inside it,
so a diagram author can pick an icon without opening 20 files.

    python3 tools/fetch_excalidraw_libraries.py [repo_root]
"""
import json
import os
import sys
import urllib.request

SOURCES = [
    # network / infrastructure topology
    "dwelle/network-topology-icons.excalidrawlib",
    "mateuszbaransanok/it-icons.excalidrawlib",
    "jgodoy/racks-and-servers-components.excalidrawlib",
    "jgodoy/network-locations.excalidrawlib",
    "dday987/common-home-network-basics.excalidrawlib",
    "anna-pastushko/architecture-diagram-components.excalidrawlib",
    "arach/systems-design-components.excalidrawlib",
    # virtualisation / platform
    "odraghi/vmware-architecture-design.excalidrawlib",
    "novakkkarel/nsx-t-vmware.excalidrawlib",
    "lowess/kubernetes-icons-set.excalidrawlib",
    "boemska-nik/kubernetes-icons.excalidrawlib",
    "markopolo123/dev_ops.excalidrawlib",
    # vendors we run
    "fortijosh/fortinet.excalidrawlib",
    "drwnio/drwnio.excalidrawlib",
    "maeddes/technology-logos.excalidrawlib",
    "zesty-lemur/microsoft-apps.excalidrawlib",
    "wictorwilen/microsoft-365-icons.excalidrawlib",
    "rockssk/microsoft-azure-cloud-icons.excalidrawlib",
]

URL = "https://libraries.excalidraw.com/libraries/%s"


def item_names(item):
    """Best-effort human label for one library item."""
    names, texts = [], []
    for el in item.get("elements", []):
        if el.get("type") == "text" and (el.get("text") or "").strip():
            texts.append(el["text"].strip().replace("\n", " "))
        if el.get("type") == "image" and el.get("fileId"):
            f = (item.get("files") or {}).get(el["fileId"]) or {}
            if f.get("created"):
                names.append(f["created"].rsplit("/", 1)[-1])
    if item.get("name"):
        names.insert(0, str(item["name"]))
    return (names[:1] or texts[:2] or ["(no label)"])[:2]


def main():
    root = sys.argv[1] if len(sys.argv) > 1 else "."
    outdir = os.path.join(root, "icons", "excalidraw-libraries")
    os.makedirs(outdir, exist_ok=True)
    rows, total_items, failures = [], 0, []
    for src in SOURCES:
        author, fname = src.split("/")
        local = "%s__%s" % (author, fname)
        path = os.path.join(outdir, local)
        try:
            with urllib.request.urlopen(URL % src, timeout=120) as r:
                raw = r.read()
            lib = json.loads(raw)
        except Exception as e:                                     # noqa: BLE001
            failures.append((src, repr(e)[:120]))
            continue
        items = lib.get("libraryItems", [])
        total_items += len(items)
        with open(path, "wb") as fh:
            fh.write(raw)
        samples = []
        for it in items[:4]:
            samples += item_names(it)
        rows.append({"source": src, "file": "icons/excalidraw-libraries/" + local,
                     "items": len(items), "kb": round(len(raw) / 1024, 1),
                     "sample": ", ".join(dict.fromkeys(samples))[:110]})

    rows.sort(key=lambda r: r["source"])
    with open(os.path.join(outdir, "INDEX.md"), "w") as fh:
        fh.write("# Excalidraw icon libraries\n\n")
        fh.write("Vendored from <https://libraries.excalidraw.com> (MIT). "
                 "Refresh with `python3 tools/fetch_excalidraw_libraries.py`.\n\n")
        fh.write("Load in Excalidraw: *Library panel -> menu -> Open library* and pick the "
                 "`.excalidrawlib` file (or drag it onto the canvas).\n\n")
        fh.write("| Library | Items | KB | Sample items | File |\n|---|---|---|---|---|\n")
        for r in rows:
            fh.write("| `%s` | %d | %s | %s | [%s](%s) |\n" % (
                r["source"], r["items"], r["kb"], r["sample"],
                r["file"].split("/")[-1], r["file"].split("/")[-1]))
        fh.write("\n## What lives where (pick by job)\n\n")
        fh.write("- **Network topology / sites** — `dwelle/network-topology-icons`, "
                 "`jgodoy/network-locations`, `dday987/common-home-network-basics`\n")
        fh.write("- **Servers, racks, blades** — `jgodoy/racks-and-servers-components`, "
                 "`mateuszbaransanok/it-icons`\n")
        fh.write("- **Virtualisation** — `odraghi/vmware-architecture-design`, "
                 "`novakkkarel/nsx-t-vmware`\n")
        fh.write("- **Firewalls we run** — `fortijosh/fortinet`\n")
        fh.write("- **Vendor / product logos** — `maeddes/technology-logos`, "
                 "`drwnio/drwnio`, `zesty-lemur/microsoft-apps`\n")
        fh.write("- **Local 4EOS PNG logos** (NinjaOne, Huntress, NetBird, Auvik, Proofpoint, "
                 "Fortinet, ScalePad) live one level up in `icons/`\n")
    print("downloaded %d/%d libraries, %d items, index at %s/INDEX.md" %
          (len(rows), len(SOURCES), total_items, outdir))
    for f in failures:
        print("  FAILED:", f)


if __name__ == "__main__":
    main()
