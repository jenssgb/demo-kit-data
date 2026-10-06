"""Builds profiles/<profile>/ – a copy of every demo folder with the demo people swapped for real tenant users.

    python tools/build_profile.py cdx          (reads profiles/cdx.json, written by Update-Bundle.ps1)
    .\\tools\\Update-Manifest.ps1 -All -Root profiles\\cdx

Same rules as lib/personas.ts in the Demo Kit, so deck, mails and files always match:
  1. e-mail address (case-insensitive)  2. full name (also "Megan-Bowen-Mail")
  3. first name as a whole word (+ genitive "s"), unless another surname follows ("Anne Patel" stays)
File names stay unchanged ("_" is a word character, so 1_Megan_Carrier_Update.eml keeps its name and
every reference to it stays valid). Fails if a default name is left over in a text file or Office document
(e.g. a name split across Word runs); binary files (.onepkg, .pdf, images) are copied unchanged and listed.
"""

import json
import quopri
import re
import shutil
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TEXT = {".eml", ".json", ".csv", ".txt", ".vtt", ".md", ".html"}
OFFICE = {".docx", ".xlsx", ".pptx"}


def load_rules(profile: dict):
    pairs = [(f, profile["people"][pid]) for pid, f in profile["personas"].items() if pid in profile["people"]]
    first = lambda n: n.split(" ")[0]
    rest = lambda n: " ".join(n.split(" ")[1:])
    rules = []
    rules += [(re.compile(re.escape(f["email"]), re.I), t["email"]) for f, t in pairs if f.get("email") and t.get("email")]
    # full name, also as German compound ("Megan-Bowen-Mail")
    rules += [
        (re.compile(rf"\b{re.escape(first(f['name']))}([ -]){re.escape(rest(f['name']))}\b"), first(t["name"]) + r"\1" + rest(t["name"]))
        for f, t in pairs
        if f["name"] != t["name"] and rest(f["name"]) and rest(t["name"])
    ]
    rules += [(re.compile(rf"\b{re.escape(f['name'])}\b"), t["name"]) for f, t in pairs if f["name"] != t["name"]]
    # first name alone (+ genitive s), but not when another surname follows ("Anne Patel" is someone else)
    other = r"(?![ -][A-Z][a-z]+\b)"
    rules += [
        (re.compile(rf"\b{re.escape(first(f['name']))}(s?)\b{other}"), first(t["name"]) + r"\1")
        for f, t in pairs
        if first(f["name"]) != first(t["name"])
    ]
    left = [
        re.compile(rf"\b{re.escape(w)}\b" + (other if w == first(f["name"]) else ""))
        for f, t in pairs
        for w in {*f["name"].split(" "), f.get("email", "")} - {*t["name"].split(" "), ""}
    ]
    return rules, left


def swap(text: str, rules) -> str:
    for rx, to in rules:
        text = rx.sub(to, text)
    return text


def swap_eml(data: bytes, rules) -> bytes:
    head, sep, body = data.partition(b"\n\n") if b"\r\n\r\n" not in data else data.partition(b"\r\n\r\n")
    head_s = swap(head.decode("utf-8"), rules)
    if re.search(r"(?im)^content-transfer-encoding:\s*quoted-printable", head_s):
        decoded = quopri.decodestring(body).decode("utf-8")
        nl = b"\r\n" if b"\r\n" in body else b"\n"
        out = quopri.encodestring(swap(decoded, rules).encode("utf-8"))
        if nl == b"\r\n":
            out = re.sub(rb"(?<!\r)\n", b"\r\n", out)
        return head_s.encode("utf-8") + sep + out
    return head_s.encode("utf-8") + sep + swap(body.decode("utf-8"), rules).encode("utf-8")


def plain(data: bytes, suffix: str) -> str:
    """Readable text for the leftover check (decoded mail body, XML without tags)."""
    if suffix == ".eml":
        head, _, body = data.replace(b"\r\n", b"\n").partition(b"\n\n")
        return head.decode("utf-8") + "\n" + quopri.decodestring(body).decode("utf-8", "replace")
    return data.decode("utf-8", "replace")


def office(src: Path, dst: Path, rules, left, problems):
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith((".xml", ".rels")):
                text = swap(data.decode("utf-8"), rules)
                data = text.encode("utf-8")
                runs = re.sub(r"<[^>]+>", "", text)
                for rx in left:
                    if rx.search(runs) or rx.search(text):
                        problems.append(f"{dst.relative_to(ROOT)}:{item.filename}: '{rx.pattern}' left")
            zout.writestr(item, data)


def main(pid: str):
    profile = json.loads((ROOT / "profiles" / f"{pid}.json").read_text("utf-8"))
    rules, left = load_rules(profile)
    out_root = ROOT / "profiles" / pid
    if out_root.exists():
        shutil.rmtree(out_root)
    demos = [d for d in ROOT.iterdir() if d.is_dir() and (d / "manifest.json").exists()]
    problems, binaries, changed = [], [], 0
    for demo in demos:
        for src in demo.rglob("*"):
            if not src.is_file():
                continue
            dst = out_root / src.relative_to(ROOT)
            dst.parent.mkdir(parents=True, exist_ok=True)
            suffix = src.suffix.lower()
            if suffix in TEXT:
                data = src.read_bytes()
                new = swap_eml(data, rules) if suffix == ".eml" else swap(data.decode("utf-8"), rules).encode("utf-8")
                dst.write_bytes(new)
                changed += new != data
                for rx in left:
                    if rx.search(plain(new, suffix)):
                        problems.append(f"{dst.relative_to(ROOT)}: '{rx.pattern}' left")
            elif suffix in OFFICE:
                office(src, dst, rules, left, problems)
                changed += 1
            else:
                shutil.copy2(src, dst)
                if suffix not in {".png", ".jpg", ".jpeg", ".gif", ".svg"}:
                    binaries.append(str(dst.relative_to(ROOT)))
    print(f"profiles/{pid}: {len(demos)} demos, {changed} text/Office files rewritten")
    for b in binaries:
        print(f"  copied unchanged (check by hand): {b}")
    if problems:
        print("Default names left over:", *problems, sep="\n  ")
        sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])
