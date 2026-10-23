"""Verify the generated bilingual corpus and installer contract without tenant access."""
import hashlib
import json
import xml.etree.ElementTree as ET
from email import policy
from email.parser import BytesParser
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parent.parent
FOLDER = ROOT / "workiq-port-operations"
cfg = json.loads((FOLDER / "tenant.json").read_text(encoding="utf-8"))
manifest = json.loads((FOLDER / "manifest.json").read_text(encoding="utf-8-sig"))
gold = json.loads((FOLDER / "tenant" / "expected.json").read_text(encoding="utf-8"))
assert len(manifest["files"]) == gold["physicalDocuments"] == 48
assert set(cfg["workIqShared"]["documents"]) == {f["path"] for f in manifest["files"]}
texts = {}
for file in manifest["files"]:
    path = FOLDER / file["path"]
    assert hashlib.sha256(path.read_bytes()).hexdigest() == file["sha256"], path
    with ZipFile(path) as package:
        xml = ET.fromstring(package.read("word/document.xml"))
        texts[file["path"]] = " ".join(t.text or "" for t in xml.iter() if t.tag.endswith("}t"))
        settings = ET.fromstring(package.read("word/settings.xml"))
        ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        assert settings.find(f"{ns}zoom").get(f"{ns}percent") == "100"
        size = xml.find(f".//{ns}pgSz")
        assert size.get(f"{ns}w") == "12240" and size.get(f"{ns}h") == "15840"
    assert "WIQ26" in texts[file["path"]]
    assert not any(word in texts[file["path"]].lower() for word in ("megan bowen", "adele vance"))
assert len({hashlib.sha256((FOLDER / p).read_bytes()).hexdigest() for p in texts}) == 48
assert "14 containers remain on hold" in texts["en/WIQ-D03_status_v2.docx"]
assert "vier" in texts["de/WIQ-D01_brief.docx"].lower()
for lang in ("en", "de"):
    assert len([p for p in texts if p.startswith(lang + "/")]) == 24
    for index, template in enumerate(cfg["mails"], 1):
        path = FOLDER / template.replace("{lang}", lang)
        msg = BytesParser(policy=policy.default).parsebytes(path.read_bytes())
        assert f"[WIQ26-M{index:02}]" in msg["Subject"]
        assert msg["From"].addresses[0].addr_spec in cfg["requiredUsers"]
        assert len(msg.get_content()) > 150
    second = BytesParser(policy=policy.default).parsebytes((FOLDER / cfg["mails"][1].replace("{lang}", lang)).read_bytes())
    assert ("NOT" if lang == "en" else "NICHT") in second.get_content()
assert not any("tenant/" in f["path"] for f in manifest["files"]), "Test gold and mail templates must not be indexed files."
print("PASS: 24 distinct bilingual documents, 48 verified hashes/Office packages, 8 original EML, stable truths and excluded test gold.")
