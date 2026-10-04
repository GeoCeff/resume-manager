"""Preserve DOCX package parts while changing only the resume contact paragraph."""

from copy import deepcopy
from hashlib import sha256
from pathlib import Path
import argparse
import json
import re
from urllib.parse import unquote
from zipfile import ZipFile

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def load_settings(path):
    global PRIVATE_NAME, PUBLIC_NAME, PRIVATE_CONTACT, PUBLIC_CONTACT, CONTACT_ANCHOR, PRIVATE_ONLY
    settings = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    PRIVATE_NAME = settings.get("ArchivePrivateName", "resume-private.docx")
    PUBLIC_NAME = settings.get("ArchivePublicName", "resume-public.docx")
    PRIVATE_CONTACT = settings["PrivateContact"]
    PUBLIC_CONTACT = settings["PublicContact"]
    CONTACT_ANCHOR = settings["ContactAnchor"]
    public_fields = {field.strip().casefold() for field in PUBLIC_CONTACT.split("|")}
    PRIVATE_ONLY = list(settings["PrivateOnly"])
    PRIVATE_ONLY += [field.strip() for field in PRIVATE_CONTACT.split("|") if field.strip().casefold() not in public_fields]
    if not CONTACT_ANCHOR or CONTACT_ANCHOR not in PRIVATE_CONTACT or CONTACT_ANCHOR not in PUBLIC_CONTACT:
        raise ValueError("ContactAnchor must occur in both contact lines")
    if not settings["PrivateOnly"] or any(not isinstance(field, str) or not field.strip() for field in PRIVATE_ONLY):
        raise ValueError("PrivateOnly must list nonempty private contact values")
    if has_private_details(PUBLIC_CONTACT):
        raise ValueError("PublicContact contains private-only contact details")


def parse_xml(data):
    return etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True))


def contact_paragraph(root):
    body = root.find(W + "body")
    if body is not None:
        for paragraph in body.findall(W + "p")[:6]:
            text = "".join(paragraph.itertext())
            if CONTACT_ANCHOR in text:
                return paragraph
    raise RuntimeError("Contact line not found in the first six body paragraphs; keep it below your name")


def set_contact(path, contact):
    with ZipFile(path) as package:
        parts = [(info, package.read(info.filename)) for info in package.infolist()]
    root = parse_xml(dict((info.filename, data) for info, data in parts)["word/document.xml"])
    paragraph = contact_paragraph(root)
    properties = paragraph.find(".//" + W + "rPr")
    properties = deepcopy(properties) if properties is not None else None
    for child in list(paragraph):
        if child.tag != W + "pPr":
            paragraph.remove(child)
    run = etree.SubElement(paragraph, W + "r")
    if properties is not None:
        run.append(properties)
    etree.SubElement(run, W + "t").text = contact
    replacement = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with ZipFile(path, "w") as package:
        for info, data in parts:
            package.writestr(info, replacement if info.filename == "word/document.xml" else data)


def has_private_details(text):
    decoded = unquote(text).lower()
    for field in PRIVATE_ONLY:
        digits = re.sub(r"\D", "", field)
        if field.casefold() in decoded or (len(digits) >= 7 and not re.search(r"[a-zA-Z]", field) and digits in re.sub(r"\D", "", decoded)):
            return True
    return False


def validate_variants(directory):
    with ZipFile(directory / PRIVATE_NAME) as private, ZipFile(directory / PUBLIC_NAME) as public:
        if set(private.namelist()) != set(public.namelist()):
            raise RuntimeError("Private and public DOCX package parts differ")
        for name in private.namelist():
            a, b = private.read(name), public.read(name)
            if name == "word/document.xml":
                ar, br = parse_xml(a), parse_xml(b)
                ap, bp = contact_paragraph(ar), contact_paragraph(br)
                if "".join(ap.itertext()) != PRIVATE_CONTACT or "".join(bp.itertext()) != PUBLIC_CONTACT:
                    raise RuntimeError("Unexpected contact line")
                for p in (ap, bp):
                    p.getparent().remove(p)
                if etree.tostring(ar) != etree.tostring(br):
                    raise RuntimeError("Private and public documents differ outside the contact paragraph")
            elif a != b:
                raise RuntimeError(f"Private and public package part differs: {name}")
            if name.endswith((".xml", ".rels")):
                xml = parse_xml(b)
                # Inspect split runs, text boxes, comments, metadata, and hyperlink targets too.
                text = "".join(xml.itertext()) + " ".join(value for e in xml.iter() for value in e.attrib.values())
                if has_private_details(text):
                    raise RuntimeError(f"Private contact details found in public {name}; remove them outside the contact line")


def fingerprint(path):
    digest = sha256()
    with ZipFile(path) as package:
        for name in sorted(package.namelist()):
            data = package.read(name)
            if name.endswith((".xml", ".rels")):
                root = parse_xml(data)
                for element in list(root.iter()):
                    if not isinstance(element.tag, str):
                        continue
                    local = etree.QName(element).localname
                    if (name == "docProps/core.xml" and local in {"modified", "revision", "lastModifiedBy"}) or (
                        name == "docProps/app.xml" and local in {"TotalTime", "Pages", "Words", "Characters", "CharactersWithSpaces", "Paragraphs", "Lines"}
                    ) or (name == "word/settings.xml" and local in {"rsids", "proofState"}):
                        element.getparent().remove(element)
                        continue
                    for key in list(element.attrib):
                        if etree.QName(key).localname.startswith("rsid"):
                            del element.attrib[key]
                data = etree.tostring(root, method="c14n")
            digest.update(name.encode())
            digest.update(b"\0")
            digest.update(data)
    return digest.hexdigest()


def validate_pdfs(directory):
    from pypdf import PdfReader

    for name in (PRIVATE_NAME, PUBLIC_NAME):
        reader = PdfReader(directory / Path(name).with_suffix(".pdf"))
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
        if not reader.pages or CONTACT_ANCHOR not in text:
            raise RuntimeError(f"PDF missing readable resume contact: {name}")
        if name == PUBLIC_NAME and has_private_details(text + str(reader.metadata)):
            raise RuntimeError("Private contact details found in public PDF")
        expected_contact = PRIVATE_CONTACT if name == PRIVATE_NAME else PUBLIC_CONTACT
        if "".join(expected_contact.split()) not in "".join(text.split()):
            raise RuntimeError(f"PDF missing expected contact line: {name}")
        print(f"Validated PDF: {name} ({len(reader.pages)} pages)")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", nargs="?", type=Path)
    parser.add_argument("--config", type=Path)
    parser.add_argument("--fingerprint", type=Path)
    parser.add_argument("--validate-pdfs", type=Path)
    args = parser.parse_args()
    if args.fingerprint:
        print(fingerprint(args.fingerprint))
        return
    if not args.config:
        parser.error("--config must point to settings stored outside the program repository")
    load_settings(args.config)
    if args.validate_pdfs:
        validate_pdfs(args.validate_pdfs)
        return
    directory = args.directory or Path.cwd()
    for name, contact in ((PRIVATE_NAME, PRIVATE_CONTACT), (PUBLIC_NAME, PUBLIC_CONTACT)):
        set_contact(directory / name, contact)
    validate_variants(directory)
    print("Validated private/public package preservation and privacy")


if __name__ == "__main__":
    main()
