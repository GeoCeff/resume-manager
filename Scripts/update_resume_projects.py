"""Preserve DOCX package parts while changing only the resume contact paragraph."""

from copy import deepcopy
from hashlib import sha256
from pathlib import Path
import argparse
import difflib
import json
import re
from urllib.parse import unquote
from zipfile import ZipFile

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
RELS_PART = 'word/_rels/document.xml.rels'


def text_of(element):
    return ''.join(element.itertext()) if element.tag != W + 'p' else ''.join(e.text or '' for e in element.iter(W + 't'))


def check_source(path, strict=False):
    if Path(path).suffix.lower() != '.docx':
        raise ValueError('Keep the resume in DOCX format; macro-enabled documents are not supported')
    with ZipFile(path) as package:
        names = package.namelist()
        if len(names) != len(set(names)):
            raise ValueError('DOCX contains duplicate package entries')
        if 'word/document.xml' not in names:
            raise ValueError('DOCX is missing its document body')
        types = package.read('[Content_Types].xml').lower() if '[Content_Types].xml' in names else b''
        if any('vbaproject' in name.lower() for name in names) or b'macroenabled' in types or b'vbaproject' in types:
            raise ValueError('Macro-enabled package rejected, even when renamed to DOCX')
        if strict and not types:
            raise ValueError('DOCX is missing its content-type declarations')


def load_settings(path):
    global PRIVATE_NAME, PUBLIC_NAME, PRIVATE_CONTACT, PUBLIC_CONTACT, CONTACT_ANCHOR, PRIVATE_ONLY
    settings = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    if not isinstance(settings.get('PrivateOnly'),list):
        raise ValueError('PrivateOnly must be an array of private contact values')
    PRIVATE_NAME = settings.get("ArchivePrivateName", "resume-private.docx")
    PUBLIC_NAME = settings.get("ArchivePublicName", "resume-public.docx")
    PRIVATE_CONTACT = settings["PrivateContact"]
    PUBLIC_CONTACT = settings["PublicContact"]
    CONTACT_ANCHOR = settings["ContactAnchor"]
    if any(not isinstance(value,str) or '\n' in value or '\r' in value for value in (PRIVATE_CONTACT,PUBLIC_CONTACT,CONTACT_ANCHOR)):
        raise ValueError('Contact lines and anchor must be single-line strings')
    if any(value != ' | '.join(field.strip() for field in value.split('|')) for value in (PRIVATE_CONTACT,PUBLIC_CONTACT)):
        raise ValueError('Separate contact fields with a space, pipe, and space')
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
            text = text_of(paragraph)
            if CONTACT_ANCHOR in text:
                return paragraph
    raise RuntimeError("Contact line not found in the first six body paragraphs; keep it below your name")


def validate_contact_markup(paragraph):
    for child in paragraph:
        if child.tag not in {W+'pPr', W+'r', W+'hyperlink', W+'bookmarkStart', W+'bookmarkEnd', W+'proofErr'}:
            raise ValueError('Unsupported contact markup; use plain runs or email hyperlinks in the contact line')
    if any(paragraph.iter(W+'fldChar')) or any(paragraph.iter(W+'instrText')) or any(paragraph.iter(W+'drawing')):
        raise ValueError('Fields or graphics in the contact line require manual adjustment; source was not changed')
    if any(child.tag not in {W+'rPr',W+'t'} for run in paragraph.iter(W+'r') for child in run):
        raise ValueError('Non-text contact runs require manual adjustment; source was not changed')


def inspect_contact(path, anchor=None):
    check_source(path, strict=True)
    with ZipFile(path) as package:
        body = parse_xml(package.read('word/document.xml')).find(W+'body')
    if body is None:
        raise ValueError('Resume body is missing')
    # ponytail: contact line near the name; ambiguous layouts need Advanced settings.
    candidates = [p for p in body.findall(W+'p')[:6] if
                  (anchor in text_of(p) if anchor else ('|' in text_of(p) and re.search(r'[^\s|@]+@[^\s|@]+\.[^\s|@]+', text_of(p))) or re.fullmatch(r'[^\s|@]+@[^\s|@]+\.[^\s|@]+',text_of(p).strip()))]
    if len(candidates) != 1:
        raise ValueError('Could not identify one contact line. Use Advanced settings to specify contact text to locate.')
    paragraph = candidates[0]
    validate_contact_markup(paragraph)
    fields = [field.strip() for field in text_of(paragraph).split('|')]
    if any(not field or '\n' in field or '\r' in field for field in fields) or len({field.casefold() for field in fields}) != len(fields):
        raise ValueError('Use unique contact details separated by a pipe in one paragraph, or configure contacts manually.')
    return {'Fields': fields, 'ContactLine': ' | '.join(fields)}


def set_contact(path, contact):
    check_source(path)
    with ZipFile(path) as package:
        parts = [(info, package.read(info.filename)) for info in package.infolist()]
    original = {info.filename: data for info, data in parts}
    root = parse_xml(original["word/document.xml"])
    paragraph = contact_paragraph(root)
    validate_contact_markup(paragraph)
    old_text = text_of(paragraph)
    spans = {}
    for match in re.finditer(r'[^|]+', old_text):
        field = match.group().strip()
        if field:
            start = match.start() + len(match.group()) - len(match.group().lstrip())
            spans[field] = (start, start+len(field))
    rels = parse_xml(original[RELS_PART]) if RELS_PART in original else etree.Element(REL+'Relationships', nsmap={None:REL[1:-1]})
    old_ids = {e.get(R+'id') for e in paragraph.iter(W+'hyperlink') if e.get(R+'id')}
    link_style = next((deepcopy(e.find('.//'+W+'rPr')) for e in paragraph.iter(W+'hyperlink') if e.find('.//'+W+'rPr') is not None), None)
    linked_emails = any('@' in text_of(e) for e in paragraph.iter(W+'hyperlink'))
    properties = paragraph.find(".//" + W + "rPr")
    properties = deepcopy(properties) if properties is not None else None
    original_children = [deepcopy(child) for child in paragraph if child.tag in {W+'r', W+'hyperlink'}]
    bookmarks = [deepcopy(child) for child in paragraph if child.tag in {W+'bookmarkStart',W+'bookmarkEnd'}]

    def segment(start, end):
        position = 0
        result = []
        for original_child in original_children:
            child = deepcopy(original_child)
            kept = False
            for node in list(child.iter(W+'t')):
                value = node.text or ''
                left, right = max(0,start-position), min(len(value),end-position)
                position += len(value)
                node.text = value[left:right] if right > left else ''
                if node.text:
                    kept = True
                    node.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
            for run in list(child.iter(W+'r')):
                if not any(node.text for node in run.iter(W+'t')) and run.getparent() is not None:
                    run.getparent().remove(run)
            if kept:
                result.append(child)
        return result

    def new_run(value, style):
        run = etree.Element(W+'r')
        if style is not None:
            run.append(deepcopy(style))
        node = etree.SubElement(run,W+'t'); node.text = value
        node.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
        return run

    for child in list(paragraph):
        if child.tag != W + "pPr":
            paragraph.remove(child)
    for bookmark in bookmarks:
        if bookmark.tag == W+'bookmarkStart': paragraph.append(bookmark)
    for index, field in enumerate(contact.split(' | ')):
        if index:
            separator = re.search(r'\s*\|\s*',old_text)
            if separator and separator.group() == ' | ':
                paragraph.extend(segment(separator.start(),separator.end()))
            else:
                paragraph.append(new_run(' | ',properties))
        if field in spans:
            paragraph.extend(segment(*spans[field]))
        elif '@' in field and linked_emails:
            identifier = 'rIdContact1'
            counter = 1
            used = {e.get('Id') for e in rels}
            while identifier in used:
                counter += 1; identifier = f'rIdContact{counter}'
            etree.SubElement(rels,REL+'Relationship',Id=identifier,Type=R[1:-1]+'/hyperlink',Target='mailto:'+field,TargetMode='External')
            link = etree.SubElement(paragraph,W+'hyperlink'); link.set(R+'id',identifier)
            link.append(new_run(field,link_style if link_style is not None else properties))
        else:
            paragraph.append(new_run(field,properties))
    for bookmark in bookmarks:
        if bookmark.tag == W+'bookmarkEnd': paragraph.append(bookmark)
    referenced = {e.get(R+'id') for e in root.iter() if e.get(R+'id')}
    for relationship in list(rels):
        if relationship.get('Id') in old_ids - referenced:
            rels.remove(relationship)
    replacement = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with ZipFile(path, "w") as package:
        for info, data in parts:
            if info.filename == 'word/document.xml': data = replacement
            elif info.filename == RELS_PART and (old_ids or linked_emails): data = etree.tostring(rels,xml_declaration=True,encoding='UTF-8',standalone=True)
            package.writestr(info,data)


def has_private_details(text):
    decoded = unquote(text).casefold()
    for field in PRIVATE_ONLY:
        digits = re.sub(r"\D", "", field)
        if field.casefold() in decoded or (len(digits) >= 7 and not re.search(r"[a-zA-Z]", field) and digits in re.sub(r"\D", "", decoded)):
            return True
    return False


def validate_variants(directory):
    with ZipFile(directory / PRIVATE_NAME) as private, ZipFile(directory / PUBLIC_NAME) as public:
        if set(private.namelist()) != set(public.namelist()):
            raise RuntimeError("Private and public DOCX package parts differ")
        ar, br = parse_xml(private.read('word/document.xml')), parse_xml(public.read('word/document.xml'))
        ap, bp = contact_paragraph(ar), contact_paragraph(br)
        contact_ids = {e.get(R+'id') for p in (ap,bp) for e in p.iter() if e.get(R+'id')}
        for p in (ap,bp): p.getparent().remove(p)
        other_ids = {e.get(R+'id') for root in (ar,br) for e in root.iter() if e.get(R+'id')}
        contact_ids -= other_ids
        if text_of(ap) != PRIVATE_CONTACT or text_of(bp) != PUBLIC_CONTACT:
            raise RuntimeError('Unexpected contact line')
        if etree.tostring(ar) != etree.tostring(br):
            raise RuntimeError('Private and public documents differ outside the contact paragraph')
        for name in private.namelist():
            a, b = private.read(name), public.read(name)
            if name == "word/document.xml":
                pass
            elif name == RELS_PART:
                roots = [parse_xml(a),parse_xml(b)]
                for relroot in roots:
                    for relationship in list(relroot):
                        if relationship.get('Id') in contact_ids:
                            relroot.remove(relationship)
                if etree.tostring(roots[0]) != etree.tostring(roots[1]):
                    raise RuntimeError('Non-contact relationships differ')
            elif a != b:
                raise RuntimeError(f"Private and public package part differs: {name}")
            if name.endswith((".xml", ".rels")):
                xml = parse_xml(b)
                # Inspect split runs, text boxes, comments, metadata, and hyperlink targets too.
                text = "".join(xml.itertext()) + " ".join(value for e in xml.iter() for value in e.attrib.values())
                if has_private_details(text):
                    raise RuntimeError(f"Private contact details found in public {name}; remove them outside the contact line")


def fingerprint(path):
    check_source(path)
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


def review(directory):
    from pypdf import PdfReader
    validate_variants(directory)
    pages = {}
    warnings = []
    for name, key in ((PRIVATE_NAME,'PrivatePages'),(PUBLIC_NAME,'PublicPages')):
        pdf = directory / Path(name).with_suffix('.pdf')
        if not pdf.is_file():
            raise ValueError('Archived PDF missing; this version is incomplete')
        reader = PdfReader(pdf)
        text = '\n'.join(page.extract_text() or '' for page in reader.pages)
        if key == 'PublicPages' and has_private_details(text+str(reader.metadata)):
            raise ValueError('Configured private details found in public PDF')
        pages[key] = len(reader.pages)
        if len(reader.pages)>1: warnings.append(f'{key.replace("Pages", "")} resume has more than one page; review layout')
    with ZipFile(directory/PUBLIC_NAME) as package:
        for name in package.namelist():
            if name.startswith('word/embeddings/'): warnings.append('Embedded file requires manual review')
            if name.startswith('word/media/'): warnings.append('Images are not checked for private text')
            if 'comments' in name.lower(): warnings.append('Comments require manual review')
            if name.endswith(('.xml','.rels')):
                root = parse_xml(package.read(name))
                if any(e.tag in {W+'ins',W+'del',W+'moveFrom',W+'moveTo'} for e in root.iter()): warnings.append('Tracked changes require manual review')
                if any(e.get('TargetMode')=='External' for e in root.iter()): warnings.append('External linked targets require manual review')
                if name=='docProps/core.xml' and any(etree.QName(e).localname in {'creator','lastModifiedBy'} and e.text for e in root): warnings.append('Author metadata requires manual review')
    return dict(pages, Checks='Configured contact checks passed; this is not an anonymity guarantee.', Warnings=list(dict.fromkeys(warnings)) or ['Review the public PDF before sharing.'])


def validate_export(path, kind):
    expected = PUBLIC_CONTACT if kind == 'Public' else PRIVATE_CONTACT
    if path.suffix.lower() == '.docx':
        check_source(path, strict=True)
        with ZipFile(path) as package:
            if text_of(contact_paragraph(parse_xml(package.read('word/document.xml')))) != expected:
                raise ValueError('Export contact line does not match the active privacy settings')
            for name in package.namelist():
                if kind == 'Public' and name.endswith(('.xml', '.rels')):
                    root = parse_xml(package.read(name))
                    text = ''.join(root.itertext()) + ' '.join(value for e in root.iter() for value in e.attrib.values())
                    if has_private_details(text):
                        raise ValueError('Configured private details found in public DOCX export')
    elif path.suffix.lower() == '.pdf':
        from pypdf import PdfReader
        reader = PdfReader(path)
        text = '\n'.join(page.extract_text() or '' for page in reader.pages)
        if not reader.pages or ''.join(expected.split()) not in ''.join(text.split()):
            raise ValueError('Export PDF is missing its expected readable contact line')
        if kind == 'Public' and has_private_details(text + str(reader.metadata)):
            raise ValueError('Configured private details found in public PDF export')
    else:
        raise ValueError('Only DOCX and PDF exports are supported')


def compare_text(first, second):
    def lines(path):
        check_source(path, strict=True)
        with ZipFile(path) as package:
            # ponytail: body text only; existing PDF previews cover visual/layout differences.
            body = parse_xml(package.read('word/document.xml')).find(W+'body')
            return [''.join(e.text or '' for e in p.iter() if e.tag in {W+'t', W+'delText'})
                    for p in body.iter(W+'p')] if body is not None else []
    return '\n'.join(difflib.unified_diff(lines(first), lines(second), fromfile=first.parent.name,
                                         tofile=second.parent.name, lineterm='')) or 'No body text changes.'


def create_example(path):
    # A synthetic OOXML package avoids personal Normal.dotm templates and needs no new library.
    paragraphs = [
        ('Example Candidate', 44, True),
        ('public@example.com | personal@example.com | Example City', 22, False),
        ('Developer seeking an entry level role', 22, False),
        ('Education', 24, True), ('Example University - Computing', 22, False),
        ('Experience', 24, True), ('Built a local document versioning tool and tested recovery paths.', 22, False),
        ('Projects', 24, True), ('Example Project - Python and native Windows tooling', 22, False),
        ('Skills', 24, True), ('Python, PowerShell, Git, documentation', 22, False),
    ]
    root = etree.Element(W+'document', nsmap={'w': W[1:-1]})
    body = etree.SubElement(root, W+'body')
    for value, size, bold in paragraphs:
        p = etree.SubElement(body, W+'p')
        properties = etree.SubElement(p, W+'pPr')
        etree.SubElement(properties, W+'spacing', {W+'after': '160'})
        if value == 'Example Candidate':
            etree.SubElement(properties, W+'pStyle', {W+'val': 'Title'})
        run = etree.SubElement(p, W+'r'); style = etree.SubElement(run, W+'rPr')
        etree.SubElement(style, W+'rFonts', {W+'ascii': 'Calibri', W+'hAnsi': 'Calibri'})
        etree.SubElement(style, W+'sz', {W+'val': str(size)})
        if bold: etree.SubElement(style, W+'b')
        etree.SubElement(run, W+'t').text = value
    section = etree.SubElement(body, W+'sectPr')
    etree.SubElement(section, W+'pgSz', {W+'w': '12240', W+'h': '15840'})
    etree.SubElement(section, W+'pgMar', {W+'top': '1080', W+'bottom': '1080', W+'left': '1080', W+'right': '1080'})
    parts = {
        '[Content_Types].xml': b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
        '_rels/.rels': b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
        'word/document.xml': etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True),
    }
    if path.suffix.lower() != '.docx': raise ValueError('Choose a DOCX filename')
    with ZipFile(path, 'x') as package:
        for name, data in parts.items(): package.writestr(name, data)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", nargs="?", type=Path)
    parser.add_argument("--config", type=Path)
    parser.add_argument("--fingerprint", type=Path)
    parser.add_argument("--validate-pdfs", type=Path)
    parser.add_argument('--check-source',type=Path)
    parser.add_argument('--validate-settings',action='store_true')
    parser.add_argument('--review',type=Path)
    parser.add_argument('--inspect-contact',type=Path)
    parser.add_argument('--anchor')
    parser.add_argument('--compare-text', nargs=2, type=Path)
    parser.add_argument('--validate-export', type=Path)
    parser.add_argument('--variant', choices=['Private', 'Public'], default='Public')
    parser.add_argument('--create-example', type=Path)
    args = parser.parse_args()
    if args.fingerprint:
        print(fingerprint(args.fingerprint))
        return
    if args.inspect_contact:
        print(json.dumps(inspect_contact(args.inspect_contact,args.anchor))); return
    if args.compare_text:
        print(json.dumps(compare_text(*args.compare_text))); return
    if args.create_example:
        create_example(args.create_example); print('Synthetic example created'); return
    if not args.config:
        parser.error("--config must point to settings stored outside the program repository")
    load_settings(args.config)
    if args.validate_settings:
        print('Settings validated'); return
    if args.validate_export:
        validate_export(args.validate_export, args.variant); print('Export validated'); return
    if args.check_source:
        check_source(args.check_source,strict=True)
        with ZipFile(args.check_source) as package: contact_paragraph(parse_xml(package.read('word/document.xml')))
        print('Source validated'); return
    if args.review:
        print(json.dumps(review(args.review))); return
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
