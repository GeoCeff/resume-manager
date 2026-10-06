"""Small synthetic-only package and privacy regression check; no user documents."""

from pathlib import Path
import json
import tempfile
from zipfile import ZipFile

import update_resume_projects as program


with tempfile.TemporaryDirectory(prefix="resume-program-test-") as temp:
    directory = Path(temp)
    settings = {
        "PrivateContact": "public@example.com | personal@example.com | 00000000000 | Example City",
        "PublicContact": "public@example.com | Example City",
        "ContactAnchor": "public@example.com",
        "PrivateOnly": ["personal@example.com", "00000000000"],
    }
    config = directory / "settings.local.json"
    config.write_text(json.dumps(settings), encoding="utf-8")
    program.load_settings(config)
    xml = b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Example Name</w:t></w:r></w:p><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>public@example.com | Example City</w:t></w:r></w:p><w:p><w:r><w:t>Example experience</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'
    for name in (program.PRIVATE_NAME, program.PUBLIC_NAME):
        with ZipFile(directory / name, "w") as package:
            package.writestr("word/document.xml", xml)
            package.writestr("word/media/image.bin", b"synthetic-payload")
            package.writestr("word/styles.xml", b"<styles/>")
    for name, contact in ((program.PRIVATE_NAME, program.PRIVATE_CONTACT), (program.PUBLIC_NAME, program.PUBLIC_CONTACT)):
        program.set_contact(directory / name, contact)
    program.validate_variants(directory)
    for name in (program.PRIVATE_NAME, program.PUBLIC_NAME):
        with ZipFile(directory / name) as package:
            assert package.read("word/media/image.bin") == b"synthetic-payload"
            assert package.read("word/styles.xml") == b"<styles/>"
    assert program.fingerprint(directory / program.PRIVATE_NAME) != program.fingerprint(directory / program.PUBLIC_NAME)
    assert program.has_private_details("personal%40example.com")
    assert program.has_private_details("000-0000-0000")
    assert not program.has_private_details(program.PUBLIC_CONTACT)
    previous_private = program.PRIVATE_ONLY
    program.PRIVATE_ONLY = previous_private + ['Stra'+chr(0xDF)+'e']
    assert program.has_private_details('Stra'+chr(0xDF)+'e') and program.has_private_details('STRASSE')
    program.PRIVATE_ONLY = previous_private
    # Retained field styles and unrelated relationships survive linked contact replacement.
    linked = b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Example Name</w:t></w:r></w:p><w:p><w:hyperlink r:id="rId1"><w:r><w:rPr><w:b/></w:rPr><w:t>public@example.com</w:t></w:r></w:hyperlink><w:r><w:t> | </w:t></w:r><w:hyperlink r:id="rId2"><w:r><w:t>personal@example.com</w:t></w:r></w:hyperlink><w:r><w:t> | 00000000000 | </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>Example City</w:t></w:r></w:p><w:p><w:hyperlink r:id="rId3"><w:r><w:t>Other link</w:t></w:r></w:hyperlink></w:p></w:body></w:document>'
    relationships = b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mailto:public@example.com" TargetMode="External"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mailto:personal@example.com" TargetMode="External"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org" TargetMode="External"/></Relationships>'
    for name, contact in ((program.PRIVATE_NAME,program.PRIVATE_CONTACT),(program.PUBLIC_NAME,program.PUBLIC_CONTACT)):
        with ZipFile(directory/name,'w') as package:
            package.writestr('word/document.xml',linked); package.writestr(program.RELS_PART,relationships)
            package.writestr('word/media/image.bin',b'synthetic-payload')
        program.set_contact(directory/name,contact)
    program.validate_variants(directory)
    with ZipFile(directory/program.PUBLIC_NAME) as package:
        root = program.parse_xml(package.read('word/document.xml'))
        assert len(list(program.contact_paragraph(root).iter(program.W+'b'))) == 1
        assert len(list(program.contact_paragraph(root).iter(program.W+'i'))) == 1
        rels = program.parse_xml(package.read(program.RELS_PART))
        assert not any('personal@example.com' in e.get('Target','') for e in rels)
        assert any(e.get('Id')=='rId3' and e.get('Target')=='https://example.org' for e in rels)
    # A shared private contact relationship must fail, not silently change unrelated content.
    shared = linked.replace(b'r:id="rId3"',b'r:id="rId2"')
    for name,contact in ((program.PRIVATE_NAME,program.PRIVATE_CONTACT),(program.PUBLIC_NAME,program.PUBLIC_CONTACT)):
        with ZipFile(directory/name,'w') as package:
            package.writestr('word/document.xml',shared); package.writestr(program.RELS_PART,relationships)
        program.set_contact(directory/name,contact)
    try: program.validate_variants(directory)
    except RuntimeError as error: assert 'Private contact details' in str(error)
    else: raise AssertionError('Shared private target escaped privacy validation')
    # Restore plain packages for subsequent hidden-text checks.
    for name,contact in ((program.PRIVATE_NAME,program.PRIVATE_CONTACT),(program.PUBLIC_NAME,program.PUBLIC_CONTACT)):
        with ZipFile(directory/name,'w') as package: package.writestr('word/document.xml',xml)
        program.set_contact(directory/name,contact)
    for name in (program.PRIVATE_NAME, program.PUBLIC_NAME):
        with ZipFile(directory / name, "a") as package:
            package.writestr("word/comments.xml", b'<comments><t>personal@</t><t>example.com</t></comments>')
    try:
        program.validate_variants(directory)
    except RuntimeError as error:
        assert "Private contact details" in str(error)
    else:
        raise AssertionError("Split private details were not rejected")
    settings["PublicContact"] += " | personal@example.com"
    config.write_text(json.dumps(settings), encoding="utf-8")
    try:
        program.load_settings(config)
    except ValueError:
        pass
    else:
        raise AssertionError("Leaking contact settings were not rejected")
    unsupported = directory/'unsupported.docx'
    with ZipFile(unsupported,'w') as package:
        package.writestr('word/document.xml',xml.replace(b'<w:t>public@',b'<w:tab/><w:t>public@'))
    unchanged = unsupported.read_bytes()
    try: program.set_contact(unsupported,program.PUBLIC_CONTACT)
    except ValueError as error: assert 'Non-text contact' in str(error)
    else: raise AssertionError('Unsupported contact run was silently flattened')
    assert unchanged == unsupported.read_bytes()
    inspected = directory/'inspect.docx'
    with ZipFile(inspected,'w') as package:
        package.writestr('[Content_Types].xml',b'<Types/>');package.writestr('word/document.xml',xml)
    source_bytes = inspected.read_bytes()
    assert program.inspect_contact(inspected)['Fields'] == ['public@example.com','Example City']
    assert source_bytes == inspected.read_bytes()
    assert program.inspect_contact(inspected,'public@example.com')['ContactLine'] == 'public@example.com | Example City'
    with ZipFile(inspected,'w') as package:
        package.writestr('[Content_Types].xml',b'<Types/>');package.writestr('word/document.xml',xml.replace(b'public@example.com | Example City',b'public@example.com'))
    assert program.inspect_contact(inspected)['Fields'] == ['public@example.com']
    with ZipFile(inspected,'w') as package:
        package.writestr('[Content_Types].xml',b'<Types/>')
        package.writestr('word/document.xml',xml.replace(b'</w:body>',b'<w:p><w:r><w:t>other@example.com | Other City</w:t></w:r></w:p></w:body>'))
    try: program.inspect_contact(inspected)
    except ValueError as error: assert 'one contact line' in str(error)
    else: raise AssertionError('Ambiguous contact detection was accepted')
    assert program.inspect_contact(inspected,'public@example.com')['Fields'][0] == 'public@example.com'
    with ZipFile(inspected,'w') as package:
        package.writestr('[Content_Types].xml',b'<Types/>');package.writestr('word/document.xml',xml.replace(b'<w:t>public@',b'<w:tab/><w:t>public@'))
    try: program.inspect_contact(inspected)
    except ValueError as error: assert 'Non-text contact' in str(error)
    else: raise AssertionError('Unsupported inspection markup was accepted')
    disguised = directory/'renamed.docx'
    with ZipFile(disguised,'w') as package:
        package.writestr('word/document.xml',xml); package.writestr('word/vbaProject.bin',b'synthetic-macro-marker')
    try: program.check_source(disguised)
    except ValueError as error: assert 'Macro-enabled' in str(error)
    else: raise AssertionError('Disguised macro package was accepted')

with tempfile.TemporaryDirectory(prefix='resume-export-test-') as temp:
    directory = Path(temp)
    example = directory/'example.docx'
    program.create_example(example)
    original = example.read_bytes()
    assert program.inspect_contact(example)['Fields'][0] == 'public@example.com'
    try: program.create_example(example)
    except FileExistsError: pass
    else: raise AssertionError('Example generation overwrote an existing file')
    assert example.read_bytes() == original
    config = directory/'settings.local.json'
    config.write_text(json.dumps({'PrivateContact':'public@example.com | personal@example.com | 00000000000 | Example City',
                                 'PublicContact':'public@example.com | Example City','ContactAnchor':'public@example.com',
                                 'PrivateOnly':['personal@example.com','00000000000']}), encoding='utf-8')
    program.load_settings(config)
    for name, contact in ((program.PRIVATE_NAME,program.PRIVATE_CONTACT),(program.PUBLIC_NAME,program.PUBLIC_CONTACT)):
        (directory/name).write_bytes(original); program.set_contact(directory/name,contact)
    program.validate_variants(directory)
    program.validate_export(directory/program.PRIVATE_NAME,'Private')
    program.validate_export(directory/program.PUBLIC_NAME,'Public')
    assert program.compare_text(directory/program.PUBLIC_NAME,directory/program.PUBLIC_NAME) == 'No body text changes.'
    changed = directory/'changed.docx'
    with ZipFile(directory/program.PUBLIC_NAME) as source, ZipFile(changed,'w') as destination:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == 'word/document.xml': data = data.replace(b'Example Project',b'Updated Project')
            destination.writestr(info,data)
    assert '+Updated Project' in program.compare_text(directory/program.PUBLIC_NAME,changed)
    with ZipFile(changed,'a') as package:
        package.writestr('docProps/custom.xml',b'<properties><value>personal@example.com</value></properties>')
    try: program.validate_export(changed,'Public')
    except ValueError as error: assert 'private details' in str(error)
    else: raise AssertionError('Hidden private contact escaped export validation')

print("PASS: synthetic preservation, linked contact styles, privacy, fingerprints, macro rejection, example, export, and text comparison")
