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

print("PASS: synthetic package preservation, fingerprints, and private contact rejection")
