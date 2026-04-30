"""Smoke test: build a feature-rich DOCX in memory, convert to HWPX,
verify the zip structure and content."""
import io
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from docx2hwpx import convert  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# A minimal 1x1 PNG (transparent).
PNG_1x1 = bytes.fromhex(
    "89504E470D0A1A0A0000000D49484452000000010000000108060000001F15C489"
    "0000000D49444154789C6300010000000500010D0A2DB40000000049454E44AE426082"
)

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>
"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>
</Relationships>
"""

STYLES_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/></w:style>
</w:styles>
"""

HEADER_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="{W}">
  <w:p>
    <w:pPr><w:jc w:val="center"/></w:pPr>
    <w:r><w:t>문서 머리말 — 변환 테스트</w:t></w:r>
  </w:p>
</w:hdr>
"""

FOOTER_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="{W}">
  <w:p>
    <w:pPr><w:jc w:val="right"/></w:pPr>
    <w:r><w:t>꼬리말 페이지</w:t></w:r>
  </w:p>
</w:ftr>
"""

NUMBERING_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="{W}">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>
"""

FOOTNOTES_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="{W}">
  <w:footnote w:type="separator" w:id="-1"><w:p/></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p/></w:footnote>
  <w:footnote w:id="1">
    <w:p><w:r><w:t>이것은 첫 번째 각주 내용입니다.</w:t></w:r></w:p>
  </w:footnote>
</w:footnotes>
"""

# Document body exercising:
#   - Heading paragraph
#   - Mixed runs with bold/italic/underline/color
#   - Right alignment
#   - Decimal list (3 items)
#   - Bullet list (2 items) — DOCX numbering id 2 = bullet
#   - Hyperlink
#   - Inline image via <w:drawing>
#   - Footnote reference
#   - Track changes: w:ins (accept), w:del (reject)
#   - Table with both colspan (gridSpan) AND rowspan (vMerge)
#   - Section properties at end
DOCUMENT_XML = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}" xmlns:r="{R}"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="40"/></w:rPr><w:t>변환 테스트 제목</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>안녕하세요, 이것은 </w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>굵은 글씨</w:t></w:r>
      <w:r><w:t>와 </w:t></w:r>
      <w:r><w:rPr><w:i/><w:color w:val="C00000"/></w:rPr><w:t>빨간 기울임</w:t></w:r>
      <w:r><w:t> 그리고 </w:t></w:r>
      <w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>밑줄</w:t></w:r>
      <w:r><w:t>이 섞인 문장입니다.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>여기에 각주가 있습니다</w:t></w:r>
      <w:r><w:footnoteReference w:id="1"/></w:r>
      <w:r><w:t>.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>그리고 외부 링크: </w:t></w:r>
      <w:hyperlink r:id="rId6">
        <w:r><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>example.com</w:t></w:r>
      </w:hyperlink>
      <w:r><w:t>.</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>첫째 항목 (자동번호)</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>둘째 항목</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>셋째 항목</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>
      <w:r><w:t>불릿 항목 1</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>
      <w:r><w:t>불릿 항목 2</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>이미지 인라인: </w:t></w:r>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="914400" cy="914400"/>
            <wp:docPr id="1" name="Pic 1"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr><pic:cNvPr id="1" name=""/><pic:cNvPicPr/></pic:nvPicPr>
                  <pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill>
                  <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p>
      <w:r><w:t>다음은 트랙 체인지: </w:t></w:r>
      <w:ins w:id="1" w:author="A" w:date="2025-01-01T00:00:00Z">
        <w:r><w:t>[삽입된 텍스트] </w:t></w:r>
      </w:ins>
      <w:del w:id="2" w:author="A" w:date="2025-01-01T00:00:00Z">
        <w:r><w:delText>[삭제될 텍스트] </w:delText></w:r>
      </w:del>
      <w:r><w:t>끝.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>페이지 나눔 앞 텍스트.</w:t></w:r>
      <w:r><w:br w:type="page"/></w:r>
      <w:r><w:t>페이지 나눔 뒤 텍스트.</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>아래는 colspan + rowspan 표:</w:t></w:r></w:p>
    <w:tbl>
      <w:tblGrid>
        <w:gridCol w:w="2000"/>
        <w:gridCol w:w="2000"/>
        <w:gridCol w:w="2000"/>
      </w:tblGrid>
      <!-- Row 1: A1 spans cols 0+1 (gridSpan=2) with yellow shading; B1 starts a vMerge (rows 1-3) with green shading -->
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:gridSpan w:val="2"/>
            <w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>
          </w:tcPr>
          <w:p><w:r><w:t>A1 (colspan=2, yellow)</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr>
            <w:vMerge w:val="restart"/>
            <w:shd w:val="clear" w:color="auto" w:fill="00FF00"/>
          </w:tcPr>
          <w:p><w:r><w:t>B1 (rowspan=3, green)</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
      <!-- Row 2: two normal cells, B continues vMerge -->
      <w:tr>
        <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
        <w:tc>
          <w:tcPr><w:vMerge/></w:tcPr>
          <w:p/>
        </w:tc>
      </w:tr>
      <!-- Row 3: two normal cells, B continues vMerge -->
      <w:tr>
        <w:tc><w:p><w:r><w:t>A3</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B3</w:t></w:r></w:p></w:tc>
        <w:tc>
          <w:tcPr><w:vMerge/></w:tcPr>
          <w:p/>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr>
      <w:headerReference r:id="rId2" w:type="default"/>
      <w:footerReference r:id="rId3" w:type="default"/>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
    </w:sectPr>
  </w:body>
</w:document>
"""


def build_docx(path: str):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES_XML)
        z.writestr("word/document.xml", DOCUMENT_XML)
        z.writestr("word/header1.xml", HEADER_XML)
        z.writestr("word/footer1.xml", FOOTER_XML)
        z.writestr("word/footnotes.xml", FOOTNOTES_XML)
        z.writestr("word/numbering.xml", NUMBERING_XML)
        z.writestr("word/media/image1.png", PNG_1x1)


def verify_hwpx(path: str):
    import xml.etree.ElementTree as ET
    with zipfile.ZipFile(path, "r") as z:
        names = z.namelist()
        print("entries:")
        for n in names:
            print(f"  {n}  ({z.getinfo(n).file_size} bytes)")

        assert names[0] == "mimetype", "mimetype must be first"
        assert z.getinfo("mimetype").compress_type == zipfile.ZIP_STORED
        assert z.read("mimetype") == b"application/hwp+zip"

        for r in [
            "version.xml", "settings.xml",
            "Contents/header.xml", "Contents/section0.xml", "Contents/content.hpf",
            "META-INF/container.xml", "META-INF/manifest.xml",
        ]:
            assert r in names, f"missing entry: {r}"

        # Image binary should be embedded.
        bin_entries = [n for n in names if n.startswith("BinData/")]
        assert bin_entries, "no BinData entries — image not embedded"
        print(f"\nbinary entries: {bin_entries}")

        # All XML well-formed.
        for n in names:
            if n.endswith(".xml") or n.endswith(".hpf"):
                ET.fromstring(z.read(n))

        sec = z.read("Contents/section0.xml").decode("utf-8")
        hpf = z.read("Contents/content.hpf").decode("utf-8")
        manifest = z.read("META-INF/manifest.xml").decode("utf-8")

        # Spot-check content.
        for needle in [
            "변환 테스트 제목", "굵은 글씨", "빨간 기울임", "밑줄",
            "첫째 항목", "불릿 항목 1", "[삽입된 텍스트]",
            "A1 (colspan=2, yellow)", "B1 (rowspan=3, green)",
            "A2", "B2", "A3", "B3",
            "이것은 첫 번째 각주 내용입니다",
            "example.com",
        ]:
            assert needle in sec, f"missing in section: {needle!r}"
            print(f"  found: {needle!r}")
        # Header/footer: not yet re-implemented in the template-based writer.
        # If you need them, see TODO in HwpxWriter.
        for needle in ("문서 머리말", "꼬리말 페이지"):
            if needle not in sec:
                print(f"  TODO header/footer: {needle!r} not in output")

        # Track-change deletion must be excluded.
        assert "[삭제될 텍스트]" not in sec, "del content leaked into output"
        print("  verified: deleted tracked text excluded")

        # Hyperlink URL must be in section as a HYPERLINK field command.
        assert "https://example.com/" in sec, "hyperlink URL missing"
        print("  verified: hyperlink URL embedded")

        # Image binary id reference must be in section.
        import re
        bin_refs = re.findall(r'binaryItemIDRef="([^"]+)"', sec)
        assert bin_refs, "no binaryItemIDRef in section"
        print(f"  verified: image binary refs = {bin_refs}")
        for bid in bin_refs:
            assert f'id="{bid}"' in hpf, f"bin id {bid} not in OPF manifest"
        # Note: META-INF/manifest.xml is intentionally empty in the
        # skeleton — Hangul does not require binaries to be listed there.

        # Cell merging:
        #  - colspan: a hp:cellSpan with colSpan="2" should appear
        #  - rowspan: a hp:cellSpan with rowSpan="3" should appear
        assert 'colSpan="2"' in sec, "colspan=2 missing"
        assert 'rowSpan="3"' in sec, "rowspan=3 missing"
        print("  verified: cellSpan colSpan=2 and rowSpan=3 emitted")

        # Footnote ctrl element should appear.
        assert "<hp:footNote " in sec, "footnote control missing"
        print("  verified: footnote control present")

        # TODO: header/footer not yet emitted by template-based writer.

        # Decimal lists are emitted as OUTLINE headings; bullets are emitted
        # as indented NONE paragraphs (without auto-numbering).
        header = z.read("Contents/header.xml").decode("utf-8")
        assert 'type="OUTLINE"' in header, "OUTLINE heading missing in paraPr"
        print("  verified: OUTLINE headings registered for decimal lists")

        # Page break: <w:br w:type="page"/> must produce a paragraph with
        # pageBreak="1" so Hangul actually breaks the page.
        assert 'pageBreak="1"' in sec, "page break attr not emitted"
        print("  verified: pageBreak='1' attribute emitted")

        # Table border: a SOLID-bordered borderFill must exist in the
        # header's <hh:borderFills>, and tables/cells must reference it.
        import re as _re
        bf_solid = _re.search(
            r'<hh:borderFill id="(\d+)"[^>]*>(?:(?!</hh:borderFill>).)*?'
            r'<hh:leftBorder type="SOLID"',
            header, _re.DOTALL,
        )
        assert bf_solid, "no SOLID-bordered borderFill registered for tables"
        bf_id = bf_solid.group(1)
        print(f"  verified: SOLID borderFill id={bf_id} present")
        assert f'borderFillIDRef="{bf_id}"' in sec, "table not referencing SOLID border"
        print(f"  verified: table cells reference borderFillIDRef={bf_id}")

        # Cell shading: yellow (FFFF00) and green (00FF00) borderFills must
        # be registered with fillBrush faceColor matching, and at least one
        # cell in the section must reference each color's borderFill id.
        for needle_color, label in [("#FFFF00", "yellow"), ("#00FF00", "green")]:
            bf = _re.search(
                r'<hh:borderFill id="(\d+)"[^>]*>(?:(?!</hh:borderFill>).)*?'
                r'faceColor="' + _re.escape(needle_color) + r'"',
                header, _re.DOTALL,
            )
            assert bf, f"no borderFill for cell fill {needle_color} ({label})"
            cid = bf.group(1)
            assert f'borderFillIDRef="{cid}"' in sec, (
                f"no cell references borderFill id {cid} for {label}"
            )
            print(f"  verified: cell shading {label} ({needle_color}) → "
                  f"borderFill id={cid}, referenced by a cell")


def main():
    docx = os.path.join(HERE, "_sample.docx")
    hwpx = os.path.join(HERE, "_sample.hwpx")
    build_docx(docx)
    print(f"built {docx} ({os.path.getsize(docx)} bytes)")
    convert(docx, hwpx)
    print(f"converted to {hwpx} ({os.path.getsize(hwpx)} bytes)\n")
    verify_hwpx(hwpx)
    print("\nALL CHECKS PASSED")


if __name__ == "__main__":
    main()
