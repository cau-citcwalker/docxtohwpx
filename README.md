# docx2hwpx

DOCX 파일을 한글(HWPX)로 변환하는 단일 파일 Python 컨버터입니다. 외부 의존성 없이 표준 라이브러리만 사용합니다.

## 빠른 시작

```bash
python3 docx2hwpx.py input.docx output.hwpx
```

저장소를 클론한 디렉토리에서 바로 실행 가능합니다. `skeleton.hwpx`(템플릿)이 같은 폴더에 있어야 합니다.

테스트를 실행하려면:

```bash
python3 test_convert.py
```

가상의 DOCX를 만들어 변환하고, 결과 HWPX의 ZIP 구조와 콘텐츠를 검증합니다.

## 동작 방식

1. **DOCX 파싱** (`DocxReader`) — `word/document.xml`, 스타일, 번호 매기기, 각주/미주, 머리/꼬리말, 관계 파일을 모두 읽어 중간 표현(IR: `Document`/`Paragraph`/`Run`/`Table` 등)으로 변환합니다.
2. **속성 인터닝** — IR을 순회하며 고유한 글자/단락 속성, 폰트, 셀 배경색을 모아 ID를 부여합니다. HWPX는 모든 속성이 헤더에 등록된 ID로만 참조되기 때문입니다.
3. **HWPX 생성** (`HwpxWriter`) — 검증된 빈 HWPX 템플릿(`skeleton.hwpx`)을 그대로 쓰면서, 새 charPr/paraPr/font/borderFill을 헤더에 덧붙이고 `Contents/section0.xml`을 새로 작성합니다. 이미지 같은 바이너리는 `BinData/`에 넣고 OPF 매니페스트(`Contents/content.hpf`)에 등록합니다.

처음부터 HWPX를 직접 만드는 대신 **템플릿을 패치**하는 방식을 쓴 이유: HWPX(OWPML 표준)는 스키마가 매우 엄격해서 손으로 짜면 한글이 거부합니다 (속성 순서, flag 요소의 존재 유무, 네임스페이스 prefix 등 수십 가지 규칙). 한컴이 공개한 [OWPML 모델](https://github.com/hancom-io/hwpx-owpml-model)과 [python-hwpx](https://github.com/airmang/python-hwpx)에서 가져온 정상 동작 skeleton을 베이스로 깔고, 본문(section)만 동적으로 생성하면 안전합니다.

### 단위 변환

| DOCX                | HWPX                  | 변환식      |
| ------------------- | --------------------- | ----------- |
| twip (1/1440 inch)  | HWPUNIT (1/7200 inch) | `× 5`       |
| EMU (1/914400 inch) | HWPUNIT               | `// 127`    |
| half-point          | HWPX 글자 크기 (1/100pt) | `× 50`      |

## 지원 기능

- 텍스트, 단락 정렬 (LEFT / CENTER / RIGHT / JUSTIFY / DISTRIBUTE)
- 글자 굵게 / 기울임 / 밑줄 / 취소선 / 글자 크기 / 색 / 폰트
- 표 (가로 colSpan + 세로 rowSpan via `<w:vMerge>`)
- 표 셀 배경색 (`<w:shd>` → 색상별 borderFill 자동 등록)
- 이미지 (PNG / JPG / GIF / BMP / WMF / EMF / TIFF)
- 하이퍼링크 (HYPERLINK 필드)
- 각주 / 미주
- 십진 자동 번호 (DOCX 번호 목록 → HWPX OUTLINE 헤딩)
- 페이지 나눔 (`<w:br w:type="page"/>`, `<w:pageBreakBefore/>`)
- 트랙 변경 수락 (`<w:ins>` 포함, `<w:del>` 제거)

## 알려진 한계

- **수식**: DOCX의 OOXML Math와 HWPX 수식은 별개의 문법이라 자동 매핑 불가능. 텍스트만 추출합니다.
- **머리말 / 꼬리말**: 템플릿 단순화 과정에서 빠져 있습니다. 필요하면 추가 가능.
- **불릿 자동 기호**: 들여쓰기는 보존되지만 ●/■ 같은 자동 기호는 출력되지 않습니다.
- **차트 / SmartArt / VML 도형**: 무시됩니다.
- **다중 섹션 머리/꼬리말**: 첫 번째만 사용합니다.
- **트랙 변경 자체 보존**: 변경을 모두 수락하는 동작만 지원합니다.

## 파일 구성

| 파일 | 설명 |
| --- | --- |
| `docx2hwpx.py` | 컨버터 본체 — IR, DocxReader, HwpxWriter, CLI |
| `skeleton.hwpx` | 한글에서 정상 열리는 빈 HWPX 템플릿. 컨버터가 이 위에 덧대 새 파일을 만듭니다 |
| `test_convert.py` | 스모크 테스트. 가상 DOCX로 모든 기능 검증 |

## 코드 구조 요약

핵심 클래스 위치:

| 항목 | 위치 |
| --- | --- |
| 단위 변환 | [`twip_to_hwp` / `emu_to_hwp` / `halfpt_to_hwp_size`](docx2hwpx.py) |
| IR 데이터클래스 | `CharProps`, `ParaProps`, `Run`, `Hyperlink`, `Image`, `FootnoteRef`, `PageBreak`, `Paragraph`, `TableCell`, `Table`, `Document` |
| DOCX 파싱 | `DocxReader` |
| 헤더 패치 (charPr/paraPr/borderFill 추가) | `HwpxWriter._patch_header`, `_patch_charpr`, `_patch_parapr`, `_patch_borderfills` |
| 본문 생성 | `HwpxWriter._build_section`, `_paragraph_xml`, `_table_ctl_xml` |
| 셀 배경색 인터닝 | `HwpxWriter._intern_cell_fill` |
| 페이지 나눔 단락 분할 | `HwpxWriter._paragraph_xml` (PageBreak 발견 시 `pageBreak="1"` 단락으로 split) |

## 라이선스

MIT License — 자세한 내용은 [LICENSE](LICENSE) 파일을 참고하세요.

## 참고 자료

HWPX 형식 분석에 사용한 외부 자료:

- [한컴 OWPML 모델 (C++ 소스)](https://github.com/hancom-io/hwpx-owpml-model)
- [airmang/python-hwpx (Skeleton 및 fixture 제공)](https://github.com/airmang/python-hwpx)
- [한컴테크 — HWPX 포맷 구조](https://tech.hancom.com/hwpxformat/)
