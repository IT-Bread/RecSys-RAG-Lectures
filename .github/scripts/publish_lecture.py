"""Restore the supplied PPTX byte-for-byte and export its actual slide structure."""
from __future__ import annotations

import base64
import hashlib
import io
import lzma
import posixpath
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[2]
EXPECTED_SHA256 = "17eaaa7fd01b24d2f59df72a25eebbdff3b12adae19af3d0e50c8767cdb9374d"
PPTX_NAME = "RecSys_RAG_lecture.pptx"
TXT_NAME = "RecSys_RAG_slide_outline.txt"
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def text_blocks(root: ET.Element) -> list[str]:
    blocks = []
    for shape in root.findall(".//p:sp", NS):
        paragraphs = [
            "".join(node.text or "" for node in para.findall(".//a:t", NS))
            for para in shape.findall("./p:txBody/a:p", NS)
        ]
        text = "\n".join(paragraphs).strip()
        if text:
            blocks.append(text)
    return blocks


def relationship_targets(archive: zipfile.ZipFile, part: str) -> dict[str, str]:
    folder, name = posixpath.split(part)
    rel_path = posixpath.join(folder, "_rels", name + ".rels")
    root = ET.fromstring(archive.read(rel_path))
    return {
        item.attrib["Id"]: posixpath.normpath(posixpath.join(folder, item.attrib["Target"]))
        for item in root
        if item.attrib.get("TargetMode") != "External"
    }


def main() -> None:
    encoded = (ROOT / ".github/assets/lecture.pptx.xz.b64").read_text(encoding="ascii")
    data = lzma.decompress(base64.b64decode("".join(encoded.split()), validate=True))
    actual_sha = hashlib.sha256(data).hexdigest()
    if actual_sha != EXPECTED_SHA256:
        raise ValueError(f"Presentation checksum mismatch: {actual_sha}")

    sections = []
    titles = []
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        invalid_member = archive.testzip()
        if invalid_member is not None:
            raise ValueError(f"Corrupt PPTX member: {invalid_member}")
        presentation = ET.fromstring(archive.read("ppt/presentation.xml"))
        targets = relationship_targets(archive, "ppt/presentation.xml")
        slide_paths = [
            targets[item.attrib[f"{{{NS['r']}}}id"]]
            for item in presentation.findall("./p:sldIdLst/p:sldId", NS)
        ]
        if len(slide_paths) != 31:
            raise ValueError(f"Expected 31 slides, found {len(slide_paths)}")

        for number, slide_path in enumerate(slide_paths, 1):
            blocks = text_blocks(ET.fromstring(archive.read(slide_path)))
            title = blocks[1]
            subtitle = blocks[2]
            titles.append(title)
            example_marker = blocks.index("ПРИМЕР / ВЫВОД")
            content = blocks[3:9]
            diagram = blocks[9:example_marker]
            example = blocks[example_marker + 1]
            if len(content) != 6:
                raise ValueError(f"Unexpected content layout on slide {number}")

            notes = []
            for target in relationship_targets(archive, slide_path).values():
                if "/notesSlides/" not in target:
                    continue
                notes_root = ET.fromstring(archive.read(target))
                for shape in notes_root.findall(".//p:sp", NS):
                    placeholder = shape.find("./p:nvSpPr/p:nvPr/p:ph", NS)
                    if placeholder is None or placeholder.get("type") != "body":
                        continue
                    notes.extend(text_blocks(shape) or [
                        "\n".join(
                            "".join(node.text or "" for node in para.findall(".//a:t", NS))
                            for para in shape.findall("./p:txBody/a:p", NS)
                        ).strip()
                    ])

            lines = ["=" * 78, f"СЛАЙД {number:02d}. {title}", "=" * 78,
                     "", "ПОДЗАГОЛОВОК", subtitle, "", "СОДЕРЖАНИЕ СЛАЙДА"]
            for index in range(0, len(content), 2):
                lines.extend(["", content[index], content[index + 1]])
            lines.extend(["", "СХЕМА / ПОДПИСИ ИЛЛЮСТРАЦИИ", *diagram,
                          "", "ПРИМЕР / ВЫВОД", example,
                          "", "ЗАМЕТКИ ПРЕПОДАВАТЕЛЯ ИЗ ПРЕЗЕНТАЦИИ", *notes])
            sections.append("\n".join(lines))

    heading = (
        "РЕКОМЕНДАТЕЛЬНЫЕ СИСТЕМЫ И RAG: МОДЕЛИ, МЕТРИКИ И БЕЙЗЛАЙНЫ\n"
        "Текстовая схема презентации: 31 слайд\n\n"
        "Источник: приложенная презентация. Содержание слайдов, подписи схем и заметки\n"
        "перенесены без содержательных изменений. Графические объекты остаются в PPTX.\n\n"
    )
    (ROOT / PPTX_NAME).write_bytes(data)
    (ROOT / TXT_NAME).write_text(heading + "\n\n".join(sections) + "\n", encoding="utf-8")

    readme = (
        "# RecSys и RAG: модели, метрики и бейзлайны\n\n"
        "Учебная презентация для инженерных направлений: **31 слайд**.\n\n"
        "## Материалы\n\n"
        f"- [Скачать презентацию PowerPoint]({PPTX_NAME}?raw=true)\n"
        f"- [Открыть текстовую схему слайдов]({TXT_NAME})\n"
        f"- [Скачать текстовую схему]({TXT_NAME}?raw=true)\n\n"
        "В текстовом файле для каждого слайда приведены заголовок, подзаголовок, "
        "содержание, подписи схем, пример или вывод и заметки преподавателя из презентации.\n\n"
        "## Содержание\n\n"
        + "\n".join(f"{i}. {title}" for i, title in enumerate(titles, 1))
        + "\n\n## Исходный файл\n\n"
        "`RecSys_RAG_инженерная_лекция_30_слайдов.pptx` — исходное название; "
        "фактически в файле 31 слайд. Для удобства ссылок файл переименован, "
        "его содержимое не изменено.\n\n"
        f"SHA-256: `{EXPECTED_SHA256}`.\n"
    )
    (ROOT / "README.md").write_text(readme, encoding="utf-8")
    print(f"Verified and published {len(titles)} slides; SHA-256 {actual_sha}")


if __name__ == "__main__":
    main()
