from __future__ import annotations

import argparse
import posixpath
from pathlib import Path, PurePosixPath
from urllib.parse import quote
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
VIDEO_REL = f"{R_NS}/video"
MEDIA_REL = "http://schemas.microsoft.com/office/2007/relationships/media"

DEFAULT_LINKS = {
    5: Path("visuals/Raycast_Slide_8_Visual_RETIMED.mp4"),
    7: Path("visuals/Raycast_Slide_10_Visual.mp4"),
    17: Path("visuals/Raycast_Slide_35_Visual.mp4"),
}


def externalize_slide_video(
    entries: dict[str, bytes], slide_number: int, relative_video_path: Path
) -> set[str]:
    rels_name = f"ppt/slides/_rels/slide{slide_number}.xml.rels"
    slide_name = f"ppt/slides/slide{slide_number}.xml"
    rels_root = etree.fromstring(entries[rels_name])
    target = quote(relative_video_path.as_posix(), safe="/._-")
    removed_media: set[str] = set()
    media_rel_ids: set[str] = set()

    matched = 0
    for relationship in rels_root.findall(f"{{{REL_NS}}}Relationship"):
        if relationship.get("Type") not in (VIDEO_REL, MEDIA_REL):
            continue
        matched += 1
        old_target = relationship.get("Target", "")
        if old_target.startswith("../media/"):
            removed_media.add(posixpath.normpath(f"ppt/slides/{old_target}"))
        relationship.set("Target", target)
        relationship.set("TargetMode", "External")
        if relationship.get("Type") == MEDIA_REL:
            media_rel_ids.add(relationship.get("Id", ""))

    if matched != 2 or len(media_rel_ids) != 1:
        raise ValueError(
            f"Slide {slide_number} should have one media and one video relationship; "
            f"found {matched} linked-media relationships."
        )

    slide_root = etree.fromstring(entries[slide_name])
    converted = 0
    for media in slide_root.findall(f".//{{{P14_NS}}}media"):
        embedded_id = media.attrib.pop(f"{{{R_NS}}}embed", None)
        if embedded_id in media_rel_ids:
            media.set(f"{{{R_NS}}}link", embedded_id)
            converted += 1
    if converted != 1:
        raise ValueError(
            f"Slide {slide_number} should have one p14:media embed reference; "
            f"converted {converted}."
        )

    entries[rels_name] = etree.tostring(
        rels_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    entries[slide_name] = etree.tostring(
        slide_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    return {str(PurePosixPath(name)) for name in removed_media}


def build_linked_deck(input_path: Path, output_path: Path) -> None:
    input_path = input_path.resolve()
    output_path = output_path.resolve()
    if input_path == output_path:
        raise ValueError("Output must be a separate file; the embedded deck is read-only.")
    if not input_path.exists():
        raise FileNotFoundError(input_path)

    project_root = input_path.parent
    for relative_path in DEFAULT_LINKS.values():
        media_path = project_root / relative_path
        if not media_path.exists():
            raise FileNotFoundError(media_path)

    with ZipFile(input_path, "r") as source:
        infos = source.infolist()
        entries = {info.filename: source.read(info.filename) for info in infos}

    removed_media: set[str] = set()
    for slide_number, relative_path in DEFAULT_LINKS.items():
        removed_media |= externalize_slide_video(entries, slide_number, relative_path)

    if len(removed_media) != len(DEFAULT_LINKS):
        raise ValueError(
            f"Expected {len(DEFAULT_LINKS)} embedded video parts, found {len(removed_media)}."
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(output_path, "w", compression=ZIP_DEFLATED, allowZip64=True) as target:
        for info in infos:
            if info.filename in removed_media:
                continue
            target.writestr(info, entries[info.filename])

    print(f"Wrote {output_path}")
    for slide_number, relative_path in DEFAULT_LINKS.items():
        print(f"Slide {slide_number}: {relative_path.as_posix()}")
    print("Removed embedded media: " + ", ".join(sorted(removed_media)))


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Create a linked-media copy of the approved Raycast v3 deck."
    )
    parser.add_argument(
        "--input", type=Path, default=Path("presentation_v3.pptx")
    )
    parser.add_argument(
        "--output", type=Path, default=Path("presentation_v3_linked.pptx")
    )
    args = parser.parse_args()
    build_linked_deck(args.input, args.output)


if __name__ == "__main__":
    main()
