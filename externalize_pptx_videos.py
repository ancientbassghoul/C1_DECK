from __future__ import annotations

import argparse
from pathlib import Path, PurePosixPath
import posixpath
import zipfile

from lxml import etree


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
VIDEO_REL = f"{R_NS}/video"
MEDIA_REL = "http://schemas.microsoft.com/office/2007/relationships/media"


def externalize_slide_video(
    entries: dict[str, bytes], slide_number: int, video_path: Path
) -> set[str]:
    rels_name = f"ppt/slides/_rels/slide{slide_number}.xml.rels"
    slide_name = f"ppt/slides/slide{slide_number}.xml"
    rels_root = etree.fromstring(entries[rels_name])
    external_uri = video_path.resolve().as_uri()
    removed_media = set()
    media_rel_ids = []

    for relationship in rels_root.findall(f"{{{REL_NS}}}Relationship"):
        if relationship.get("Type") not in (VIDEO_REL, MEDIA_REL):
            continue
        old_target = relationship.get("Target", "")
        if old_target.startswith("../media/"):
            removed_media.add(posixpath.normpath(f"ppt/slides/{old_target}"))
        relationship.set("Target", external_uri)
        relationship.set("TargetMode", "External")
        if relationship.get("Type") == MEDIA_REL:
            media_rel_ids.append(relationship.get("Id"))

    slide_root = etree.fromstring(entries[slide_name])
    for media in slide_root.findall(f".//{{{P14_NS}}}media"):
        embedded_id = media.attrib.pop(f"{{{R_NS}}}embed", None)
        if embedded_id in media_rel_ids:
            media.set(f"{{{R_NS}}}link", embedded_id)

    entries[rels_name] = etree.tostring(
        rels_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    entries[slide_name] = etree.tostring(
        slide_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    return {str(PurePosixPath(name)) for name in removed_media}


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Replace embedded Slide 8/10 videos with external links."
    )
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--slide-8-video", required=True, type=Path)
    parser.add_argument("--slide-10-video", required=True, type=Path)
    args = parser.parse_args()

    for path in (args.input, args.slide_8_video, args.slide_10_video):
        if not path.exists():
            raise FileNotFoundError(path)
    if args.output.resolve() == args.input.resolve():
        raise ValueError("Output must be a separate file.")

    with zipfile.ZipFile(args.input, "r") as source:
        infos = source.infolist()
        entries = {info.filename: source.read(info.filename) for info in infos}

    removed_media = set()
    removed_media |= externalize_slide_video(entries, 8, args.slide_8_video)
    removed_media |= externalize_slide_video(entries, 10, args.slide_10_video)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(args.output, "w", allowZip64=True) as target:
        for info in infos:
            if info.filename in removed_media:
                continue
            target.writestr(info, entries[info.filename])

    print(args.output)
    print("Removed embedded media:", ", ".join(sorted(removed_media)))


if __name__ == "__main__":
    main()
