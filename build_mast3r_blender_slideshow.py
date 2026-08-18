from __future__ import annotations

import argparse
from pathlib import Path
import subprocess


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"}


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Build a modification-date-ordered H.265 image slideshow."
    )
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()

    source = args.source.resolve()
    output = args.output.resolve()
    images = sorted(
        (
            path
            for path in source.iterdir()
            if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS
        ),
        key=lambda path: (path.stat().st_mtime_ns, path.name.casefold()),
    )
    if not images:
        raise ValueError(f"No images found in {source}")

    output.parent.mkdir(parents=True, exist_ok=True)
    print(f"Encoding {len(images)} images in modification-date order:", flush=True)
    for index, image in enumerate(images, start=1):
        timestamp = image.stat().st_mtime
        print(f"{index:02d}. {timestamp:.6f}  {image.name}", flush=True)

    command = [
        "ffmpeg",
        "-hide_banner",
        "-y",
    ]
    for image in images:
        command.extend(
            ["-loop", "1", "-framerate", "30", "-t", "2", "-i", str(image)]
        )

    filters = []
    for index in range(len(images)):
        filters.append(
            f"[{index}:v]"
            "scale=1920:1080:force_original_aspect_ratio=decrease:flags=lanczos,"
            "pad=1920:1080:(ow-iw)/2:(oh-ih)/2:color=black,"
            "setsar=1,fps=30,format=yuv420p,trim=duration=2,setpts=PTS-STARTPTS"
            f"[v{index}]"
        )
    concat_inputs = "".join(f"[v{index}]" for index in range(len(images)))
    filters.append(f"{concat_inputs}concat=n={len(images)}:v=1:a=0[outv]")
    command.extend(
        [
        "-filter_complex",
        ";".join(filters),
        "-map",
        "[outv]",
        "-an",
        "-c:v",
        "libx265",
        "-preset",
        "slow",
        "-crf",
        "14",
        "-tag:v",
        "hvc1",
        "-movflags",
        "+faststart",
        str(output),
        ]
    )
    subprocess.run(command, check=True)

    print(output, flush=True)


if __name__ == "__main__":
    main()
