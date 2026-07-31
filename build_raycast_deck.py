from pathlib import Path
import argparse
import subprocess

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_VISUALS_DIR = SCRIPT_DIR / "visuals"
DEFAULT_DATASET_DIR = SCRIPT_DIR / "dataset"
DEFAULT_MASCOT_PATH = (
    SCRIPT_DIR / "assets" / "Time_Traveler_Spoilers_Flattened_Transparent.png"
)
DEFAULT_SERIOUS_MASCOT_PATH = (
    SCRIPT_DIR / "assets" / "Time_Traveler_Spoilers_Red_Portal_Transparent.png"
)
ANIMATION_SCRIPT = SCRIPT_DIR / "apply_speech_animations.ps1"
OUTPUT_PPTX = SCRIPT_DIR / "presentation.pptx"

SPEECH_OVERLAY_SLIDES = (1, 4, 5, 10, 11)
SPEECH_TEXT_MARKERS = {
    1: "≤10 pixels",
    4: "First commit after",
    5: "You have no idea how important",
    10: "This exact commit also complains",
    11: "Yeah, so much effort",
}
MASCOT_X = 11.75
MASCOT_Y = 5.85
MASCOT_SIZE = 1.48
MASCOT_ORIGIN_X = MASCOT_X + MASCOT_SIZE / 2
MASCOT_ORIGIN_Y = MASCOT_Y + MASCOT_SIZE / 2
SPEECH_BUBBLE_RIGHT = 11.86
SPEECH_TAIL_TARGET_Y = MASCOT_Y + 0.08


def find_asset(assets_dir, *names):
    """Return the first matching asset from a project-local asset directory."""
    for name in names:
        exact = assets_dir / name
        if exact.exists():
            return exact
        matches = sorted(assets_dir.glob(f"*{name}"))
        if matches:
            return matches[0]
    raise FileNotFoundError(
        f"Could not find any of {names!r} in {assets_dir.resolve()}."
    )

SLIDE_W = 13.333
SLIDE_H = 7.5

PRIMARY = "101827"
SECONDARY = "27E38D"
ACCENT = "FF5B5B"
OFFWHITE = "F5F2EA"
WHITE = "FFFFFF"
MUTED = "8A94A4"
PALE = "DCE2EA"
DEEP = "09111E"
SOFT_GREEN = "DDF8EC"

HEADER_FONT = "Bahnschrift SemiBold"
BODY_FONT = "Aptos"

ACT_LABELS = [
    "Setup",
    "Bring-Up",
    "By Hand",
    "Check-In",
    "Pivot",
    "New Brain",
    "Bug Hunt",
    "Silent Bugs",
    "Tuning",
    "Twist",
    "Landing",
]

ACT_SPECS = [
    {"number": "0", "label": "Setup", "header": "THE SETUP", "start": 1, "end": 3},
    {"number": "I", "label": "Bring-Up", "header": "BRING-UP", "start": 4, "end": 6},
    {
        "number": "II",
        "label": "By Hand",
        "header": "BY HAND,\nFOR NOW",
        "start": 7,
        "end": 11,
    },
    {
        "number": "III",
        "label": "Check-In",
        "header": "ASKING FOR\nDIRECTIONS",
        "start": 12,
        "end": 18,
    },
    {
        "number": "IV",
        "label": "Pivot",
        "header": "STARTING OVER",
        "start": 19,
        "end": 23,
    },
    {
        "number": "V",
        "label": "New Brain",
        "header": "HOW THE NEW\nBRAIN THINKS",
        "start": 24,
        "end": 25,
    },
    {
        "number": "VI",
        "label": "Bug Hunt",
        "header": "THE BUG HUNT",
        "start": 26,
        "end": 29,
    },
    {
        "number": "VII",
        "label": "Silent Bugs",
        "header": "SILENT FAILURES",
        "start": 30,
        "end": 32,
    },
    {
        "number": "VIII",
        "label": "Tuning",
        "header": "TUNING THE\nMACHINE",
        "start": 33,
        "end": 36,
    },
    {
        "number": "IX",
        "label": "Twist",
        "header": "THE TWIST",
        "start": 37,
        "end": 38,
    },
    {
        "number": "X",
        "label": "Landing",
        "header": "WHERE IT LANDED",
        "start": 39,
        "end": 40,
    },
]


def rgb(hex_color):
    return RGBColor.from_string(hex_color)


def set_background(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(color)


def add_rect(slide, x, y, w, h, fill, line=None, radius=True, line_width=1):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(
        shape_type, Inches(x), Inches(y), Inches(w), Inches(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = rgb(line)
        shape.line.width = Pt(line_width)
    return shape


def add_text(
    slide,
    text,
    x,
    y,
    w,
    h,
    font_size,
    color,
    font=BODY_FONT,
    bold=False,
    align=PP_ALIGN.LEFT,
    valign=MSO_ANCHOR.TOP,
    margin=0.0,
    italic=False,
):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin)
    tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = valign
    paragraph = tf.paragraphs[0]
    paragraph.alignment = align
    paragraph.space_before = Pt(0)
    paragraph.space_after = Pt(0)
    paragraph.line_spacing = 1.0
    run = paragraph.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = rgb(color)
    return box


def add_rich_text(
    slide,
    segments,
    x,
    y,
    w,
    h,
    font_size,
    color,
    valign=MSO_ANCHOR.TOP,
    margin=0.0,
):
    """Add one text item with per-segment font choices."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.vertical_anchor = valign
    frame.margin_left = Inches(margin)
    frame.margin_right = Inches(margin)
    frame.margin_top = Inches(margin)
    frame.margin_bottom = Inches(margin)
    paragraph = frame.paragraphs[0]
    paragraph.space_before = Pt(0)
    paragraph.space_after = Pt(0)
    for text, font in segments:
        run = paragraph.add_run()
        run.text = text
        run.font.name = font
        run.font.size = Pt(font_size)
        run.font.color.rgb = rgb(color)
    return box


def add_title(slide, title, dark=False, font_size=40):
    color = WHITE if dark else PRIMARY
    add_text(
        slide,
        title,
        1.75,
        0.36,
        11.1,
        0.62,
        font_size,
        color,
        font=HEADER_FONT,
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
    )
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(1.75), Inches(1.10), Inches(0.78), Inches(0.055)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = rgb(SECONDARY)
    line.line.fill.background()


def add_act_rail(slide, current_slide):
    rail = add_rect(slide, 0, 0, 1.42, SLIDE_H, PRIMARY, None, radius=False)

    add_text(
        slide,
        "ACT 0 —\nTHE SETUP",
        0.18,
        0.18,
        1.08,
        0.78,
        13.5,
        WHITE,
        font=HEADER_FONT,
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
    )
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.18), Inches(1.02), Inches(0.48), Inches(0.045)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = rgb(SECONDARY)
    accent.line.fill.background()

    current_y = 1.17
    add_rect(slide, 0.10, current_y, 1.20, 1.08, DEEP, SECONDARY, radius=True)
    add_text(
        slide,
        "Setup",
        0.55,
        current_y + 0.37,
        0.66,
        0.28,
        10.5,
        SECONDARY,
        font=HEADER_FONT,
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
    )

    dot_x = 0.32
    dot_ys = [current_y + 0.24, current_y + 0.53, current_y + 0.82]
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(dot_x - 0.018),
        Inches(dot_ys[0]),
        Inches(0.036),
        Inches(dot_ys[-1] - dot_ys[0]),
    )
    line.fill.solid()
    line.fill.fore_color.rgb = rgb(MUTED)
    line.line.fill.background()

    for index, dot_y in enumerate(dot_ys, start=1):
        dot = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(dot_x - 0.07),
            Inches(dot_y - 0.07),
            Inches(0.14),
            Inches(0.14),
        )
        if index < current_slide:
            dot.fill.solid()
            dot.fill.fore_color.rgb = rgb(MUTED)
            dot.line.color.rgb = rgb(MUTED)
        elif index == current_slide:
            dot.fill.solid()
            dot.fill.fore_color.rgb = rgb(SECONDARY)
            dot.line.color.rgb = rgb(SECONDARY)
        else:
            dot.fill.solid()
            dot.fill.fore_color.rgb = rgb(PRIMARY)
            dot.line.color.rgb = rgb(PALE)
        dot.line.width = Pt(1.25)

    row_y = 2.39
    row_h = 0.43
    for label in ACT_LABELS[1:]:
        tick = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0.25),
            Inches(row_y + 0.19),
            Inches(0.12),
            Inches(0.018),
        )
        tick.fill.solid()
        tick.fill.fore_color.rgb = rgb(MUTED)
        tick.line.fill.background()
        add_text(
            slide,
            label,
            0.47,
            row_y + 0.06,
            0.78,
            0.28,
            8.3,
            MUTED,
            font=BODY_FONT,
            valign=MSO_ANCHOR.MIDDLE,
        )
        row_y += row_h

    return rail


def add_act_i_rail(slide, current_slide):
    """Add the persistent rail for Act I without altering the Act 0 rail."""
    rail = add_rect(slide, 0, 0, 1.42, SLIDE_H, PRIMARY, None, radius=False)

    add_text(
        slide,
        "ACT I —\nBRING-UP",
        0.18,
        0.18,
        1.08,
        0.78,
        13.5,
        WHITE,
        font=HEADER_FONT,
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
    )
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.18), Inches(1.02), Inches(0.48), Inches(0.045)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = rgb(SECONDARY)
    accent.line.fill.background()

    current_y = 1.17
    add_rect(slide, 0.10, current_y, 1.20, 1.08, DEEP, SECONDARY, radius=True)
    add_text(
        slide,
        "Bring-Up",
        0.55,
        current_y + 0.37,
        0.66,
        0.28,
        10.5,
        SECONDARY,
        font=HEADER_FONT,
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
    )

    dot_x = 0.32
    dot_ys = [current_y + 0.24, current_y + 0.53, current_y + 0.82]
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(dot_x - 0.018),
        Inches(dot_ys[0]),
        Inches(0.036),
        Inches(dot_ys[-1] - dot_ys[0]),
    )
    line.fill.solid()
    line.fill.fore_color.rgb = rgb(MUTED)
    line.line.fill.background()

    for index, dot_y in enumerate(dot_ys, start=4):
        dot = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(dot_x - 0.07),
            Inches(dot_y - 0.07),
            Inches(0.14),
            Inches(0.14),
        )
        if index < current_slide:
            dot.fill.solid()
            dot.fill.fore_color.rgb = rgb(MUTED)
            dot.line.color.rgb = rgb(MUTED)
        elif index == current_slide:
            dot.fill.solid()
            dot.fill.fore_color.rgb = rgb(SECONDARY)
            dot.line.color.rgb = rgb(SECONDARY)
        else:
            dot.fill.solid()
            dot.fill.fore_color.rgb = rgb(PRIMARY)
            dot.line.color.rgb = rgb(PALE)
        dot.line.width = Pt(1.25)

    other_acts = [(0, ACT_LABELS[0])] + list(enumerate(ACT_LABELS[2:], start=2))
    row_y = 2.39
    row_h = 0.43
    for act_index, label in other_acts:
        tick = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0.25),
            Inches(row_y + 0.17),
            Inches(0.12),
            Inches(0.08),
        )
        if act_index < 1:
            tick.fill.solid()
            tick.fill.fore_color.rgb = rgb(MUTED)
            tick.line.color.rgb = rgb(MUTED)
        else:
            tick.fill.solid()
            tick.fill.fore_color.rgb = rgb(PRIMARY)
            tick.line.color.rgb = rgb(MUTED)
        tick.line.width = Pt(0.8)
        add_text(
            slide,
            label,
            0.47,
            row_y + 0.06,
            0.78,
            0.28,
            8.3,
            MUTED,
            font=BODY_FONT,
            valign=MSO_ANCHOR.MIDDLE,
        )
        row_y += row_h

    return rail


def add_act_ii_rail(slide, current_slide):
    """Add the persistent rail and five stations for Act II."""
    rail = add_rect(slide, 0, 0, 1.42, SLIDE_H, PRIMARY, None, radius=False)

    add_text(
        slide,
        "ACT II —\nBY HAND,\nFOR NOW",
        0.18,
        0.13,
        1.08,
        0.88,
        11.8,
        WHITE,
        font=HEADER_FONT,
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
    )
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.18), Inches(1.02), Inches(0.48), Inches(0.045)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = rgb(SECONDARY)
    accent.line.fill.background()

    current_y = 1.17
    add_rect(slide, 0.10, current_y, 1.20, 1.48, DEEP, SECONDARY, radius=True)
    add_text(
        slide,
        "By Hand",
        0.55,
        current_y + 0.57,
        0.66,
        0.30,
        10.5,
        SECONDARY,
        font=HEADER_FONT,
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
    )

    dot_x = 0.32
    dot_ys = [current_y + 0.24 + 0.25 * index for index in range(5)]
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(dot_x - 0.018),
        Inches(dot_ys[0]),
        Inches(0.036),
        Inches(dot_ys[-1] - dot_ys[0]),
    )
    line.fill.solid()
    line.fill.fore_color.rgb = rgb(MUTED)
    line.line.fill.background()

    for index, dot_y in enumerate(dot_ys, start=7):
        dot = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(dot_x - 0.07),
            Inches(dot_y - 0.07),
            Inches(0.14),
            Inches(0.14),
        )
        if index < current_slide:
            dot.fill.solid()
            dot.fill.fore_color.rgb = rgb(MUTED)
            dot.line.color.rgb = rgb(MUTED)
        elif index == current_slide:
            dot.fill.solid()
            dot.fill.fore_color.rgb = rgb(SECONDARY)
            dot.line.color.rgb = rgb(SECONDARY)
        else:
            dot.fill.solid()
            dot.fill.fore_color.rgb = rgb(PRIMARY)
            dot.line.color.rgb = rgb(PALE)
        dot.line.width = Pt(1.25)

    other_acts = [
        (act_index, label)
        for act_index, label in enumerate(ACT_LABELS)
        if act_index != 2
    ]
    row_y = 2.78
    row_h = 0.40
    for act_index, label in other_acts:
        tick = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0.25),
            Inches(row_y + 0.16),
            Inches(0.12),
            Inches(0.08),
        )
        if act_index < 2:
            tick.fill.solid()
            tick.fill.fore_color.rgb = rgb(MUTED)
            tick.line.color.rgb = rgb(MUTED)
        else:
            tick.fill.solid()
            tick.fill.fore_color.rgb = rgb(PRIMARY)
            tick.line.color.rgb = rgb(MUTED)
        tick.line.width = Pt(0.8)
        add_text(
            slide,
            label,
            0.47,
            row_y + 0.05,
            0.78,
            0.28,
            8.1,
            MUTED,
            font=BODY_FONT,
            valign=MSO_ANCHOR.MIDDLE,
        )
        row_y += row_h

    return rail


def get_act_spec(slide_number):
    for act_index, spec in enumerate(ACT_SPECS):
        if spec["start"] <= slide_number <= spec["end"]:
            return act_index, spec
    raise ValueError(f"Slide {slide_number} is outside the configured 40-slide deck.")


def add_inline_act_rail(slide, current_slide):
    """Render the active act as an inline accordion in chronological order."""
    active_index, active_spec = get_act_spec(current_slide)
    rail = add_rect(slide, 0, 0, 1.42, SLIDE_H, PRIMARY, None, radius=False)

    header_text = f'ACT {active_spec["number"]} —\n{active_spec["header"]}'
    header_lines = header_text.count("\n") + 1
    header_font_size = 13.0 if header_lines == 2 else 11.4
    add_text(
        slide,
        header_text,
        0.18,
        0.12,
        1.08,
        0.90,
        header_font_size,
        WHITE,
        font=HEADER_FONT,
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
    )
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.18), Inches(1.02), Inches(0.48), Inches(0.045)
    )
    accent.fill.solid()
    accent.fill.fore_color.rgb = rgb(SECONDARY)
    accent.line.fill.background()

    list_y = 1.18
    collapsed_h = 0.37
    station_step = 0.20
    station_count = active_spec["end"] - active_spec["start"] + 1
    active_h = collapsed_h + station_count * station_step + 0.13

    for act_index, spec in enumerate(ACT_SPECS):
        if act_index != active_index:
            marker = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(0.25),
                Inches(list_y + 0.145),
                Inches(0.12),
                Inches(0.08),
            )
            if act_index < active_index:
                marker.fill.solid()
                marker.fill.fore_color.rgb = rgb(MUTED)
                marker.line.color.rgb = rgb(MUTED)
            else:
                marker.fill.solid()
                marker.fill.fore_color.rgb = rgb(PRIMARY)
                marker.line.color.rgb = rgb(MUTED)
            marker.line.width = Pt(0.8)
            add_text(
                slide,
                spec["label"],
                0.47,
                list_y + 0.045,
                0.78,
                0.28,
                8.2,
                MUTED,
                font=BODY_FONT,
                valign=MSO_ANCHOR.MIDDLE,
            )
            list_y += collapsed_h
            continue

        add_rect(
            slide,
            0.10,
            list_y,
            1.20,
            active_h,
            DEEP,
            SECONDARY,
            radius=True,
            line_width=1.0,
        )
        active_marker = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0.25),
            Inches(list_y + 0.145),
            Inches(0.12),
            Inches(0.08),
        )
        active_marker.fill.solid()
        active_marker.fill.fore_color.rgb = rgb(SECONDARY)
        active_marker.line.color.rgb = rgb(SECONDARY)
        add_text(
            slide,
            spec["label"],
            0.47,
            list_y + 0.045,
            0.78,
            0.28,
            9.0,
            SECONDARY,
            font=HEADER_FONT,
            bold=True,
            valign=MSO_ANCHOR.MIDDLE,
        )

        dot_x = 0.31
        first_dot_y = list_y + collapsed_h + 0.07
        dot_ys = [first_dot_y + station_step * index for index in range(station_count)]
        if station_count > 1:
            station_line = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                Inches(dot_x - 0.014),
                Inches(dot_ys[0]),
                Inches(0.028),
                Inches(dot_ys[-1] - dot_ys[0]),
            )
            station_line.fill.solid()
            station_line.fill.fore_color.rgb = rgb(MUTED)
            station_line.line.fill.background()

        current_station = current_slide - spec["start"]
        for station_index, dot_y in enumerate(dot_ys):
            dot = slide.shapes.add_shape(
                MSO_SHAPE.OVAL,
                Inches(dot_x - 0.06),
                Inches(dot_y - 0.06),
                Inches(0.12),
                Inches(0.12),
            )
            if station_index < current_station:
                dot.fill.solid()
                dot.fill.fore_color.rgb = rgb(MUTED)
                dot.line.color.rgb = rgb(MUTED)
            elif station_index == current_station:
                dot.fill.solid()
                dot.fill.fore_color.rgb = rgb(SECONDARY)
                dot.line.color.rgb = rgb(SECONDARY)
            else:
                dot.fill.solid()
                dot.fill.fore_color.rgb = rgb(DEEP)
                dot.line.color.rgb = rgb(PALE)
            dot.line.width = Pt(1.1)

        list_y += active_h

    return rail


def normalize_all_act_rails(prs):
    """Replace legacy top-pinned rails while preserving all non-rail slide content."""
    rail_right = Inches(1.43)
    for slide_number, slide in enumerate(prs.slides, start=1):
        rail_shapes = [
            shape
            for shape in slide.shapes
            if shape.left < rail_right and shape.left + shape.width <= rail_right
        ]
        for shape in rail_shapes:
            shape._element.getparent().remove(shape._element)
        add_inline_act_rail(slide, slide_number)


def add_body_card(slide, number, text, x, y, w, h, dark=False, font_size=15.0):
    fill = DEEP if dark else WHITE
    border = "334155" if dark else PALE
    text_color = OFFWHITE if dark else PRIMARY
    card = add_rect(slide, x, y, w, h, fill, border, radius=True, line_width=1.2)
    add_text(
        slide,
        f"{number:02d}",
        x + 0.18,
        y + 0.15,
        0.42,
        0.24,
        10,
        SECONDARY,
        font=HEADER_FONT,
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
    )
    mark = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(x + 0.18),
        Inches(y + 0.48),
        Inches(0.34),
        Inches(0.035),
    )
    mark.fill.solid()
    mark.fill.fore_color.rgb = rgb(SECONDARY)
    mark.line.fill.background()
    add_text(
        slide,
        text,
        x + 0.67,
        y + 0.14,
        w - 0.85,
        h - 0.25,
        font_size,
        text_color,
        font=BODY_FONT,
        valign=MSO_ANCHOR.MIDDLE,
    )
    return card


def add_visual_placeholder(slide, instruction, x, y, w, h, dark=False, split=False):
    fill = "182338" if dark else "E8EDF3"
    border = "516177" if dark else MUTED
    text_color = PALE if dark else PRIMARY
    add_rect(slide, x, y, w, h, fill, border, radius=True, line_width=1.2)

    if split:
        gap = 0.12
        half = (w - gap - 0.42) / 2
        left_x = x + 0.15
        right_x = left_x + half + gap
        inner_y = y + 0.52
        inner_h = h - 1.43
        add_rect(
            slide,
            left_x,
            inner_y,
            half,
            inner_h,
            "25334A" if dark else "F8FAFC",
            "516177" if dark else PALE,
            radius=True,
        )
        add_rect(
            slide,
            right_x,
            inner_y,
            half,
            inner_h,
            "131C2D" if dark else "CBD5E1",
            "516177" if dark else PALE,
            radius=True,
        )
        add_text(
            slide,
            "SHARP FRAME",
            left_x,
            inner_y + inner_h / 2 - 0.18,
            half,
            0.36,
            14,
            SECONDARY,
            font=HEADER_FONT,
            bold=True,
            align=PP_ALIGN.CENTER,
            valign=MSO_ANCHOR.MIDDLE,
        )
        add_text(
            slide,
            "WORST-BLURRED FRAME",
            right_x,
            inner_y + inner_h / 2 - 0.18,
            half,
            0.36,
            13,
            ACCENT,
            font=HEADER_FONT,
            bold=True,
            align=PP_ALIGN.CENTER,
            valign=MSO_ANCHOR.MIDDLE,
        )
        add_text(
            slide,
            "VISUAL PLACEHOLDER",
            x + 0.18,
            y + 0.13,
            w - 0.36,
            0.28,
            11,
            SECONDARY,
            font=HEADER_FONT,
            bold=True,
            valign=MSO_ANCHOR.MIDDLE,
        )
        add_text(
            slide,
            instruction,
            x + 0.20,
            y + h - 0.72,
            w - 0.40,
            0.56,
            10.5,
            text_color,
            font=BODY_FONT,
            italic=True,
            valign=MSO_ANCHOR.MIDDLE,
        )
    else:
        add_text(
            slide,
            "VISUAL PLACEHOLDER",
            x + 0.25,
            y + 0.28,
            w - 0.50,
            0.35,
            12,
            SECONDARY,
            font=HEADER_FONT,
            bold=True,
            valign=MSO_ANCHOR.MIDDLE,
        )
        add_text(
            slide,
            instruction,
            x + 0.35,
            y + 0.92,
            w - 0.70,
            h - 1.35,
            14,
            text_color,
            font=BODY_FONT,
            italic=True,
            valign=MSO_ANCHOR.MIDDLE,
        )
    return


def add_speech_bubble(slide, text, x, y, w, h, font_size=15.5):
    bubble = add_rect(slide, x, y, w, h, ACCENT, WHITE, radius=True, line_width=1.2)
    bubble.name = "SpeechBubble_Background"
    bubble_text = add_text(
        slide,
        text,
        x + 0.22,
        y + 0.10,
        w - 0.44,
        h - 0.20,
        font_size,
        WHITE,
        font=HEADER_FONT,
        bold=True,
        align=PP_ALIGN.CENTER,
        valign=MSO_ANCHOR.MIDDLE,
    )
    bubble_text.name = "SpeechBubble_Text"
    tail = slide.shapes.add_shape(
        MSO_SHAPE.ISOSCELES_TRIANGLE,
        Inches(x + w - 0.58),
        Inches(y + h - 0.04),
        Inches(0.38),
        Inches(0.28),
    )
    tail.rotation = 180
    tail.fill.solid()
    tail.fill.fore_color.rgb = rgb(ACCENT)
    tail.line.fill.background()
    tail.name = "SpeechBubble_Tail"
    return bubble


def _shape_fill_rgb(shape):
    try:
        return str(shape.fill.fore_color.rgb)
    except (AttributeError, TypeError, ValueError):
        return None


def prepare_speech_overlays(prs, mascot_path, serious_mascot_path):
    """Name, position, and populate the five animated speech overlays."""
    if not mascot_path.exists():
        raise FileNotFoundError(f"Missing mascot asset: {mascot_path}")
    if not serious_mascot_path.exists():
        raise FileNotFoundError(
            f"Missing serious mascot asset: {serious_mascot_path}"
        )

    for slide_number in SPEECH_OVERLAY_SLIDES:
        if slide_number > len(prs.slides):
            continue
        slide = prs.slides[slide_number - 1]
        if any(shape.name == "SpeechOverlay_Group" for shape in slide.shapes):
            continue

        for shape in list(slide.shapes):
            if shape.name in ("Mascot", "Mascot_Origin"):
                shape._element.getparent().remove(shape._element)

        marker = SPEECH_TEXT_MARKERS[slide_number]
        bubble_text = next(
            (
                shape
                for shape in slide.shapes
                if getattr(shape, "has_text_frame", False)
                and marker in shape.text
                and shape.top >= Inches(5.4)
            ),
            None,
        )
        if bubble_text is None:
            raise ValueError(f"Could not find the speech text on slide {slide_number}.")

        accent_shapes = [
            shape
            for shape in slide.shapes
            if _shape_fill_rgb(shape) == ACCENT and shape.top >= Inches(5.4)
        ]
        bubble = next(
            (
                shape
                for shape in accent_shapes
                if shape.name == "SpeechBubble_Background"
                or shape.width >= Inches(2.0)
            ),
            None,
        )
        tail = next(
            (
                shape
                for shape in accent_shapes
                if shape.name == "SpeechBubble_Tail"
                or shape.width < Inches(1.0)
            ),
            None,
        )
        if bubble is None or tail is None:
            raise ValueError(f"Could not find all speech shapes on slide {slide_number}.")

        bubble_x = SPEECH_BUBBLE_RIGHT - bubble.width / Inches(1)
        bubble_w = bubble.width / Inches(1)
        bubble_h = bubble.height / Inches(1)
        bubble_y = SPEECH_TAIL_TARGET_Y - 0.28 - bubble_h + 0.04

        bubble.left = Inches(bubble_x)
        bubble.top = Inches(bubble_y)
        bubble.name = "SpeechBubble_Background"
        bubble_text.left = Inches(bubble_x + 0.22)
        bubble_text.top = Inches(bubble_y + 0.10)
        bubble_text.width = Inches(bubble_w - 0.44)
        bubble_text.height = Inches(bubble_h - 0.20)
        bubble_text.name = "SpeechBubble_Text"

        tail.left = Inches(SPEECH_BUBBLE_RIGHT - 0.22)
        tail.top = Inches(bubble_y + bubble_h - 0.04)
        tail.width = Inches(0.38)
        tail.height = Inches(0.28)
        tail.rotation = 180
        tail.name = "SpeechBubble_Tail"

        slide_mascot_path = (
            serious_mascot_path if slide_number == 11 else mascot_path
        )
        mascot = slide.shapes.add_picture(
            str(slide_mascot_path),
            Inches(MASCOT_X),
            Inches(MASCOT_Y),
            Inches(MASCOT_SIZE),
            Inches(MASCOT_SIZE),
        )
        mascot.name = "Mascot"

        origin = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(MASCOT_ORIGIN_X - 0.02),
            Inches(MASCOT_ORIGIN_Y - 0.02),
            Inches(0.04),
            Inches(0.04),
        )
        origin.fill.background()
        origin.line.fill.background()
        origin.name = "Mascot_Origin"


def apply_speech_animations(presentation_path, serious_mascot_path):
    """Replace speech parts with native callouts and add Zoom entrances."""
    if not ANIMATION_SCRIPT.exists():
        raise FileNotFoundError(f"Missing animation script: {ANIMATION_SCRIPT}")
    result = subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(ANIMATION_SCRIPT),
            "-PresentationPath",
            str(presentation_path.resolve()),
            "-SeriousMascotPath",
            str(serious_mascot_path.resolve()),
        ],
        cwd=SCRIPT_DIR,
        check=True,
        capture_output=True,
        text=True,
    )
    if result.stdout.strip():
        print(result.stdout.strip())


def add_notes(slide, notes, source=None, visual_spec=None):
    text_frame = slide.notes_slide.notes_text_frame
    text_frame.text = notes
    if visual_spec:
        paragraph = text_frame.add_paragraph()
        paragraph.text = f"Visual: {visual_spec}"
    if source:
        paragraph = text_frame.add_paragraph()
        paragraph.text = f"[Sources]\n- {source}"


def add_picture_cover(slide, image_path, x, y, w, h):
    """Add an image using a centered 'cover' crop, like CSS object-fit: cover."""
    from PIL import Image

    with Image.open(image_path) as image:
        image_w, image_h = image.size
    image_ratio = image_w / image_h
    frame_ratio = w / h

    picture = slide.shapes.add_picture(
        str(image_path), Inches(x), Inches(y), Inches(w), Inches(h)
    )
    if image_ratio > frame_ratio:
        visible_fraction = frame_ratio / image_ratio
        crop = (1.0 - visible_fraction) / 2.0
        picture.crop_left = crop
        picture.crop_right = crop
    elif image_ratio < frame_ratio:
        visible_fraction = image_ratio / frame_ratio
        crop = (1.0 - visible_fraction) / 2.0
        picture.crop_top = crop
        picture.crop_bottom = crop
    return picture


def add_rounded_picture_cover(
    slide, image_path, x, y, w, h, line_color="516177", line_width=1.5
):
    """Add a cover-cropped image with rounded corners and a visible frame."""
    picture = add_picture_cover(slide, image_path, x, y, w, h)
    picture._element.spPr.prstGeom.set("prst", "roundRect")

    frame = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)
    )
    frame.fill.background()
    frame.line.color.rgb = rgb(line_color)
    frame.line.width = Pt(line_width)
    return picture


def add_picture_contain(slide, image_path, x, y, w, h, dark=False):
    """Fit a complete image inside a rounded presentation frame."""
    from PIL import Image

    frame_fill = "182338" if dark else "E8EDF3"
    frame_line = "516177" if dark else MUTED
    add_rect(slide, x, y, w, h, frame_fill, frame_line, radius=True, line_width=1.2)

    with Image.open(image_path) as image:
        image_w, image_h = image.size
    scale = min((w - 0.20) / image_w, (h - 0.20) / image_h)
    placed_w = image_w * scale
    placed_h = image_h * scale
    placed_x = x + (w - placed_w) / 2
    placed_y = y + (h - placed_h) / 2
    picture = slide.shapes.add_picture(
        str(image_path),
        Inches(placed_x),
        Inches(placed_y),
        Inches(placed_w),
        Inches(placed_h),
    )
    picture._element.spPr.prstGeom.set("prst", "roundRect")
    return picture


def build_slide_1(prs, visual_1):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_picture(
        str(visual_1), 0, 0, width=prs.slide_width, height=prs.slide_height
    )

    add_rect(slide, 1.58, 0.32, 8.36, 1.68, DEEP, None, radius=True)
    add_text(
        slide,
        "The Raycast Challenge",
        1.84,
        0.53,
        7.80,
        0.65,
        50,
        WHITE,
        font=HEADER_FONT,
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
    )
    add_text(
        slide,
        "Pick a pixel in one drone frame. Find the same point in twelve others. ≤10 pixels or it doesn't count.",
        1.86,
        1.27,
        7.82,
        0.48,
        19.5,
        OFFWHITE,
        font=BODY_FONT,
        valign=MSO_ANCHOR.MIDDLE,
    )
    add_act_rail(slide, 1)
    add_speech_bubble(slide, '"≤10 pixels in your dreams, mate."', 9.27, 6.20, 3.55, 0.82)
    add_notes(
        slide,
        'Open cold with the image, not with the org chart. This is a "here\'s the trick" opener — the rest of the deck is "here\'s what it took."',
        "User-provided visual: Raycast_Slide_1_Visual.png",
        "A single dramatic frame from the footage — the van, mid-field, fisheye-distorted — with a crosshair on it and thin lines fanning out to ~12 hovering drones (cartoon). This is the whole pitch in one image. [NTPvChat]",
    )


def build_slide_2(prs, visual_2):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, OFFWHITE)
    add_title(slide, "The Assignment", dark=False)

    add_picture_cover(slide, visual_2, 1.75, 1.42, 4.18, 5.55)

    items = [
        "A sequence of drone FPV frames, one stationary van, one agricultural field",
        "Several frames badly motion/defocus-blurred",
        "Deliverable: source code, a methodology writeup, and proof — not just a claim — of ≤10px reprojection accuracy",
        'Complete architectural freedom. "Ingenuity and a tinkerer mindset," officially encouraged.',
    ]
    positions = [
        (6.22, 1.42),
        (9.57, 1.42),
        (6.22, 4.28),
        (9.57, 4.28),
    ]
    for index, (text, (x, y)) in enumerate(zip(items, positions), start=1):
        add_body_card(slide, index, text, x, y, 3.05, 2.69, dark=False)

    add_act_rail(slide, 2)
    add_notes(
        slide,
        "This slide exists so the audience knows the constraints going in: freedom, blur, repetition, and a hard numeric bar.",
    )


def build_slide_3(prs, sharp_frame, blurred_frame):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, PRIMARY)
    add_title(slide, "The Challenges", dark=True)

    visual_instruction = (
        "Side-by-side crop of a sharp frame vs. one of the worst-blurred frames, same rough patch of ground. [NTPvChat]"
    )
    add_rect(slide, 1.75, 1.42, 6.15, 5.55, "182338", "516177", radius=True, line_width=1.2)
    add_rounded_picture_cover(slide, sharp_frame, 1.90, 1.94, 2.805, 4.12)
    add_rounded_picture_cover(slide, blurred_frame, 4.825, 1.94, 2.805, 4.12)
    add_rect(slide, 1.90, 1.94, 2.805, 0.48, PRIMARY, None, radius=True)
    add_text(
        slide, "SHARP FRAME", 1.90, 1.94, 2.805, 0.48, 14, SECONDARY,
        font=HEADER_FONT, bold=True, align=PP_ALIGN.CENTER,
        valign=MSO_ANCHOR.MIDDLE,
    )
    add_rect(slide, 4.825, 1.94, 2.805, 0.48, PRIMARY, None, radius=True)
    add_text(
        slide, "WORST-BLURRED FRAME", 4.825, 1.94, 2.805, 0.48, 13, ACCENT,
        font=HEADER_FONT, bold=True, align=PP_ALIGN.CENTER,
        valign=MSO_ANCHOR.MIDDLE,
    )
    add_text(
        slide, visual_instruction, 1.95, 6.25, 5.75, 0.56, 10.5, PALE,
        font=BODY_FONT, italic=True, valign=MSO_ANCHOR.MIDDLE,
    )

    items = [
        "The field is highly repetitive — one patch of dirt looks like every other patch of dirt",
        "Several frames are so blurred that even a human struggles to place a feature confidently",
        "GPS telemetry on the drones is off by multiple metres — you can't just trust the numbers you're handed",
    ]
    ys = [1.42, 3.28, 5.14]
    for index, (text, y) in enumerate(zip(items, ys), start=1):
        add_body_card(slide, index, text, 8.22, y, 4.60, 1.65, dark=True)

    add_act_rail(slide, 3)
    add_notes(
        slide,
        "This is the thesis of the whole deck: every subsequent architectural decision is a response to one of these three problems.",
    )


def build_slide_4(prs, visual_4):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, PRIMARY)
    add_title(
        slide,
        "First Commit: Already a Real Pipeline (Thanks, Claude :-))",
        dark=True,
        font_size=31,
    )

    visual_spec = (
        "Load Frames → OCR (reads the HUD overlay) → Undistort → Pose "
        "(GPS→ENU + prior rotation) → Refine (`pitch_optimizer.py`) → "
        "Interactive Viewer (ray-cast on click).\nRefine → Blender. [NTPvChat]"
    )
    add_picture_contain(slide, visual_4, 1.72, 1.42, 2.85, 5.42, dark=True)

    items = [
        "OCR reads GPS/heading/altitude straight off the HUD overlay",
        "GPS → local ENU coordinates",
        "Fisheye lens undistortion",
        "A hand-rolled pitch optimizer",
        "Ray–ground intersection math",
        "An interactive click-to-reproject viewer",
        "Debug argument: Export solve to Blender",
    ]
    positions = [
        (4.82, 1.48),
        (8.78, 1.48),
        (4.82, 2.65),
        (8.78, 2.65),
        (4.82, 3.82),
        (8.78, 3.82),
        (4.82, 4.99),
    ]
    for index, (text, (x, y)) in enumerate(zip(items, positions), start=1):
        add_body_card(
            slide, index, text, x, y, 3.66, 1.02, dark=True, font_size=12.7
        )

    add_act_i_rail(slide, 4)
    add_speech_bubble(
        slide,
        '"First commit after ~4 workday... literally a war-crime."',
        8.38,
        6.24,
        4.36,
        0.72,
    )
    add_notes(
        slide,
        "Establish early that day one wasn't a toy — it was a mature skeleton that mostly survived. Mention roll bug which was easily spotted in Blender.",
        visual_spec=visual_spec,
    )


def build_slide_5(prs, visual_5):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, OFFWHITE)
    add_title(slide, "Lesson #1, Learned in the First Hour", font_size=36)

    body = (
        'At this stage, pitch is computed from the ground plane — and the van is '
        'a real 3D object sitting on top of that plane, not part of it. Mask it '
        'out, or "what does flat ground look like" gets answered by the wrong geometry.'
    )
    add_rect(slide, 1.75, 1.48, 4.22, 4.88, WHITE, PALE, radius=True, line_width=1.2)
    add_text(
        slide,
        body,
        2.05,
        1.84,
        3.62,
        4.16,
        20,
        PRIMARY,
        font=BODY_FONT,
        valign=MSO_ANCHOR.MIDDLE,
    )

    visual_spec = (
        'A frame with the van bounding-boxed and crossed out in red, labeled '
        '"real geometry — wrong category." [NTPvChat]'
    )
    add_rounded_picture_cover(
        slide, visual_5, 6.23, 1.48, 6.57, 4.65, line_color=MUTED, line_width=1.2
    )

    add_act_i_rail(slide, 5)
    add_speech_bubble(
        slide,
        '"You have no idea how important those masked-out van pixels will become towards the middle of the project and onwards."',
        6.98,
        6.25,
        5.76,
        0.82,
    )
    add_notes(
        slide,
        "This is specifically about excluding a real foreground object from a ground-only computation — don't conflate it with the screen-overlay masking that shows up later (Slide 7 onward); that's a related-but-distinct lesson with its own payoff.",
        visual_spec=visual_spec,
    )


def replace_slide_5_placeholder(prs, visual_5):
    """Replace the existing Slide 5 placeholder without rebuilding prior slides."""
    slide = prs.slides[4]
    visual_left = Inches(6.20)
    visual_top = Inches(1.40)
    visual_bottom = Inches(6.10)

    existing_picture = any(
        shape.shape_type == 13
        and shape.left >= visual_left
        and shape.top >= visual_top
        and shape.top < visual_bottom
        for shape in slide.shapes
    )
    if existing_picture:
        return

    placeholder_shapes = [
        shape
        for shape in slide.shapes
        if shape.left >= visual_left
        and shape.top >= visual_top
        and shape.top < visual_bottom
    ]
    for shape in placeholder_shapes:
        shape._element.getparent().remove(shape._element)

    add_rounded_picture_cover(
        slide, visual_5, 6.23, 1.48, 6.57, 4.65, line_color=MUTED, line_width=1.2
    )


def replace_slide_visual(
    prs, slide_number, image_path, x, y, w, h, shape_name, dark=True
):
    """Replace a named slide's visual placeholder without disturbing its text."""
    slide = prs.slides[slide_number - 1]
    if any(shape.name == shape_name for shape in slide.shapes):
        return

    left = Inches(x - 0.04)
    top = Inches(y - 0.04)
    right = Inches(x + w + 0.04)
    bottom = Inches(y + h + 0.04)
    placeholder_shapes = [
        shape
        for shape in slide.shapes
        if shape.left >= left
        and shape.top >= top
        and shape.left + shape.width <= right
        and shape.top + shape.height <= bottom
    ]
    for shape in placeholder_shapes:
        shape._element.getparent().remove(shape._element)

    picture = add_picture_contain(slide, image_path, x, y, w, h, dark=dark)
    picture.name = shape_name


def build_slide_6(prs, visual_6):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, PRIMARY)
    add_title(
        slide,
        '"Lots of Improvements Tried. Result Still Not Amazing."',
        dark=True,
        font_size=28,
    )

    add_rect(
        slide, 1.75, 1.48, 4.48, 4.98, DEEP, "334155", radius=True, line_width=1.2
    )
    add_rich_text(
        slide,
        [
            ("Day two: the solver gets switched to Ceres (", BODY_FONT),
            ("pyceres", "Cascadia Mono"),
            (
                ") — real nonlinear least-squares, not a hand-rolled optimizer. "
                "Two tools are born here that outlive nearly everything else in "
                "the codebase: ",
                BODY_FONT,
            ),
            ("camera_deltas.py", "Cascadia Mono"),
            (
                " (solved-vs-telemetry sanity check) and a GeoCalib-based "
                "pitch/roll estimator.",
                BODY_FONT,
            ),
        ],
        2.06,
        1.82,
        3.86,
        4.30,
        19,
        OFFWHITE,
        valign=MSO_ANCHOR.MIDDLE,
    )

    visual_spec = "User-provided visual: Raycast_Slide_6_Visual.png"
    picture = add_picture_contain(
        slide, visual_6, 6.50, 1.48, 6.30, 4.98, dark=True
    )
    picture.name = "Slide6_Visual"

    add_act_i_rail(slide, 6)
    add_notes(
        slide,
        'This is the first "one twin lives, one twin fades" beat — good rhythm-setter for the whole deck\'s structure.',
        visual_spec=visual_spec,
    )


def build_slide_7(prs, visual_7):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, PRIMARY)
    add_title(slide, "Ten Manual Graphics Overlay Masks", dark=True, font_size=36)

    add_rect(
        slide, 1.75, 1.48, 4.45, 4.98, DEEP, "334155", radius=True, line_width=1.2
    )
    add_rich_text(
        slide,
        [
            ("HUD_REGIONS", "Cascadia Mono"),
            (
                ": 10 rectangles, geometrically masking every HUD overlay element "
                "in raw camera pixels. The top-centre navigation bar alone is split "
                "into three separate boxes — because one wide box, dragged through "
                "fisheye undistortion, would bow at the corners and expose exactly "
                "the content it was supposed to hide.",
                BODY_FONT,
            ),
        ],
        2.06,
        1.82,
        3.83,
        4.30,
        18,
        OFFWHITE,
        valign=MSO_ANCHOR.MIDDLE,
    )

    visual_spec = "User-provided visual: Raycast_Slide_7_Visual.png"
    picture = add_picture_contain(
        slide, visual_7, 6.48, 1.48, 6.32, 4.98, dark=True
    )
    picture.name = "Slide7_Visual"
    add_act_ii_rail(slide, 7)
    add_notes(slide, "", visual_spec=visual_spec)


def add_video_placeholder(slide, x, y, w, h):
    """Add a clean 16:9 frame ready for a future embedded video."""
    frame = add_rect(
        slide, x, y, w, h, "E8EDF3", MUTED, radius=True, line_width=1.2
    )
    frame.name = "Slide8_Video_Placeholder"

    circle_size = 0.78
    circle = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(x + (w - circle_size) / 2),
        Inches(y + (h - circle_size) / 2 - 0.08),
        Inches(circle_size),
        Inches(circle_size),
    )
    circle.fill.solid()
    circle.fill.fore_color.rgb = rgb(PRIMARY)
    circle.line.fill.background()

    play = slide.shapes.add_shape(
        MSO_SHAPE.ISOSCELES_TRIANGLE,
        Inches(x + w / 2 - 0.13),
        Inches(y + h / 2 - 0.23),
        Inches(0.28),
        Inches(0.34),
    )
    play.rotation = 90
    play.fill.solid()
    play.fill.fore_color.rgb = rgb(WHITE)
    play.line.fill.background()
    add_text(
        slide,
        "LANDSCAPE VIDEO PLACEHOLDER",
        x + 0.35,
        y + h - 0.50,
        w - 0.70,
        0.26,
        11,
        PRIMARY,
        font=HEADER_FONT,
        bold=True,
        align=PP_ALIGN.CENTER,
        valign=MSO_ANCHOR.MIDDLE,
    )


def add_slide_8_stage_card(slide, number, text, x, font_size):
    """Add one compact stage in Slide 8's horizontal process row."""
    y, w, h = 5.72, 2.68, 1.42
    add_rect(slide, x, y, w, h, WHITE, PALE, radius=True, line_width=1.0)
    add_text(
        slide,
        f"{number:02d}",
        x + 0.15,
        y + 0.10,
        0.36,
        0.20,
        9.2,
        SECONDARY,
        font=HEADER_FONT,
        bold=True,
        valign=MSO_ANCHOR.MIDDLE,
    )
    mark = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(x + 0.15),
        Inches(y + 0.35),
        Inches(0.34),
        Inches(0.03),
    )
    mark.fill.solid()
    mark.fill.fore_color.rgb = rgb(SECONDARY)
    mark.line.fill.background()
    add_text(
        slide,
        text,
        x + 0.16,
        y + 0.44,
        w - 0.32,
        h - 0.53,
        font_size,
        PRIMARY,
        font=BODY_FONT,
        valign=MSO_ANCHOR.MIDDLE,
        margin=0.02,
    )


def layout_slide_8(slide):
    """Lay out Slide 8 around a central landscape video."""
    add_title(slide, "The Manual Correspondence Picker Is Born", font_size=35)
    add_video_placeholder(slide, 3.70, 1.35, 7.25, 4.08)

    items = [
        "Multi-frame grid, independent zoom/pan, adjustable marker size",
        (
            "Four feature types, cycled with a key: ground, roof, wheel_axis, "
            "roof_edge — the last two are pairs with known real-world distances "
            "(wheelbase, van width) baked in as hard constraints"
        ),
        (
            "Right-click removes a single point; d deletes a whole correspondence "
            "or just one frame's mark, depending on context; [/] step through saved "
            "correspondences filtered by type"
        ),
        (
            "A quit-safety net: unsaved changes require a second keypress to "
            "actually discard"
        ),
    ]
    positions = [1.75, 4.54, 7.33, 10.12]
    font_sizes = [10.0, 8.8, 8.6, 9.4]
    for index, (text, x, font_size) in enumerate(
        zip(items, positions, font_sizes), start=1
    ):
        add_slide_8_stage_card(slide, index, text, x, font_size)


def normalize_slide_8_layout(prs):
    """Rebuild Slide 8's content area without disturbing its rail or notes."""
    if len(prs.slides) < 8:
        return
    slide = prs.slides[7]
    content_shapes = [
        shape for shape in slide.shapes if shape.left >= Inches(1.43)
    ]
    for shape in content_shapes:
        shape._element.getparent().remove(shape._element)
    set_background(slide, OFFWHITE)
    layout_slide_8(slide)


def build_slide_8(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, OFFWHITE)
    layout_slide_8(slide)
    add_act_ii_rail(slide, 8)
    add_notes(
        slide,
        "This is worth a real video, not a mockup — it genuinely looks like a "
        "purpose-built desktop app, which surprises people expecting a debug script.",
        visual_spec="Landscape video of the manual correspondence picker in action.",
    )


def build_slide_9(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, PRIMARY)
    add_title(slide, "The Scorer: A Heads-Up, Not a Verdict", dark=True, font_size=36)

    body = (
        'For every correspondence, crop the clicked point in each frame, embed '
        'with CLIP, average pairwise similarity. A low score is genuinely '
        'ambiguous: it can mean "wrong pick" — or "right pick, just doesn\'t look '
        'alike from this angle/blur." CLIP only sees pixels, not the context that '
        "told you it's the same rivet."
    )
    add_rect(
        slide, 1.75, 1.48, 4.52, 4.98, DEEP, "334155", radius=True, line_width=1.2
    )
    add_text(
        slide,
        body,
        2.07,
        1.82,
        3.88,
        4.30,
        18.5,
        OFFWHITE,
        font=BODY_FONT,
        valign=MSO_ANCHOR.MIDDLE,
    )

    visual_spec = (
        'Three crop pairs side by side, unlabeled at first glance: (1) high score, '
        'obviously the same feature; (2) low score, still correct — just an '
        'oblique-angle or blur mismatch; (3) low score, actually a mis-click. '
        'The point: (2) and (3) look equally "bad" to the algorithm. Only a human '
        'looking at the full frame — not just the crop — can tell them apart. '
        "[NTPvCode]"
    )
    add_visual_placeholder(
        slide, visual_spec, 6.54, 1.48, 6.26, 4.98, dark=True
    )
    add_act_ii_rail(slide, 9)
    add_notes(
        slide,
        "Emphasize what it's not: not a 3D geometric check, purely \"do these crops look alike\" — and that's a real limitation, not a footnote. The project's own README says as much: low scores across oblique frame pairs are expected, not necessarily wrong.",
        visual_spec=visual_spec,
    )


def build_slide_10(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, OFFWHITE)
    add_title(
        slide, '"Manual Feature Matching Works Great on 7 Frames"', font_size=28
    )

    visual_spec = (
        "The best-looking proof-sheet reprojection from this era, one source pixel "
        "accurately landing across all 7 frames. [NTPvCode]"
    )
    add_visual_placeholder(
        slide, visual_spec, 1.75, 1.48, 6.18, 4.82, dark=False
    )

    body = (
        "Known van geometry (wheel axis, roof edges) as hard anchors + hand-picked "
        "ground correspondences → a genuinely good calibration on the 7 clearest "
        "frames. This becomes the quality bar the rest of the project measures "
        "itself against."
    )
    add_rect(slide, 8.20, 1.48, 4.60, 4.82, WHITE, PALE, radius=True, line_width=1.2)
    add_text(
        slide,
        body,
        8.52,
        1.84,
        3.96,
        4.08,
        19,
        PRIMARY,
        font=BODY_FONT,
        valign=MSO_ANCHOR.MIDDLE,
    )

    add_act_ii_rail(slide, 10)
    add_speech_bubble(
        slide,
        '"This exact commit also complains, in its own message, about frame 12035 '
        "hitting the solver's yaw and roll limits, bit it was a small matter of "
        'releasing the limits in the config, and the hard frame got solved as well."',
        6.62,
        6.22,
        6.12,
        0.98,
        font_size=10.8,
    )
    add_notes(
        slide,
        "Good beat to pause on — first real proof of concept, worth letting it breathe before complicating it.",
        visual_spec=visual_spec,
    )


def build_slide_11(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, PRIMARY)
    add_title(slide, "The Other Six", dark=True, font_size=40)

    body = (
        "Two-stage auto-matching: manual correspondences as a trusted anchor, "
        "LightGlue auto-matches extending coverage outward from there. First "
        "appearance of a pattern that gets reinvented twice more before this "
        "project is done: trust a good core, extend outward carefully."
    )
    add_rect(
        slide, 1.75, 1.48, 4.45, 4.82, DEEP, "334155", radius=True, line_width=1.2
    )
    add_text(
        slide,
        body,
        2.07,
        1.84,
        3.81,
        4.10,
        18.5,
        OFFWHITE,
        font=BODY_FONT,
        valign=MSO_ANCHOR.MIDDLE,
    )

    visual_spec = (
        'A simple diagram — 7 solid "trusted" nodes in the center, 6 dashed '
        '"extended" nodes reaching out from them. [NTPvChat]'
    )
    add_visual_placeholder(
        slide, visual_spec, 6.48, 1.48, 6.32, 4.82, dark=True
    )
    add_act_ii_rail(slide, 11)
    add_speech_bubble(
        slide,
        '"Yeah, so much effort, and great thinking, but it just didn\'t F***ing solve."',
        7.74,
        6.23,
        5.00,
        0.82,
        font_size=13.2,
    )
    add_notes(
        slide,
        "Plant this pattern explicitly here — it pays off big later (CLIP weighting, then the two-stage MASt3R split) and the audience should recognize it when it returns.",
        visual_spec=visual_spec,
    )


def normalize_slide_10_title(prs):
    """Keep the long Slide 10 title clear of the accent rule in existing decks."""
    if len(prs.slides) < 10:
        return
    expected_title = '"Manual Feature Matching Works Great on 7 Frames"'
    for shape in prs.slides[9].shapes:
        if getattr(shape, "has_text_frame", False) and shape.text == expected_title:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(28)
            return


def main():
    parser = argparse.ArgumentParser(
        description="Build the Raycast Challenge deck through Act II."
    )
    parser.add_argument(
        "--visuals-dir",
        type=Path,
        default=DEFAULT_VISUALS_DIR,
        help="Folder containing generated slide visuals.",
    )
    parser.add_argument(
        "--dataset-dir",
        type=Path,
        default=DEFAULT_DATASET_DIR,
        help="Folder containing source dataset images.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=OUTPUT_PPTX,
        help="Destination .pptx path.",
    )
    parser.add_argument(
        "--mascot",
        type=Path,
        default=DEFAULT_MASCOT_PATH,
        help="Transparent mascot PNG used by speech overlays.",
    )
    parser.add_argument(
        "--serious-mascot",
        type=Path,
        default=DEFAULT_SERIOUS_MASCOT_PATH,
        help="Red-portal mascot PNG used on Slide 11 only.",
    )
    parser.add_argument(
        "--skip-animations",
        action="store_true",
        help="Save without applying PowerPoint entrance animations.",
    )
    args = parser.parse_args()

    visual_1 = find_asset(args.visuals_dir, "Raycast_Slide_1_Visual.png")
    visual_2 = find_asset(args.visuals_dir, "Raycast_Slide_2_Visual.png")
    visual_4 = find_asset(args.visuals_dir, "Raycast_Slide_4_Visual.png")
    visual_5 = find_asset(args.visuals_dir, "Raycast_Slide_5_Visual.png")
    visual_6 = find_asset(args.visuals_dir, "Raycast_Slide_6_Visual.png")
    visual_7 = find_asset(args.visuals_dir, "Raycast_Slide_7_Visual.png")
    sharp_frame = find_asset(args.dataset_dir, "2026-02-15_16-25-03_12035.png")
    blurred_frame = find_asset(args.dataset_dir, "2026-02-15_16-25-03_04681.png")

    args.output.parent.mkdir(parents=True, exist_ok=True)
    if args.output.exists():
        prs = Presentation(str(args.output))
        if len(prs.slides) not in (3, 6, 11):
            raise ValueError(
                f"Expected 3, 6, or 11 existing slides in {args.output}, "
                f"found {len(prs.slides)}."
            )
    else:
        prs = Presentation()
        prs.slide_width = Inches(SLIDE_W)
        prs.slide_height = Inches(SLIDE_H)
        prs.core_properties.title = "The Raycast Challenge — Acts 0 through II"
        prs.core_properties.subject = "The Setup, Bring-Up, and By Hand"
        prs.core_properties.author = "Ofer Shkedi"
        prs.core_properties.keywords = "Raycast, drone, reprojection, calibration"

        build_slide_1(prs, visual_1)
        build_slide_2(prs, visual_2)
        build_slide_3(prs, sharp_frame, blurred_frame)

    if len(prs.slides) == 3:
        build_slide_4(prs, visual_4)
        build_slide_5(prs, visual_5)
        build_slide_6(prs, visual_6)
    else:
        replace_slide_5_placeholder(prs, visual_5)

    if len(prs.slides) == 6:
        build_slide_7(prs, visual_7)
        build_slide_8(prs)
        build_slide_9(prs)
        build_slide_10(prs)
        build_slide_11(prs)

    replace_slide_visual(
        prs, 6, visual_6, 6.50, 1.48, 6.30, 4.98, "Slide6_Visual", dark=True
    )
    replace_slide_visual(
        prs, 7, visual_7, 6.48, 1.48, 6.32, 4.98, "Slide7_Visual", dark=True
    )
    normalize_slide_8_layout(prs)
    normalize_slide_10_title(prs)
    normalize_all_act_rails(prs)
    prepare_speech_overlays(prs, args.mascot, args.serious_mascot)
    prs.save(args.output)
    if not args.skip_animations:
        apply_speech_animations(args.output, args.serious_mascot)
    print(args.output)


if __name__ == "__main__":
    main()
