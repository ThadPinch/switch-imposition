import glob
import math
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Tuple

def ensure_pypdf_dependency():
    try:
        import PyPDF2  # noqa: F401
        return
    except ModuleNotFoundError:
        pass

    candidate_roots = []
    conda_prefix = os.environ.get("CONDA_PREFIX")
    if conda_prefix:
        candidate_roots.append(conda_prefix)

    home = Path.home()
    candidate_roots.extend(
        [
            str(home / "opt" / "anaconda3"),
            str(home / "anaconda3"),
            str(home / "miniconda3"),
            "/opt/anaconda3",
            "/opt/miniconda3",
        ]
    )

    seen = set()
    for root in candidate_roots:
        if not root or root in seen:
            continue
        seen.add(root)
        for site_packages in glob.glob(os.path.join(root, "lib", "python*", "site-packages")):
            has_pypdf2 = os.path.exists(os.path.join(site_packages, "PyPDF2", "__init__.py"))
            has_pypdf = os.path.exists(os.path.join(site_packages, "pypdf", "__init__.py"))
            if has_pypdf2 or has_pypdf:
                if site_packages not in sys.path:
                    sys.path.append(site_packages)
                return


ensure_pypdf_dependency()

try:
    from PyPDF2 import PdfReader, PdfWriter, Transformation
    from PyPDF2._page import PageObject
    from PyPDF2.generic import DecodedStreamObject, NameObject
except ModuleNotFoundError as exc:
    raise ModuleNotFoundError(
        "PyPDF2 is not available. Use the Anaconda Python that already has it, or install PyPDF2/pypdf for this interpreter."
    ) from exc


PT = 72.0
OUTER_MARGIN_IN = 0.125


@dataclass
class ImpositionConfig:
    mode: str
    cut_width_in: float
    cut_height_in: float
    bleed_width_in: float
    bleed_height_in: float
    sheet_width_in: float
    sheet_height_in: float
    gap_horizontal_in: float
    gap_vertical_in: float
    sheet_orientation: str
    best_fit: bool
    duplex: bool
    auto_correct_art_orientation: bool
    art_rotation: str
    binding_edge: str
    rotate_first_column_or_row: bool
    image_shift_x_in: float
    image_shift_y_in: float


class ImpositionError(RuntimeError):
    pass


def impose_pdf(source_path: str, output_path: str, config: ImpositionConfig) -> Dict[str, object]:
    with open(source_path, "rb") as source_file:
        reader = PdfReader(source_file, strict=False)
        page_count = len(reader.pages)
        if page_count == 0:
            raise ImpositionError("The uploaded PDF has no pages.")

        resolved_sheet = resolve_sheet_configuration(config)
        if resolved_sheet["capacity"] <= 0:
            raise ImpositionError(
                "The cut size does not fit on the selected sheet with the current gaps and 0.125 inch outer margins."
            )

        layout = plan_layout(
            resolved_sheet["sheet_w_in"],
            resolved_sheet["sheet_h_in"],
            config.cut_width_in,
            config.cut_height_in,
            config.gap_horizontal_in,
            config.gap_vertical_in,
            resolved_sheet["capacity"],
        )
        if layout is None:
            raise ImpositionError("The chosen dimensions could not produce a valid layout.")

        auto_rotate_source = (
            config.auto_correct_art_orientation
            and should_auto_correct_orientation(config, reader.pages[0])
        )
        positions_per_sheet = layout["cols"] * layout["rows"]
        writer = PdfWriter()

        if config.mode == "cutAndStack":
            plan = build_cut_and_stack_plan(page_count, positions_per_sheet)
            output_sheet_count = plan["sheet_count"]
            for sheet_index in range(output_sheet_count):
                output_page = create_output_page(writer, layout)
                draw_crop_overlay(output_page, layout, config, sheet_index)

                for row in range(layout["rows"]):
                    for column in range(layout["cols"]):
                        position_index = row * layout["cols"] + column
                        source_index = get_cut_and_stack_source_index(plan, sheet_index, position_index)
                        if source_index < 0:
                            continue
                        prepared_page, prepared_size = prepare_source_page(
                            reader.pages[source_index],
                            config,
                            auto_rotate_source,
                        )
                        place_prepared_page(
                            output_page,
                            prepared_page,
                            prepared_size,
                            layout,
                            config,
                            row,
                            column,
                            sheet_index,
                        )
                output_page.compress_content_streams()
                writer.add_page(output_page)
        else:
            output_sheet_count = page_count
            for source_index in range(page_count):
                output_page = create_output_page(writer, layout)
                draw_crop_overlay(output_page, layout, config, source_index)
                prepared_page, prepared_size = prepare_source_page(
                    reader.pages[source_index],
                    config,
                    auto_rotate_source,
                )
                for row in range(layout["rows"]):
                    for column in range(layout["cols"]):
                        place_prepared_page(
                            output_page,
                            prepared_page,
                            prepared_size,
                            layout,
                            config,
                            row,
                            column,
                            source_index,
                        )
                output_page.compress_content_streams()
                writer.add_page(output_page)

        with open(output_path, "wb") as output_file:
            writer.write(output_file)

    return {
        "pageCount": page_count,
        "outputSheetCount": output_sheet_count,
        "cols": layout["cols"],
        "rows": layout["rows"],
        "sheetWidthIn": resolved_sheet["sheet_w_in"],
        "sheetHeightIn": resolved_sheet["sheet_h_in"],
        "sheetOrientation": resolved_sheet["actual_orientation"],
        "outputFileName": build_output_file_name(source_path, config.mode, layout["cols"], layout["rows"]),
        "autoRotatedSource": auto_rotate_source,
    }


def parse_config(values: Dict[str, str]) -> ImpositionConfig:
    try:
        return ImpositionConfig(
            mode=parse_mode(values.get("mode")),
            cut_width_in=parse_positive_float(values.get("cutWidthIn"), "Cut width"),
            cut_height_in=parse_positive_float(values.get("cutHeightIn"), "Cut height"),
            bleed_width_in=parse_positive_float(values.get("bleedWidthIn"), "Bleed width"),
            bleed_height_in=parse_positive_float(values.get("bleedHeightIn"), "Bleed height"),
            sheet_width_in=parse_positive_float(values.get("sheetWidthIn"), "Sheet width"),
            sheet_height_in=parse_positive_float(values.get("sheetHeightIn"), "Sheet height"),
            gap_horizontal_in=parse_non_negative_float(values.get("gapHorizontalIn"), "Horizontal gap"),
            gap_vertical_in=parse_non_negative_float(values.get("gapVerticalIn"), "Vertical gap"),
            sheet_orientation=str(values.get("sheetOrientation") or "auto"),
            best_fit=parse_bool(values.get("bestFit")),
            duplex=parse_bool(values.get("duplex")),
            auto_correct_art_orientation=parse_bool(values.get("autoCorrectArtOrientation")),
            art_rotation=str(values.get("artRotation") or "None"),
            binding_edge=str(values.get("bindingEdge") or "Left"),
            rotate_first_column_or_row=parse_bool(values.get("rotateFirstColumnOrRow")),
            image_shift_x_in=parse_signed_float(values.get("imageShiftXIn")),
            image_shift_y_in=parse_signed_float(values.get("imageShiftYIn")),
        )
    except ValueError as exc:
        raise ImpositionError(str(exc)) from exc


def parse_mode(value: str) -> str:
    return "cutAndStack" if str(value or "").strip() == "cutAndStack" else "repeat"


def parse_bool(value: str) -> bool:
    return str(value or "").strip().lower() == "true"


def parse_positive_float(value: str, label: str) -> float:
    number = float(value)
    if not math.isfinite(number) or number <= 0:
        raise ValueError(f"{label} must be greater than zero.")
    return number


def parse_non_negative_float(value: str, label: str) -> float:
    number = float(value)
    if not math.isfinite(number) or number < 0:
        raise ValueError(f"{label} must be zero or greater.")
    return number


def parse_signed_float(value: str) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return 0.0
    return number if math.isfinite(number) else 0.0


def create_output_page(writer: PdfWriter, layout: Dict[str, float]) -> PageObject:
    return PageObject.create_blank_page(
        width=pt(layout["sheet_w_in"]),
        height=pt(layout["sheet_h_in"]),
    )


def resolve_sheet_configuration(config: ImpositionConfig) -> Dict[str, object]:
    sheet_w_in = config.sheet_width_in
    sheet_h_in = config.sheet_height_in

    if config.sheet_orientation == "portrait":
        sheet_w_in = min(config.sheet_width_in, config.sheet_height_in)
        sheet_h_in = max(config.sheet_width_in, config.sheet_height_in)
    elif config.sheet_orientation == "landscape":
        sheet_w_in = max(config.sheet_width_in, config.sheet_height_in)
        sheet_h_in = min(config.sheet_width_in, config.sheet_height_in)

    capacity = max_placements_for_sheet(
        sheet_w_in,
        sheet_h_in,
        config.cut_width_in,
        config.cut_height_in,
        config.gap_horizontal_in,
        config.gap_vertical_in,
    )
    swapped_capacity = max_placements_for_sheet(
        sheet_h_in,
        sheet_w_in,
        config.cut_width_in,
        config.cut_height_in,
        config.gap_horizontal_in,
        config.gap_vertical_in,
    )

    best_fit_swapped = False
    if config.best_fit and swapped_capacity > capacity:
        sheet_w_in, sheet_h_in = sheet_h_in, sheet_w_in
        capacity = swapped_capacity
        best_fit_swapped = True

    return {
        "sheet_w_in": sheet_w_in,
        "sheet_h_in": sheet_h_in,
        "capacity": capacity,
        "best_fit_swapped": best_fit_swapped,
        "actual_orientation": orientation_from_dimensions(sheet_w_in, sheet_h_in),
    }


def max_placements_for_sheet(
    sheet_w_in: float,
    sheet_h_in: float,
    cut_w_in: float,
    cut_h_in: float,
    gap_h_in: float,
    gap_v_in: float,
) -> int:
    avail_w_in = sheet_w_in - OUTER_MARGIN_IN * 2
    avail_h_in = sheet_h_in - OUTER_MARGIN_IN * 2
    if avail_w_in <= 0 or avail_h_in <= 0 or cut_w_in <= 0 or cut_h_in <= 0:
        return 0

    cols_max, rows_max = grid_fit(avail_w_in, avail_h_in, cut_w_in, cut_h_in, gap_h_in, gap_v_in)
    return cols_max * rows_max


def plan_layout(
    sheet_w_in: float,
    sheet_h_in: float,
    cut_w_in: float,
    cut_h_in: float,
    gap_h_in: float,
    gap_v_in: float,
    required_count: int,
):
    required = max(1, int(required_count))
    avail_w_in = sheet_w_in - OUTER_MARGIN_IN * 2
    avail_h_in = sheet_h_in - OUTER_MARGIN_IN * 2
    if avail_w_in <= 0 or avail_h_in <= 0:
        return None

    cols_max, rows_max = grid_fit(avail_w_in, avail_h_in, cut_w_in, cut_h_in, gap_h_in, gap_v_in)
    max_placements = cols_max * rows_max
    if max_placements < required:
        return None

    cols = min(cols_max, required)
    rows_needed = math.ceil(required / cols)
    while rows_needed > rows_max and cols > 1:
        cols -= 1
        rows_needed = math.ceil(required / cols)

    if rows_needed > rows_max:
        return None

    rows = rows_needed
    cell_w_pt = pt(cut_w_in)
    cell_h_pt = pt(cut_h_in)
    gap_h_pt = pt(gap_h_in)
    gap_v_pt = pt(gap_v_in)
    sheet_w_pt = pt(sheet_w_in)
    sheet_h_pt = pt(sheet_h_in)
    arr_w_pt = cols * cell_w_pt + (cols - 1) * gap_h_pt
    arr_h_pt = rows * cell_h_pt + (rows - 1) * gap_v_pt
    off_x = pt(OUTER_MARGIN_IN) + (sheet_w_pt - pt(OUTER_MARGIN_IN) * 2 - arr_w_pt) / 2
    off_y = pt(OUTER_MARGIN_IN) + (sheet_h_pt - pt(OUTER_MARGIN_IN) * 2 - arr_h_pt) / 2

    return {
        "sheet_w_in": sheet_w_in,
        "sheet_h_in": sheet_h_in,
        "cols": cols,
        "rows": rows,
        "cell_w_pt": cell_w_pt,
        "cell_h_pt": cell_h_pt,
        "gap_h_pt": gap_h_pt,
        "gap_v_pt": gap_v_pt,
        "off_x": off_x,
        "off_y": off_y,
    }


def grid_fit(avail_w_in: float, avail_h_in: float, cell_w_in: float, cell_h_in: float, gap_h_in: float, gap_v_in: float) -> Tuple[int, int]:
    cols_max = max(1, math.floor((avail_w_in + gap_h_in) / (cell_w_in + gap_h_in)))
    rows_max = max(1, math.floor((avail_h_in + gap_v_in) / (cell_h_in + gap_v_in)))
    return cols_max, rows_max


def build_cut_and_stack_plan(total_pages: int, positions_per_sheet: int) -> Dict[str, int]:
    per_sheet = max(1, positions_per_sheet)
    sheet_count = math.ceil(total_pages / per_sheet) if total_pages > 0 else 0
    return {"per_sheet": per_sheet, "sheet_count": sheet_count, "total_pages": total_pages}


def get_cut_and_stack_source_index(plan: Dict[str, int], sheet_index: int, position_index: int) -> int:
    source_index = position_index * plan["sheet_count"] + sheet_index
    return source_index if source_index >= 0 and source_index < plan["total_pages"] else -1


def prepare_source_page(source_page: PageObject, config: ImpositionConfig, auto_rotate_source: bool):
    source_width, source_height = get_source_page_size(source_page)
    working_page = clone_page(source_page, source_width, source_height)

    if auto_rotate_source:
        rotated_page = PageObject.create_blank_page(width=source_height, height=source_width)
        rotated_copy = clone_page(working_page, source_width, source_height)
        rotated_copy.add_transformation(Transformation().rotate(90).translate(source_height, 0))
        rotated_page.merge_page(rotated_copy)
        working_page = rotated_page

    target_w_pt = pt(config.bleed_width_in if config.bleed_width_in > config.cut_width_in or config.bleed_height_in > config.cut_height_in else config.cut_width_in)
    target_h_pt = pt(config.bleed_height_in if config.bleed_width_in > config.cut_width_in or config.bleed_height_in > config.cut_height_in else config.cut_height_in)
    source_width, source_height = get_source_page_size(working_page)
    crop_w_pt = min(target_w_pt, source_width)
    crop_h_pt = min(target_h_pt, source_height)
    left = (source_width - crop_w_pt) / 2
    bottom = (source_height - crop_h_pt) / 2

    cropped_page = PageObject.create_blank_page(width=crop_w_pt, height=crop_h_pt)
    cropped_copy = clone_page(working_page, source_width, source_height)
    cropped_copy.add_transformation(Transformation().translate(-left, -bottom))
    cropped_page.merge_page(cropped_copy)
    return cropped_page, {"width": crop_w_pt, "height": crop_h_pt}


def place_prepared_page(
    output_page: PageObject,
    prepared_page: PageObject,
    prepared_size: Dict[str, float],
    layout: Dict[str, float],
    config: ImpositionConfig,
    row: int,
    column: int,
    page_index: int,
):
    flip_positions_this_page = config.duplex and page_index % 2 == 1
    r_eff, c_eff = get_effective_position(row, column, layout, config.binding_edge, flip_positions_this_page)
    center_x = layout["off_x"] + c_eff * (layout["cell_w_pt"] + layout["gap_h_pt"]) + layout["cell_w_pt"] / 2
    center_y = layout["off_y"] + r_eff * (layout["cell_h_pt"] + layout["gap_v_pt"]) + layout["cell_h_pt"] / 2

    place_w_pt = pt(config.bleed_width_in if config.bleed_width_in > config.cut_width_in or config.bleed_height_in > config.cut_height_in else config.cut_width_in)
    place_h_pt = pt(config.bleed_height_in if config.bleed_width_in > config.cut_width_in or config.bleed_height_in > config.cut_height_in else config.cut_height_in)
    draw_w = min(place_w_pt, prepared_size["width"])
    draw_h = min(place_h_pt, prepared_size["height"])
    rotation = rotation_for_page(config, r_eff, c_eff, flip_positions_this_page, page_index)
    shift_x, shift_y = get_shift_adjustments(
        config.binding_edge,
        flip_positions_this_page,
        config.image_shift_x_in,
        config.image_shift_y_in,
    )
    pre_shift_x, pre_shift_y = pre_rotation_shift_for(rotation, shift_x, shift_y)
    draw_x = center_x - place_w_pt / 2 + pre_shift_x + (place_w_pt - draw_w) / 2
    draw_y = center_y - place_h_pt / 2 + pre_shift_y + (place_h_pt - draw_h) / 2
    adjusted_x, adjusted_y = adjust_xy_for_rotation(draw_x, draw_y, draw_w, draw_h, rotation)

    placed_page = clone_page(prepared_page, prepared_size["width"], prepared_size["height"])
    transform = Transformation().rotate(rotation).translate(adjusted_x, adjusted_y)
    placed_page.add_transformation(transform)
    output_page.merge_page(placed_page)


def draw_crop_overlay(output_page: PageObject, layout: Dict[str, float], config: ImpositionConfig, sheet_index: int):
    overlay_page = PageObject.create_blank_page(
        width=pt(layout["sheet_w_in"]),
        height=pt(layout["sheet_h_in"]),
    )
    commands = ["q", "0 0 0 RG", "0.5 w", "1 J"]
    flip_positions_this_page = config.duplex and sheet_index % 2 == 1

    for row in range(layout["rows"]):
        for column in range(layout["cols"]):
            r_eff, c_eff = get_effective_position(row, column, layout, config.binding_edge, flip_positions_this_page)
            center_x = layout["off_x"] + c_eff * (layout["cell_w_pt"] + layout["gap_h_pt"]) + layout["cell_w_pt"] / 2
            center_y = layout["off_y"] + r_eff * (layout["cell_h_pt"] + layout["gap_v_pt"]) + layout["cell_h_pt"] / 2
            commands.extend(
                individual_crop_commands(
                    center_x,
                    center_y,
                    pt(config.cut_width_in),
                    pt(config.cut_height_in),
                    0.25,
                    0.125,
                    c_eff == 0,
                    c_eff == layout["cols"] - 1,
                    r_eff == 0,
                    r_eff == layout["rows"] - 1,
                    layout["gap_h_pt"],
                    layout["gap_v_pt"],
                )
            )

    commands.append("Q")
    stream = DecodedStreamObject()
    stream.set_data("\n".join(commands).encode("ascii"))
    overlay_page[NameObject("/Contents")] = stream
    output_page.merge_page(overlay_page)


def individual_crop_commands(
    center_x: float,
    center_y: float,
    cut_w_pt: float,
    cut_h_pt: float,
    gap_in: float,
    len_in: float,
    is_left_edge: bool,
    is_right_edge: bool,
    is_bottom_edge: bool,
    is_top_edge: bool,
    gap_horizontal_pt: float,
    gap_vertical_pt: float,
):
    off = pt(gap_in)
    perimeter_len = pt(len_in)
    max_interior_len_h = max(0.0, (gap_horizontal_pt - off * 2) * 0.4)
    max_interior_len_v = max(0.0, (gap_vertical_pt - off * 2) * 0.4)
    interior_len_h = min(pt(0.03125), max_interior_len_h)
    interior_len_v = min(pt(0.03125), max_interior_len_v)
    half_w = cut_w_pt / 2
    half_h = cut_h_pt / 2
    x_l = center_x - half_w
    x_r = center_x + half_w
    y_b = center_y - half_h
    y_t = center_y + half_h
    top_len = perimeter_len if is_top_edge else interior_len_v
    bottom_len = perimeter_len if is_bottom_edge else interior_len_v
    left_len = perimeter_len if is_left_edge else interior_len_h
    right_len = perimeter_len if is_right_edge else interior_len_h

    return [
        line_command(x_l, y_t + off, x_l, y_t + off + top_len),
        line_command(x_l - off - left_len, y_t, x_l - off, y_t),
        line_command(x_r, y_t + off, x_r, y_t + off + top_len),
        line_command(x_r + off, y_t, x_r + off + right_len, y_t),
        line_command(x_l, y_b - off, x_l, y_b - off - bottom_len),
        line_command(x_l - off - left_len, y_b, x_l - off, y_b),
        line_command(x_r, y_b - off, x_r, y_b - off - bottom_len),
        line_command(x_r + off, y_b, x_r + off + right_len, y_b),
    ]


def line_command(x1: float, y1: float, x2: float, y2: float) -> str:
    return f"{pdf_num(x1)} {pdf_num(y1)} m {pdf_num(x2)} {pdf_num(y2)} l S"


def pdf_num(value: float) -> str:
    text = f"{value:.4f}"
    text = text.rstrip("0").rstrip(".")
    return text or "0"


def get_effective_position(
    row: int,
    column: int,
    layout: Dict[str, float],
    binding_edge: str,
    flip_positions_this_page: bool,
) -> Tuple[int, int]:
    if not flip_positions_this_page:
        return row, column

    normalized_edge = str(binding_edge or "Left").lower()
    if normalized_edge in ("left", "right"):
        return row, layout["cols"] - 1 - column
    if normalized_edge in ("top", "bottom"):
        return layout["rows"] - 1 - row, column
    return row, column


def get_shift_adjustments(binding_edge: str, flip_positions_this_page: bool, image_shift_x: float, image_shift_y: float) -> Tuple[float, float]:
    if not flip_positions_this_page:
        return image_shift_x or 0.0, image_shift_y or 0.0

    normalized_edge = str(binding_edge or "Left").lower()
    if normalized_edge in ("left", "right"):
        return -(image_shift_x or 0.0), image_shift_y or 0.0
    if normalized_edge in ("top", "bottom"):
        return image_shift_x or 0.0, -(image_shift_y or 0.0)
    return image_shift_x or 0.0, image_shift_y or 0.0


def pre_rotation_shift_for(degrees_value: int, shift_x_in: float, shift_y_in: float) -> Tuple[float, float]:
    shift_x_pt = pt(shift_x_in or 0.0)
    shift_y_pt = pt(shift_y_in or 0.0)
    normalized = normalize_degrees(degrees_value)
    if normalized == 180:
        return -shift_x_pt, -shift_y_pt
    if normalized == 90:
        return -shift_y_pt, shift_x_pt
    if normalized == 270:
        return shift_y_pt, -shift_x_pt
    return shift_x_pt, shift_y_pt


def adjust_xy_for_rotation(x: float, y: float, width: float, height: float, degrees_value: int) -> Tuple[float, float]:
    normalized = normalize_degrees(degrees_value)
    if normalized == 180:
        return x + width, y + height
    if normalized == 90:
        return x + height, y
    if normalized == 270:
        return x, y + width
    return x, y


def rotation_for_page(
    config: ImpositionConfig,
    row: int,
    column: int,
    flip_positions_this_page: bool,
    page_index: int,
) -> int:
    degrees_value = compute_art_rotation_degrees(config, row, column, page_index)
    if flip_positions_this_page:
        degrees_value = (degrees_value + 180) % 360
    return degrees_value


def compute_art_rotation_degrees(config: ImpositionConfig, row: int, column: int, page_index: int) -> int:
    mode = str(config.art_rotation or "None").strip().lower()
    start_rotated = bool(config.rotate_first_column_or_row)
    if mode == "evenpages":
        return 180 if page_index % 2 == 1 else 0
    if mode == "rows":
        return 180 if (row % 2 == 0 and start_rotated) or (row % 2 == 1 and not start_rotated) else 0
    if mode.startswith("col"):
        return 180 if (column % 2 == 0 and start_rotated) or (column % 2 == 1 and not start_rotated) else 0
    return 0


def should_auto_correct_orientation(config: ImpositionConfig, first_page: PageObject) -> bool:
    cut_orientation = orientation_from_dimensions(config.cut_width_in, config.cut_height_in)
    source_width, source_height = get_source_page_size(first_page)
    source_orientation = orientation_from_dimensions(source_width, source_height)
    if cut_orientation == "square" or source_orientation == "square":
        return False
    return cut_orientation != source_orientation


def orientation_from_dimensions(width: float, height: float) -> str:
    if width <= 0 or height <= 0:
        return "square"
    if abs(width - height) <= 0.0001:
        return "square"
    return "portrait" if height > width else "landscape"


def get_source_page_size(page: PageObject) -> Tuple[float, float]:
    crop_box = getattr(page, "cropbox", None)
    if crop_box is not None and crop_box.width > 0 and crop_box.height > 0:
        return float(crop_box.width), float(crop_box.height)
    media_box = getattr(page, "mediabox", None)
    if media_box is not None and media_box.width > 0 and media_box.height > 0:
        return float(media_box.width), float(media_box.height)
    raise ImpositionError("A source page is missing a valid media box.")


def clone_page(page: PageObject, width: float, height: float) -> PageObject:
    cloned_page = PageObject.create_blank_page(width=width, height=height)
    cloned_page.merge_page(page)
    return cloned_page


def build_output_file_name(source_path: str, mode: str, cols: int, rows: int) -> str:
    base_name = os.path.splitext(os.path.basename(source_path))[0]
    safe_base = "".join(ch if ch.isalnum() or ch in "._-" else "-" for ch in base_name).strip("-") or "imposed"
    mode_slug = "cut-stack" if mode == "cutAndStack" else "repeat"
    return f"{safe_base}-{mode_slug}-{cols}x{rows}.pdf"


def pt(inches: float) -> float:
    return float(inches) * PT


def normalize_degrees(value: int) -> int:
    return int((value % 360 + 360) % 360)
