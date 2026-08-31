"""Lossless text grouping for Word table pagination.

These helpers establish semantic boundaries only. They do not estimate page
height, change text, or decide whether a block fits in a particular editor.
"""

from collections.abc import Callable
import math
import unicodedata


def estimate_text_height(text, width_pt, size_pt=12, line_spacing=1.15, paragraph_spacing=1):
    """Approximate packing cost only; never impose a Word row/page height."""
    columns = max(1, width_pt / size_pt)
    lines = 0
    for paragraph in text.split("\n"):
        units = sum(1 if unicodedata.east_asian_width(char) in "WF" else 0.6 for char in paragraph)
        lines += max(1, math.ceil(units / columns))
    return lines * size_pt * line_spacing + len(text.split("\n")) * paragraph_spacing


def balance_cell_blocks(cell_blocks, measure, max_height=160):
    """Pack adjacent atomic blocks into existing tall rows without reordering.

    Output cells are lists of original blocks (empty list means padding). The
    tallest first block provides a local budget; no atomic block is split and
    long blocks never pull additional content into an oversized row.
    """
    cursors = [0] * len(cell_blocks)
    rows = []
    while any(cursor < len(blocks) for cursor, blocks in zip(cursors, cell_blocks)):
        cells = []
        for index, blocks in enumerate(cell_blocks):
            cells.append(blocks[cursors[index]:cursors[index] + 1])
            cursors[index] += bool(cells[-1])
        target = max(measure(index, "\n".join(cell)) for index, cell in enumerate(cells))
        if target > max_height:
            rows.append(cells)
            continue
        for index, blocks in enumerate(cell_blocks):
            while cursors[index] < len(blocks):
                candidate = cells[index] + [blocks[cursors[index]]]
                if measure(index, "\n".join(candidate)) > target:
                    break
                cells[index] = candidate
                cursors[index] += 1
        rows.append(cells)
    return rows


def split_paragraph_blocks(text: str, is_heading: Callable[[str], bool]) -> list[str]:
    """Group consecutive headings with their first following body paragraph.

    Every remaining body paragraph becomes its own block. Empty lines stay in
    the pending block and never end a heading chain. A heading with no following
    body paragraph is retained as a final block; a caller can flag it rather
    than silently inventing or dropping content.

    The result is always nonempty and lossless::

        "\\n".join(split_paragraph_blocks(text, is_heading)) == text

    ``is_heading`` receives each original, nonblank line (without stripping its
    spaces). Blank-only input and trailing blank lines are preserved too.
    """
    if not isinstance(text, str):
        raise TypeError("text must be a string")
    if not callable(is_heading):
        raise TypeError("is_heading must be callable")

    blocks: list[str] = []
    pending: list[str] = []
    for line in text.split("\n"):
        pending.append(line)
        if line.strip() and not is_heading(line):
            blocks.append("\n".join(pending))
            pending = []

    if pending:
        blocks.append("\n".join(pending))
    return blocks
