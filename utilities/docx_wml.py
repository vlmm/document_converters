from __future__ import annotations

from dataclasses import dataclass

# Namespace map for WordprocessingML used by lxml xpath() calls.
# Without this, expressions like .//w:br can raise: "Undefined namespace prefix".
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def xpath(el, expr: str):
    """Run xpath with the WordprocessingML namespace map.

    Some python-docx elements expose an lxml element with an xpath() method that
    requires a namespace mapping when prefixes are used.
    """
    try:
        return el.xpath(expr, namespaces=NS)
    except TypeError:
        # Some xpath implementations don't accept the namespaces kwarg.
        return el.xpath(expr)


def paragraph_contains_page_break(p) -> bool:
    try:
        # Use explicit namespaces to avoid "Undefined namespace prefix".
        return bool(xpath(p._p, './/w:br[@w:type="page"]'))
    except Exception:
        return False


def extract_numpr(p) -> tuple[bool, int, int]:
    """Return (has_numpr, numId, ilvl)."""
    try:
        pPr = p._p.pPr
        if pPr is None or pPr.numPr is None:
            return False, 0, 0
        numPr = pPr.numPr
        ilvl_elem = numPr.ilvl
        ilvl_val = (
            int(ilvl_elem.val)
            if ilvl_elem is not None and ilvl_elem.val is not None
            else 0
        )
        numId_elem = numPr.numId
        num_id_val = (
            int(numId_elem.val)
            if numId_elem is not None and numId_elem.val is not None
            else 0
        )
        if num_id_val <= 0:
            return False, 0, 0
        return True, num_id_val, ilvl_val
    except Exception:
        return False, 0, 0


@dataclass
class ListFormat:
    kind: str  # "decimal" | "upperRoman" | "lowerRoman" | "upperLetter" | "lowerLetter" | "bullet" | "unknown"
    lvl_text: str  # e.g. "%1." or "%1)"
    start: int = 1


def build_list_formats(doc) -> dict[tuple[int, int], ListFormat]:
    """
    Best-effort parse of numbering definitions for list marker formats.
    """
    formats: dict[tuple[int, int], ListFormat] = {}
    try:
        numbering = doc.part.numbering_part.element  # CT_Numbering
    except Exception:
        return formats

    # Map abstractNumId -> {ilvl -> lvl element}
    abstract_lvls: dict[int, dict[int, object]] = {}
    for abs_el in xpath(numbering, "./w:abstractNum"):
        try:
            abs_id = int(abs_el.get(f"{{{W_NS}}}abstractNumId"))
        except Exception:
            continue
        lvl_map: dict[int, object] = {}
        for lvl in xpath(abs_el, "./w:lvl"):
            try:
                ilvl = int(lvl.get(f"{{{W_NS}}}ilvl"))
            except Exception:
                continue
            lvl_map[ilvl] = lvl
        abstract_lvls[abs_id] = lvl_map

    # numId -> abstractNumId
    num_to_abs: dict[int, int] = {}
    # numId overrides: (numId, ilvl) -> startOverride
    overrides_start: dict[tuple[int, int], int] = {}

    for num in xpath(numbering, "./w:num"):
        try:
            num_id = int(num.get(f"{{{W_NS}}}numId"))
        except Exception:
            continue

        abs_ref = xpath(num, "./w:abstractNumId")
        if abs_ref:
            try:
                abs_id = int(abs_ref[0].get(f"{{{W_NS}}}val"))
                num_to_abs[num_id] = abs_id
            except Exception:
                pass

        for lvl_override in xpath(num, "./w:lvlOverride"):
            try:
                ilvl = int(lvl_override.get(f"{{{W_NS}}}ilvl"))
            except Exception:
                continue
            start_ov = xpath(lvl_override, "./w:startOverride")
            if start_ov:
                try:
                    start_val = int(start_ov[0].get(f"{{{W_NS}}}val"))
                    overrides_start[(num_id, ilvl)] = start_val
                except Exception:
                    pass

    def _kind_from_numfmt(numfmt: str) -> str:
        m = (numfmt or "").lower()
        if m in ("decimal",):
            return "decimal"
        if m in ("upperroman",):
            return "upperRoman"
        if m in ("lowerroman",):
            return "lowerRoman"
        if m in ("upperletter", "upperalpha"):
            return "upperLetter"
        if m in ("lowerletter", "loweralpha"):
            return "lowerLetter"
        if m in ("bullet",):
            return "bullet"
        return "unknown"

    # Build (numId, ilvl) -> ListFormat
    for num_id, abs_id in num_to_abs.items():
        lvl_map = abstract_lvls.get(abs_id, {})
        for ilvl, lvl in lvl_map.items():
            numfmt_el = xpath(lvl, "./w:numFmt")
            lvltext_el = xpath(lvl, "./w:lvlText")
            start_el = xpath(lvl, "./w:start")

            numfmt = numfmt_el[0].get(f"{{{W_NS}}}val") if numfmt_el else ""
            lvltext = lvltext_el[0].get(f"{{{W_NS}}}val") if lvltext_el else "%1."
            start = 1
            if start_el:
                try:
                    start = int(start_el[0].get(f"{{{W_NS}}}val"))
                except Exception:
                    start = 1

            if (num_id, ilvl) in overrides_start:
                start = overrides_start[(num_id, ilvl)]

            formats[(num_id, ilvl)] = ListFormat(
                kind=_kind_from_numfmt(numfmt),
                lvl_text=lvltext,
                start=start,
            )

    return formats
