"""Модул за комбиниране и разделяне на PDF файлове.

Поддържа:
- Обединяване на множество PDF файлове в един
- Разделяне на PDF файл по отделни страници или диапазони от страници
- Използване от командния ред (merge / split подкоманди)
"""

import argparse
import re
import sys
from pathlib import Path
from typing import List, Optional

try:
    from pypdf import PdfReader, PdfWriter  # type: ignore
except ImportError:  # pragma: no cover - clear message if dependency missing
    PdfReader = None  # type: ignore
    PdfWriter = None  # type: ignore


def _check_dependency() -> None:
    """Проверява дали pypdf е инсталирано; хвърля RuntimeError ако не е."""
    if PdfReader is None:
        raise RuntimeError(
            "Липсваща зависимост: pypdf.\n"
            "Инсталирайте я с: pip install pypdf"
        )


def merge_pdfs(input_paths: List[Path], output_path: Path) -> Path:
    """Обединява множество PDF файлове в един изходен файл.

    Args:
        input_paths: Наредена последователност от пътища до входните PDF файлове.
        output_path: Път до изходния PDF файл.

    Returns:
        Пътят до създадения изходен файл.

    Raises:
        RuntimeError: Ако pypdf не е инсталирано.
        FileNotFoundError: Ако някой от входните файлове не съществува.
        ValueError: Ако списъкът с входни файлове е празен.
    """
    _check_dependency()

    if not input_paths:
        raise ValueError("Необходим е поне един входен PDF файл.")

    for path in input_paths:
        if not path.exists():
            raise FileNotFoundError(f"Входен файл не е намерен: {path}")

    writer = PdfWriter()

    for path in input_paths:
        reader = PdfReader(str(path))
        for page in reader.pages:
            writer.add_page(page)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "wb") as fout:
        writer.write(fout)

    return output_path


def _parse_page_ranges(spec: str, total_pages: int) -> List[int]:
    """Разбира низ с диапазони от страници и връща списък с 0-базирани индекси.

    Форматът е запетайно-разделени диапазони/единични номера (1-базирани).
    Примери: ``"1"`` → [0], ``"1-3"`` → [0,1,2], ``"1,3-5"`` → [0,2,3,4].

    Args:
        spec: Низ с диапазони (напр. ``"1-3,5,7-9"``).
        total_pages: Общ брой на страниците в документа.

    Returns:
        Списък с уникални 0-базирани индекси на страниците, запазвайки реда.

    Raises:
        ValueError: При невалиден формат или номера извън допустимия диапазон.
    """
    indices: List[int] = []
    seen = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            bounds = part.split("-", 1)
            try:
                start = int(bounds[0].strip())
                end = int(bounds[1].strip())
            except ValueError:
                raise ValueError(f"Невалиден диапазон от страници: '{part}'")
            if start < 1 or end < 1:
                raise ValueError(
                    f"Номерата на страниците трябва да са положителни: '{part}'"
                )
            if start > end:
                raise ValueError(
                    f"Началната страница трябва да е ≤ крайната: '{part}'"
                )
            if end > total_pages:
                raise ValueError(
                    f"Страница {end} е извън документа ({total_pages} стр.)"
                )
            for i in range(start - 1, end):
                if i not in seen:
                    indices.append(i)
                    seen.add(i)
        else:
            try:
                page_num = int(part)
            except ValueError:
                raise ValueError(f"Невалиден номер на страница: '{part}'")
            if page_num < 1:
                raise ValueError(
                    f"Номерата на страниците трябва да са положителни: '{part}'"
                )
            if page_num > total_pages:
                raise ValueError(
                    f"Страница {page_num} е извън документа ({total_pages} стр.)"
                )
            if (page_num - 1) not in seen:
                indices.append(page_num - 1)
                seen.add(page_num - 1)
    return indices


def split_pdf(
    input_path: Path,
    output_dir: Path,
    pages: Optional[str] = None,
) -> List[Path]:
    """Разделя PDF файл по отделни страници или по зададени диапазони.

    Когато ``pages`` не е зададен, всяка страница се записва в отделен файл
    (``<stem>_page_1.pdf``, ``<stem>_page_2.pdf``, …).

    Когато ``pages`` е зададен, само зоните от там се извличат в *един* файл.
    Форматът е запетайно-разделени диапазони (1-базирани), напр. ``"1-3,5"``.

    Args:
        input_path: Път до входния PDF файл.
        output_dir: Директория, в която да се запишат изходните файлове.
        pages: По желание – диапазон от страници за извличане (напр. ``"2-4"``).

    Returns:
        Списък с пътища до всички създадени изходни файлове.

    Raises:
        RuntimeError: Ако pypdf не е инсталирано.
        FileNotFoundError: Ако входният файл не съществува.
        ValueError: При невалиден формат на страниците.
    """
    _check_dependency()

    if not input_path.exists():
        raise FileNotFoundError(f"Входен файл не е намерен: {input_path}")

    reader = PdfReader(str(input_path))
    total_pages = len(reader.pages)
    stem = input_path.stem
    output_dir.mkdir(parents=True, exist_ok=True)

    output_paths: List[Path] = []

    if pages is not None:
        # Extract the specified pages into a single output file
        indices = _parse_page_ranges(pages, total_pages)
        writer = PdfWriter()
        for idx in indices:
            writer.add_page(reader.pages[idx])
        # Build a safe filename from the sanitised page spec (keep only digits, commas and hyphens)
        safe_spec = re.sub(r"[^\d,\-]", "", pages.replace(" ", ""))
        out_file = output_dir / f"{stem}_pages_{safe_spec}.pdf"
        with open(out_file, "wb") as fout:
            writer.write(fout)
        output_paths.append(out_file)
    else:
        # One file per page
        for idx in range(total_pages):
            writer = PdfWriter()
            writer.add_page(reader.pages[idx])
            out_file = output_dir / f"{stem}_page_{idx + 1}.pdf"
            with open(out_file, "wb") as fout:
                writer.write(fout)
            output_paths.append(out_file)

    return output_paths


def main(argv: Optional[List[str]] = None) -> int:
    """Главна функция за командния ред.

    Поддържа две подкоманди:

    * ``merge`` – обединява PDF файлове::

        pdf_utils.py merge -o combined.pdf file1.pdf file2.pdf

    * ``split`` – разделя PDF файл::

        pdf_utils.py split input.pdf -o output_dir/
        pdf_utils.py split input.pdf -o output_dir/ --pages 1-3,5
    """
    parser = argparse.ArgumentParser(
        description="Комбиниране и разделяне на PDF файлове."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # --- merge subcommand ---
    merge_parser = subparsers.add_parser(
        "merge", help="Обединява множество PDF файлове в един."
    )
    merge_parser.add_argument(
        "inputs",
        nargs="+",
        metavar="INPUT",
        help="Входни PDF файлове (по ред).",
    )
    merge_parser.add_argument(
        "-o", "--output",
        required=True,
        metavar="OUTPUT",
        help="Изходен PDF файл.",
    )

    # --- split subcommand ---
    split_parser = subparsers.add_parser(
        "split", help="Разделя PDF файл на отделни файлове."
    )
    split_parser.add_argument("input", help="Входен PDF файл.")
    split_parser.add_argument(
        "-o", "--output-dir",
        default=".",
        metavar="DIR",
        help="Директория за изходните файлове (по подразбиране: текущата).",
    )
    split_parser.add_argument(
        "--pages",
        default=None,
        metavar="RANGES",
        help=(
            "Диапазони от страници за извличане (напр. '1-3,5'). "
            "Ако не е зададено, всяка страница се записва поотделно."
        ),
    )

    args = parser.parse_args(argv)

    try:
        if args.command == "merge":
            output = merge_pdfs(
                [Path(p).expanduser().resolve() for p in args.inputs],
                Path(args.output).expanduser().resolve(),
            )
            print(f"Записан обединен PDF: {output}")

        else:  # split
            outputs = split_pdf(
                Path(args.input).expanduser().resolve(),
                Path(args.output_dir).expanduser().resolve(),
                pages=args.pages,
            )
            for out in outputs:
                print(f"Записан: {out}")

    except Exception as exc:  # pragma: no cover - CLI error path
        sys.stderr.write(f"Грешка: {exc}\n")
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
