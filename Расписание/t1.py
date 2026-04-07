#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
docx_to_html.py
----------------
Convert a Microsoft Word (.docx) document to an HTML file while preserving
tables (including merged cells, borders, and basic styling).

Requirements
------------
pip install mammoth tqdm beautifulsoup4

Usage
-----
python docx_to_html.py input.docx output.html
"""

import sys
import pathlib
from typing import Tuple

import mammoth
from tqdm import tqdm
from bs4 import BeautifulSoup


def _add_border_css(html: str) -> str:
    """
    Insert a <style> block that:
        * draws visible borders for all tables,
        * centers the text inside every table cell,
        * keeps the original cell padding.
    The block is placed right after </ opening <head> tag (if present)
    or at the very beginning of the document.
    """
    css = """
    <style>
        /* Make tables look like Word tables */
        table {
            border-collapse: collapse;
            width: auto;               /* keep original width */
        }
        th, td {
            border: 1px solid #000;    /* solid 1 px black border */
            padding: 4px;              /* default cell padding */
            text-align: center;        /* <-- центрируем текст */
        }
    </style>
    """

    # Insert after <head> if it exists, otherwise prepend to the document.
    if "<head>" in html:
        html = html.replace("</head>", f"{css}\n</head>")
    else:
        html = css + "\n" + html
    return html


def convert_docx_to_html(
    docx_path: pathlib.Path,
    html_path: pathlib.Path,
    embed_images: bool = True,
    image_dir: pathlib.Path | None = None,
) -> Tuple[int, int]:
    """
    Convert a DOCX file to HTML.

    Parameters
    ----------
    docx_path : pathlib.Path
        Path to the source .docx file.
    html_path : pathlib.Path
        Destination path for the generated HTML file.
    embed_images : bool, default=True
        If True, images are embedded as base64 data URIs.
        If False, images are saved to ``image_dir`` and referenced by <img src>.
    image_dir : pathlib.Path | None, default=None
        Directory where extracted images will be stored when ``embed_images`` is False.
        If ``None`` and ``embed_images`` is False, a folder named ``<html_file>_files``
        will be created next to the HTML output.

    Returns
    -------
    Tuple[int, int]
        (number_of_paragraphs, number_of_tables) processed.
    """
    # ------------------------------------------------------------------ #
    # 1️⃣  Prepare image handling callbacks
    # ------------------------------------------------------------------ #
    if not embed_images:
        if image_dir is None:
            image_dir = html_path.with_suffix("").with_name(html_path.stem + "_files")
        image_dir.mkdir(parents=True, exist_ok=True)

        def convert_image(image):
            # Save the image file and return a relative URL.
            image_extension = image.content_type.split("/")[-1]  # e.g. "png"
            image_name = f"image_{image.index}.{image_extension}"
            image_path = image_dir / image_name
            with open(image_path, "wb") as f:
                f.write(image.binary)
            return {"src": str(image_path.relative_to(html_path.parent))}

        image_converter = mammoth.images.img_elem(convert_image)
    else:
        # Embed as base64 – no extra folder needed.
        image_converter = mammoth.images.inline

    # ------------------------------------------------------------------ #
    # 2️⃣  Convert the document
    # ------------------------------------------------------------------ #
    with open(docx_path, "rb") as docx_file:
        # OPTIONAL: you can add a style_map to keep more Word styles.
        # For borders we rely on the CSS injected later, but you could also
        # map specific Word table styles to inline CSS, e.g.:
        #   "table[style-name='MyTable'] => table:style='border:2px solid #ff0000;'" 
        result = mammoth.convert_to_html(
            docx_file,
            convert_image=image_converter,
            # style_map = "p[style-name='Heading 1'] => h1:f:f"
        )
        html = result.value  # The generated HTML
        messages = result.messages  # Any conversion warnings

    # ------------------------------------------------------------------ #
    # 3️⃣  Inject CSS for table borders **and** centered text
    # ------------------------------------------------------------------ #
    html = _add_border_css(html)

    # ------------------------------------------------------------------ #
    # 4️⃣  Write output
    # ------------------------------------------------------------------ #
    html_path.write_text(html, encoding="utf-8")

    # ------------------------------------------------------------------ #
    # 5️⃣  Simple statistics (optional)
    # ------------------------------------------------------------------ #
    soup = BeautifulSoup(html, "html.parser")
    n_paragraphs = len(soup.find_all("p"))
    n_tables = len(soup.find_all("table"))

    # Print any conversion messages (useful for debugging)
    if messages:
        sys.stderr.write("Conversion messages:\n")
        for m in messages:
            sys.stderr.write(f"  - {m}\n")

    return n_paragraphs, n_tables

def converter(docx_file, html_file):

    input_path = pathlib.Path(docx_file).expanduser().resolve()
    output_path = pathlib.Path(html_file).expanduser().resolve()

    if not input_path.is_file():
        sys.stderr.write(f"❌ Input file not found: {input_path}\n")
        sys.exit(1)

    # Progress bar for large documents (optional)
    with tqdm(total=1, desc="Converting", unit="file") as pbar:
        n_par, n_tbl = convert_docx_to_html(
            input_path, output_path, embed_images=True
        )
        pbar.update(1)

    print(f"✅  Conversion finished → {output_path}")
    print(f"   Paragraphs: {n_par}, Tables: {n_tbl}")


def main(argv=None):
    if argv is None:
        argv = sys.argv[1:]

    if len(argv) not in (2, 3):
        sys.stderr.write(
            "Usage: python docx_to_html.py <input.docx> <output.html> [--no-embed]\n"
        )
        sys.exit(1)

    input_path = pathlib.Path(argv[0]).expanduser().resolve()
    output_path = pathlib.Path(argv[1]).expanduser().resolve()
    embed = True
    if len(argv) == 3 and argv[2] == "--no-embed":
        embed = False

    if not input_path.is_file():
        sys.stderr.write(f"❌ Input file not found: {input_path}\n")
        sys.exit(1)

    # Progress bar for large documents (optional)
    with tqdm(total=1, desc="Converting", unit="file") as pbar:
        n_par, n_tbl = convert_docx_to_html(
            input_path, output_path, embed_images=embed
        )
        pbar.update(1)

    print(f"✅  Conversion finished → {output_path}")
    print(f"   Paragraphs: {n_par}, Tables: {n_tbl}")


if __name__ == "__main__":
    main()
