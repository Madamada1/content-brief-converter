"""Tests for product-grid / product-comparison detection and rendering:
repeated "Myer product URL + name + metadata" records in a brief's Draft
Content should become one row per product in a PRODUCT GRID block (with
the URL only in the LINK column), not one flattened blob of copy.

See docx_parser._split_product_entries (detection) and
xlsx_builder._product_grid_block (rendering).
"""

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx_helpers import build_brief_docx

from docx_parser import parse_docx
from xlsx_builder import COL_LINK, COL_HEADLINE, COL_SUBCOPY, build_workbook

KETTLES = [
    (
        "https://www.myer.com.au/p/breville-the-soft-top-luxe-kettle-bke735nre4ian1",
        "Breville Soft Top Luxe Kettle",
        [
            "Best for: Busy households",
            "Pros: Rapid boil and premium finish",
            "Cons: Fixed temperature",
            "Star rating: 5/5",
            "Cost: $229 RRP",
        ],
    ),
    (
        "https://www.myer.com.au/p/smeg-klf01blau-50s-style-kettle-black",
        "Smeg 50's Style Kettle in Black",
        [
            "Best for: Style-led kitchens",
            "Pros: Iconic look and smooth pouring",
            "Cons: Weightier build",
            "Star rating: 4.8/5",
            "Cost: $299 RRP",
        ],
    ),
    (
        "https://www.myer.com.au/p/russell-hobbs-legacy-kettle-rhk52aus",
        "Russell Hobbs Legacy Kettle",
        [
            "Best for: Classic kitchens",
            "Pros: Retro design and quiet boil",
            "Cons: No temperature control",
            "Star rating: 4.5/5",
            "Cost: $89 RRP",
        ],
    ),
]


class ProductGridDetectionTest(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmpdir.cleanup)

    def _path(self, name):
        return os.path.join(self.tmpdir.name, name)

    def test_plain_text_url_products_detected_as_separate_records(self):
        """The reported bug: a generated doc that writes the URL and name
        as plain (non-bulleted, non-bold) paragraphs must still be split
        into one product per record, not left as one flattened blob."""
        path = self._path("plain.docx")
        build_brief_docx(
            path,
            heading="Best Kettles",
            intro_lines=["Discover our favourite kettles for every kitchen."],
            products=KETTLES,
            bullet_url=False,
            bold_name=False,
        )
        parsed = parse_docx(path)
        section = parsed["sections"][0]

        self.assertEqual(section["heading"], "Best Kettles")
        self.assertEqual(
            section["copy"].strip(), "Discover our favourite kettles for every kitchen."
        )
        self.assertEqual(len(section["products"]), len(KETTLES))

        for (url, name, detail_lines), product in zip(KETTLES, section["products"]):
            self.assertEqual(product["url"], url)
            self.assertEqual(product["name"], name)
            # The URL must not leak into the product's copy text.
            self.assertNotIn(url, product["copy"])
            for line in detail_lines:
                self.assertIn(line, product["copy"])

    def test_bulleted_hyperlink_bold_name_products_still_detected(self):
        """The original supported shape (bulleted hyperlink + bold name)
        must keep working."""
        path = self._path("bulleted.docx")
        build_brief_docx(
            path,
            heading="Best Kettles",
            intro_lines=["Discover our favourite kettles for every kitchen."],
            products=KETTLES,
            bullet_url=True,
            bold_name=True,
        )
        parsed = parse_docx(path)
        section = parsed["sections"][0]
        self.assertEqual(len(section["products"]), len(KETTLES))
        for (url, name, _), product in zip(KETTLES, section["products"]):
            self.assertEqual(product["url"], url)
            self.assertEqual(product["name"], name)

    def test_non_product_section_unaffected(self):
        """Ordinary H2/body-copy sections with no Myer /p/ URL must be
        left exactly as-is: no products, full copy intact."""
        path = self._path("normal.docx")
        build_brief_docx(
            path,
            heading="Mattress Sizes",
            intro_lines=[
                "Choosing the right mattress size depends on your room and who's sleeping in it.",
                "Best for: Kids, teenagers and adults in smaller rooms",
            ],
            products=[],
        )
        parsed = parse_docx(path)
        section = parsed["sections"][0]
        self.assertEqual(section["products"], [])
        self.assertIn("Best for: Kids, teenagers and adults in smaller rooms", section["copy"])

    def test_unrelated_link_in_body_copy_is_not_a_delimiter(self):
        """A non-product-page link mentioned inline in ordinary body copy
        must not be mistaken for a product-grid marker."""
        path = self._path("unrelated_link.docx")
        build_brief_docx(
            path,
            heading="Shop Kettles",
            intro_lines=[
                "Browse the full range at https://www.myer.com.au/c/kettles for more options.",
            ],
            products=[],
        )
        parsed = parse_docx(path)
        section = parsed["sections"][0]
        self.assertEqual(section["products"], [])
        self.assertIn("myer.com.au/c/kettles", section["copy"])

    def test_products_end_at_next_heading(self):
        """A product grid must not swallow content from the next H2/H3
        section."""
        path = self._path("trailing.docx")
        build_brief_docx(
            path,
            heading="Best Kettles",
            intro_lines=["Discover our favourite kettles for every kitchen."],
            products=KETTLES[:2],
            trailing_heading="Care Guide",
            trailing_lines=["Descale your kettle every three months."],
        )
        parsed = parse_docx(path)
        self.assertEqual(len(parsed["sections"]), 2)
        kettles_section, care_section = parsed["sections"]
        self.assertEqual(len(kettles_section["products"]), 2)
        self.assertEqual(care_section["products"], [])
        self.assertIn("Descale your kettle every three months.", care_section["copy"])


class ProductGridExcelOutputTest(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmpdir.cleanup)

    def test_each_product_gets_its_own_row_with_url_in_link_column(self):
        path = os.path.join(self.tmpdir.name, "kettles.docx")
        build_brief_docx(
            path,
            heading="Best Kettles",
            intro_lines=["Discover our favourite kettles for every kitchen."],
            products=KETTLES,
        )
        parsed = parse_docx(path)
        parsed["source_filename"] = "kettles.docx"
        wb = build_workbook(parsed)
        ws = wb.active

        # Find the PRODUCT GRID divider row.
        divider_row = next(
            r for r in range(1, ws.max_row + 1) if ws.cell(row=r, column=1).value == "PRODUCT GRID"
        )

        product_rows = []
        r = divider_row + 1
        while r <= ws.max_row and ws.cell(row=r, column=COL_LINK).value:
            product_rows.append(r)
            r += 1

        # One Excel row per detected product -- not one row for the whole grid.
        self.assertEqual(len(product_rows), len(KETTLES))

        for row, (url, name, detail_lines) in zip(product_rows, KETTLES):
            self.assertEqual(ws.cell(row=row, column=COL_LINK).value, url)
            self.assertEqual(ws.cell(row=row, column=COL_HEADLINE).value, name)
            copy_cell = ws.cell(row=row, column=COL_SUBCOPY).value
            copy_text = str(copy_cell)
            # The raw product URL must not appear in the copy cell.
            self.assertNotIn(url, copy_text)
            for line in detail_lines:
                self.assertIn(line, copy_text)


if __name__ == "__main__":
    unittest.main()
