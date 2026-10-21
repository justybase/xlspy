"""Create and update a review template. No optional dependencies required."""

import argparse
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from xlspy import (  # noqa: E402
    Comment, DefinedName, ExcelReader, Formula, Protection, SheetProtection,
    Style, WorkbookProtection, XlsxUpdater, XlsxWriter, XlsbUpdater, XlsbWriter,
)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", type=Path, default=Path("reviewed.xlsx"))
    args = parser.parse_args()
    binary = args.out.suffix.lower() == ".xlsb"
    writer_class = XlsbWriter if binary else XlsxWriter
    updater_class = XlsbUpdater if binary else XlsxUpdater
    with writer_class(args.out) as writer:
        writer.add_sheet("Data")
        writer.write_sheet(
            [["Amount", "Tax"], [120, Formula("A2*0.23", value=27.6)]],
            comments={"A2": Comment("Confirm this amount", "Jan", 240, 100)},
            styles={"A2": Style(protection=Protection(locked=False))},
            sheet_protection=SheetProtection(password="review"),
        )
        writer.set_defined_name(DefinedName("Amount", "'Data'!$A$2"))
        writer.set_workbook_protection(WorkbookProtection(password="review"))
    updater = updater_class(args.out)
    try:
        updater.update_cells("Data", {"A2": 125, "B2": Formula("A2*0.23", value=28.75)})
        updater.set_comments("Data", {"A2": Comment("Confirmed", "Ola", 240, 100)})
        updater.save()
    finally:
        updater.close()
    with ExcelReader(str(args.out)) as reader:
        print(reader.get_comments("Data"))
        print(reader.get_defined_names())
        print(reader.read_all("Data", data_only=False))
    print(f"Written: {args.out.resolve()}")


if __name__ == "__main__":
    main()
