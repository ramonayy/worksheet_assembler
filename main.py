from pathlib import Path
import csv

from openpyxl import Workbook


def create_excel_from_csvs(
    data_path: Path,
    csv_pattern: str,
    output_filename: str
):
    output_path = data_path / output_filename

    workbook = Workbook(write_only=True)

    csv_files = sorted(data_path.glob(csv_pattern))

    for csv_file in csv_files:
        sheet_name = csv_file.stem[:31]
        worksheet = workbook.create_sheet(title=sheet_name)

        with csv_file.open(
            "r",
            encoding="utf-8-sig",
            newline=""
        ) as file:
            reader = csv.reader(file, delimiter=",")

            for row_number, row in enumerate(reader):
                if row_number == 0:
                    worksheet.append(row)
                else:
                    converted_row = []
                    for value in row:
                        try:
                            converted_value = int(value)
                        except ValueError:
                            converted_value = float(value)
                        converted_row.append(converted_value)
                    worksheet.append(converted_row)
                row_number += 1

    workbook.save(output_path)

    print(f"Erstellt: {output_path}")


def main():
    data_path = Path(
        r"C:\Users\ma1021716\IdeaProjects\RoulettePlus\data"
    )

    # Alle spieler_*.csv → spieler.xlsx
    create_excel_from_csvs(
        data_path,
        "spieler*.csv",
        "spieler.xlsx"
    )

    # Alle quoten_*.csv → quoten.xlsx
    create_excel_from_csvs(
        data_path,
        "quoten*.csv",
        "quoten.xlsx"
    )


if __name__ == "__main__":
    main()