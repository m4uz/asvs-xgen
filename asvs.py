#!/usr/bin/env python3

import argparse
import csv
import io
import logging
import requests
import xlsxwriter
from collections import defaultdict

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

ASVS_CSV_URLS = {
    4: 'https://raw.githubusercontent.com/OWASP/ASVS/v4.0.3/4.0/docs_en/OWASP%20Application%20Security%20Verification%20Standard%204.0.3-en.csv',
    5: 'https://raw.githubusercontent.com/OWASP/ASVS/v5.0.0/5.0/docs_en/OWASP_Application_Security_Verification_Standard_5.0.0_en.csv',
}
DEFAULT_OUTPUTS = {
    4: 'OWASP-ASVS-4.0.3.xlsx',
    5: 'OWASP-ASVS-5.0.0.xlsx',
}

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "-a",
        "--asvs-version",
        type=int,
        choices=[4, 5],
        default=5,
        help="ASVS version (4 or 5).",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Output .xlsx file path.",
    )
    args = parser.parse_args()

    if args.output is not None and not args.output.lower().endswith(".xlsx"):
        parser.error("Output file must have a .xlsx extension.")

    if args.output is None:
        args.output = DEFAULT_OUTPUTS[args.asvs_version]

    return args

def main(asvs_version: int, output_path: str):
    asvs_csv = download_asvs_csv(asvs_version)
    worksheet_data = prepare_worksheet_data(asvs_csv, asvs_version)
    create_workbook(worksheet_data, output_path)

def download_asvs_csv(asvs_version: int):
    logging.info(f'Downloading ASVS from {ASVS_CSV_URLS[asvs_version]}')

    r = requests.get(ASVS_CSV_URLS[asvs_version])
    r.raise_for_status()
    return r.text

def prepare_worksheet_data(asvs_csv: str, asvs_version: int) -> defaultdict[str, list[list]]:
    """
    Parses ASVS requirements in CVS format and creates a dictionary where the keys correspond to chapter names and
    values to the list of requirements.
    :param asvs_csv: ASVS requirements CSV content
    :return: Dictionary of chapter names to requirements
    """

    logging.info('Preparing worksheet data')

    csv_reader = csv.reader(io.StringIO(asvs_csv))
    worksheets = defaultdict(list)

    next(csv_reader)  # Skip header row

    # Expected headers - version 4:
    # 0 - chapter_id
    # 1 - chapter_name
    # 2 - section_id
    # 3 - section_name
    # 4 - req_id
    # 5 - req_description
    # 6 - level1
    # 7 - level2
    # 8 - level3
    # 9 - cwe

    # Expected headers - version 5:
    # 0 - chapter_id
    # 1 - chapter_name
    # 2 - section_id
    # 3 - section_name
    # 4 - req_id
    # 5 - req_description
    # 6 - L

    for row in csv_reader:
        # Skip empty rows
        if not row:
            continue

        chapter = f'{row[0]} {row[1]}'             # chapter_id + chapter_name
        req_id = row[4][1:]                        # req_id
        section = row[3]                           # section_name
        req = row[5]                               # req_description
        if asvs_version == 4:
            l1 = row[6]                            # level1
            l2 = row[7]                            # level2
            l3 = row[8]                            # level3
        elif asvs_version == 5:
            l1 = '✓' if int(row[6]) <= 1 else ''   # level1
            l2 = '✓' if int(row[6]) <= 2 else ''   # level2
            l3 = '✓' if int(row[6]) <= 3 else ''   # level3
        else:
            raise ValueError('Version must be 4 or 5')

        fulfilled = ''
        comment = ''

        worksheets[chapter].append([
            req_id,
            section,
            req,
            l1,
            l2,
            l3,
            fulfilled,
            comment])

    return worksheets

def create_workbook(worksheets: defaultdict[str, list[list]], output_path: str) -> str:
    """
    Creates ASVS Excel workbook based on the provided data.
    :param worksheets: Dictionary of chapter names to requirements
    :param output_path: Path for the output workbook
    :return: Absolute path to the created workbook
    """

    logging.info('Creating workbook')

    workbook = xlsxwriter.Workbook(output_path)

    # Summary worksheet - Added now so that it is the first one.
    workbook.add_worksheet("Summary")

    # --------------------------------------------------
    # Chapter worksheets
    # --------------------------------------------------
    
    for sheet_name, data in worksheets.items():
        # Worksheet - Sheet name cannot exceed 30 chars
        worksheet = workbook.add_worksheet(sheet_name[:30])

        # Zoom
        worksheet.set_zoom(150)

        # Column widths
        worksheet.set_column('A:A', 7)   # Requirement ID
        worksheet.set_column('B:B', 50)  # Section
        worksheet.set_column('C:C', 50)  # Requirement
        worksheet.set_column('D:D', 10)  # Level 1
        worksheet.set_column('E:E', 10)  # Level 2
        worksheet.set_column('F:F', 10)  # Level 3
        worksheet.set_column('G:G', 10)  # Fulfilled
        worksheet.set_column('H:H', 50)  # Comment

        # Requirements table column formatting
        requirement_id_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
        section_fmt = workbook.add_format({'align': 'left', 'valign': 'vcenter'})
        requirement_fmt = workbook.add_format({'align': 'left', 'valign': 'top', 'text_wrap': True})
        level_fmt_props = {'align': 'center', 'valign': 'vcenter', 'text_wrap': True}
        l1_fmt = workbook.add_format({**level_fmt_props, 'bg_color': '#DCE6F1'})
        l2_fmt = workbook.add_format({**level_fmt_props, 'bg_color': '#B8CCE4'})
        l3_fmt = workbook.add_format({**level_fmt_props, 'bg_color': '#95B3D7'})
        fulfilled_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        comment_fmt = workbook.add_format({'align': 'left', 'valign': 'top', 'text_wrap': True})

        # Requirements table - Table name contains chapter postfix, i.e., table_v1, table_v2, table_v3, ...etc...
        table_name = f'table_{sheet_name.split(" ")[0].lower()}'
        worksheet.add_table(0, 0, len(data), 7, {
            'name': table_name,
            'columns': [
                {'header': 'Requirement ID', 'format': requirement_id_fmt},
                {'header': 'Section', 'format': section_fmt},
                {'header': 'Requirement', 'format': requirement_fmt},
                {'header': 'Level 1', 'format': l1_fmt},
                {'header': 'Level 2', 'format': l2_fmt},
                {'header': 'Level 3', 'format': l3_fmt},
                {'header': 'Fulfilled', 'format': fulfilled_fmt},
                {'header': 'Comment', 'format': comment_fmt},
            ],
            'style': 'Table Style Light 9',
            'data': data})

        # 'Fulfilled' column dropdown selection
        for row in range(1, len(data) + 2):
            worksheet.data_validation(f'G{row}', {
                'validate': 'list',
                'source': ['Yes', 'No', 'Partially', 'Not applicable'],
                'error_message': 'Invalid input. Choose Yes, No, Partially or Not applicable.',
                'error_title': 'Invalid Input',
            })

        # 'Fulfilled' column conditional formatting
        fulfilled_range = f"G2:G{len(data) + 1}"
        worksheet.conditional_format(fulfilled_range, {'type': 'cell',
                                                       'criteria': '==',
                                                       'value': '"Yes"',
                                                       'format': workbook.add_format({'bg_color': '#ECF1DF'})})
        worksheet.conditional_format(fulfilled_range, {'type': 'cell',
                                                       'criteria': '==',
                                                       'value': '"No"',
                                                       'format': workbook.add_format({'bg_color': '#FFC7CE'})})
        worksheet.conditional_format(fulfilled_range, {'type': 'cell',
                                                       'criteria': '==',
                                                       'value': '"Partially"',
                                                       'format': workbook.add_format({'bg_color': '#FFEB9C'})})
        worksheet.conditional_format(fulfilled_range, {'type': 'cell',
                                                       'criteria': '==',
                                                       'value': '"Not applicable"',
                                                       'format': workbook.add_format({'bg_color': '#D3D3D3'})})

    # --------------------------------------------------
    # Summary worksheet
    # --------------------------------------------------

    worksheet = workbook.get_worksheet_by_name('Summary')

    # Zoom
    worksheet.set_zoom(150)

    # Formatting
    heading_format = workbook.add_format({
        "bold": 1,
        "align": "left",
        "valign": "vcenter",
    })

    # Table headers
    summary_headers = [
        'Level',
        'Total',
        'Yes',
        'No',
        'Partially',
        'Not applicable',
        'No Answer',
    ]

    summary_columns = [{'header': header} for header in summary_headers]

    # --------------------------------------------------
    # Summary worksheet - Per-chapter summary tables
    # --------------------------------------------------

    heading_row = 0
    table_first_row = 1
    table_last_row = 4

    def build_level_formulas(target_table: str, level: int) -> list[str]:
        level_name = f'Level {level}'
        return [
            level_name,
            f'=COUNTA({target_table}[{level_name}])',
            f'=COUNTIFS({target_table}[Fulfilled], "Yes", {target_table}[{level_name}], "<>")',
            f'=COUNTIFS({target_table}[Fulfilled], "No", {target_table}[{level_name}], "<>")',
            f'=COUNTIFS({target_table}[Fulfilled], "Partially", {target_table}[{level_name}], "<>")',
            f'=COUNTIFS({target_table}[Fulfilled], "Not applicable", {target_table}[{level_name}], "<>")',
            f'=COUNTIFS({target_table}[Fulfilled], "", {target_table}[{level_name}], "<>")',
        ]

    for sheet_name in worksheets.keys():
        # Chapter heading
        worksheet.write(heading_row, 0, sheet_name, heading_format)
        worksheet.merge_range(heading_row, 0, heading_row, 5, None)

        # Fulfillment statistics per level and chapter:
        target_table = f'table_{sheet_name.split(" ")[0].lower()}'
        data = [
            build_level_formulas(target_table, 1),
            build_level_formulas(target_table, 2),
            build_level_formulas(target_table, 3),
        ]

        worksheet.add_table(table_first_row, 0, table_last_row, 6, {
            'columns': summary_columns,
            'style': 'Table Style Light 9',
            'autofilter': False,
            'data': data,
        })

        heading_row += 5
        table_first_row += 5
        table_last_row += 5

    # --------------------------------------------------
    # Summary worksheet - Summary of all chapters
    # --------------------------------------------------

    row_offset = 5
    data_row_offsets = {
        1: 3,
        2: 4,
        3: 5,
    }

    def rows_for_level(level: int) -> list[int]:
        first_row = data_row_offsets[level]
        return list(range(first_row, len(worksheets) * row_offset, row_offset))

    def sum_column(col_letter: str, rows: list[int]) -> str:
        return f'=SUM({",".join([f"{col_letter}{row}" for row in rows])})'

    formulas = [
        # Fulfillment statistics per level across all chapters:
        [
            'Level 1',
            sum_column("B", rows_for_level(1)),
            sum_column("C", rows_for_level(1)),
            sum_column("D", rows_for_level(1)),
            sum_column("E", rows_for_level(1)),
            sum_column("F", rows_for_level(1)),
            sum_column("G", rows_for_level(1)),
        ],
        [
            'Level 2',
            sum_column("B", rows_for_level(2)),
            sum_column("C", rows_for_level(2)),
            sum_column("D", rows_for_level(2)),
            sum_column("E", rows_for_level(2)),
            sum_column("F", rows_for_level(2)),
            sum_column("G", rows_for_level(2)),
        ],
        [
            'Level 3',
            sum_column("B", rows_for_level(3)),
            sum_column("C", rows_for_level(3)),
            sum_column("D", rows_for_level(3)),
            sum_column("E", rows_for_level(3)),
            sum_column("F", rows_for_level(3)),
            sum_column("G", rows_for_level(3)),
        ],
    ]

    summary_heading_row = len(worksheets) * 5
    summary_table_first_row = summary_heading_row + 1

    # Summary heading
    worksheet.write(summary_heading_row, 0, "Summary", heading_format)
    worksheet.merge_range(summary_heading_row, 0, summary_heading_row, 5, None)

    # Summary table
    worksheet.add_table(summary_table_first_row, 0, summary_table_first_row + 3, 6, {
        'columns': summary_columns,
        'style': 'Table Style Light 9',
        'autofilter': False,
        'data': formulas,
    })

    # --------------------------------------------------
    # Summary worksheet - Fulfillment chart
    # --------------------------------------------------

    chart = workbook.add_chart({"type": "bar", "subtype": "percent_stacked"})
    chart.set_title({"name": "Fulfillment Summary"})
    chart.set_x_axis({"label_position": "none"})

    summary_first_data_row = summary_table_first_row + 1
    summary_last_data_row = summary_table_first_row + len(formulas)

    series_definitions = [
        ("Yes", 2, "#ECF1DF"),
        ("No", 3, "#FFC7CE"),
        ("Partially", 4, "#FFEB9C"),
        ("Not applicable", 5, "#D9D9D9"),
        ("No Answer", 6, "#F2F2F2"),
    ]

    for _, column_index, color in series_definitions:
        chart.add_series({
            "name": ["Summary", summary_table_first_row, column_index],
            "categories": ["Summary", summary_first_data_row, 0, summary_last_data_row, 0],
            "values": ["Summary", summary_first_data_row, column_index, summary_last_data_row, column_index],
            "fill": {"color": color},
            "border": {"color": color},
            "data_labels": {"value": True},
        })

    chart_start_cell = "I2"
    worksheet.insert_chart(chart_start_cell, chart)

    workbook.close()

    logging.info(f'Workbook saved as {output_path}')

if __name__ == '__main__':
    args = parse_args()
    main(args.asvs_version, args.output)
