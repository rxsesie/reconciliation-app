from openpyxl import load_workbook, Workbook
from openpyxl.cell.cell import Cell
import datetime
from rapidfuzz import fuzz, utils
from dataclasses import dataclass
from typing import List, Union, Dict, Tuple
import os
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

CURRENT_YEAR = 2025
STATEMENT_MONTH = 2
client_list_path = "Client List.xlsx"
statements_list = [
    "BCBS- JANUARY- 2025- DIANELKYS (1).xlsx",
    "BCBS- FEBRERO-2025- DIANELKYS (1).xlsx",
]


@dataclass
class BCBSRecord:
    member_name: str
    effective_date: str
    members: Union[int, str]  # number of members
    payment_period: str
    commission: str
    state: str
    action: str


@dataclass
class ClientRecord:
    effective_date: datetime.datetime
    policy_name: str
    members: int
    company: str
    state: str
    members_list: List[str]


@dataclass
class Payment:
    action: str
    payment_period: str
    commission: float


@dataclass
class AllRecord:
    member_name: str
    effective_date: str
    state: str
    company: str
    in_client_list: bool
    payments: List[Payment]


def read_bcbs(file_path="BCBS.xlsx") -> List[BCBSRecord]:
    """Read BCBS.xlsx file and extract relevant commission data."""

    if not os.path.exists(file_path):
        print(f"Error: File {file_path} not found.")
        return []

    wb = load_workbook(filename=file_path, read_only=True)
    sheet = wb.active
    records = []

    for row in sheet.iter_rows(values_only=True):
        if not any(row):  # Skip completely empty rows
            continue
        if row[1] == "CANCELLATIONS":
            print("CCANCELLATIONS")
            break  # skip concellations
        try:
            # Check if this is a client data row (has a numeric ID in column B)
            int(row[1])

            # Extract relevant columns
            policy_name = row[8]
            effective_date = (
                row[14].strftime("%m/%d/%Y")
                if isinstance(row[14], datetime.datetime)
                else row[14]
            )
            members = row[34]
            payment_period = row[41]
            commission = round(row[64], 2)

            # Create record
            record = BCBSRecord(
                member_name=policy_name,
                effective_date=effective_date,
                members=members,
                payment_period=payment_period,
                commission=commission,
                state="",
                action="",
            )
            records.append(record)
        except Exception as e:
            # This is likely a header row or other non-data row
            # print(e)
            row = [r for r in row if r]
            continue

    return records


def read_client_list(file_path="Client List.xlsx") -> List[ClientRecord]:
    """Read client list from Excel and create structured records."""

    if not os.path.exists(file_path):
        print(f"Error: File {file_path} not found.")
        return []

    wb = load_workbook(filename=file_path, read_only=True)
    sheet = wb.active
    records = []

    # Skip the header row
    rows = list(sheet.iter_rows(values_only=True))[1:]

    for row in rows:
        if not any(row):  # Skip completely empty rows
            continue

        effective_date = row[0]
        policy_name = row[1]
        members_count = row[2]
        company = row[3]
        state = row[4]
        client_apply = row[5]
        spouse_name = row[6]
        spouse_apply = row[8]

        # Get all potential members
        member_1 = row[9] if len(row) > 9 else None
        member_1_apply = row[11] if len(row) > 11 else None
        member_2 = row[12] if len(row) > 12 else None
        member_2_apply = row[14] if len(row) > 14 else None
        member_3 = row[15] if len(row) > 15 else None
        member_3_apply = row[17] if len(row) > 17 else None

        # Compile list of active members
        members_list = []
        if client_apply == "SI" and policy_name:
            members_list.append(policy_name)
        if spouse_apply == "SI" and spouse_name:
            members_list.append(spouse_name)
        if member_1_apply == "SI" and member_1:
            members_list.append(member_1)
        if member_2_apply == "SI" and member_2:
            members_list.append(member_2)
        if member_3_apply == "SI" and member_3:
            members_list.append(member_3)

        record = ClientRecord(
            effective_date=effective_date,
            policy_name=policy_name,
            members=members_count or 0,
            company=company,
            state=state,
            members_list=members_list,
        )
        records.append(record)
    return records


def parse_payment_period(period_str: str) -> Tuple[int, int]:
    """Parse payment period string into (month, year) tuple."""
    if len(period_str) == 6:  # Format: YYYYMM
        year = int(period_str[:4])
        month = int(period_str[4:])
    else:  # Try other format
        parts = period_str.split("/")
        if len(parts) == 2:
            month, year = int(parts[0]), int(parts[1])
        else:
            # Default to January 2025 if format is unknown
            month, year = 1, 2025
            print(f"Warning: Unknown payment period format: {period_str}")

    return month, year


def parse_effective_date(date_str: str) -> datetime.datetime:
    """Parse effective date string into datetime object."""
    if isinstance(date_str, datetime.datetime):
        return date_str

    if not date_str:
        return None

    # Try various formats
    formats = [
        "%m/%d/%Y",  # 01/01/2025
        "%m/%d/%Y",  # 01/01/2025
        "%Y-%m-%d",  # 2025-01-01
        "%m/%d/%y",  # 01/01/25
    ]

    for fmt in formats:
        try:
            return datetime.datetime.strptime(date_str, fmt)
        except ValueError:
            continue

    # If we can't parse, return None and print warning
    print(f"Warning: Could not parse date: {date_str}")
    return None


def apply_excel_formatting(ws):
    """Apply formatting to the Excel worksheet."""
    # Define styles
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(
        start_color="336699", end_color="336699", fill_type="solid"
    )

    # Border styles
    thin_border = Side(border_style="thin", color="000000")
    thick_border = Side(border_style="medium", color="000000")

    header_border = Border(
        left=thin_border, right=thin_border, top=thick_border, bottom=thick_border
    )

    data_border = Border(
        left=thin_border, right=thin_border, top=thin_border, bottom=thin_border
    )

    # Center alignment
    center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Month colors - light blue for alternating columns
    light_blue_fill = PatternFill(
        start_color="E6F2FF", end_color="E6F2FF", fill_type="solid"
    )

    # Apply header styles
    for col_num, cell in enumerate(next(ws.rows), 1):
        cell.font = header_font
        cell.fill = header_fill
        cell.border = header_border
        cell.alignment = center_alignment

    # Adjust column widths
    ws.column_dimensions["A"].width = 15  # Effective date
    ws.column_dimensions["B"].width = 40  # Member Name
    ws.column_dimensions["C"].width = 10  # State

    # Format month columns
    for col_num in range(4, 16):  # Months are columns D through O
        col_letter = get_column_letter(col_num)
        ws.column_dimensions[col_letter].width = 12

        # Alternate light blue fill for month columns
        if col_num % 2 == 0:  # Even columns
            for row_num in range(2, ws.max_row + 1):  # Skip header row
                ws.cell(row=row_num, column=col_num).fill = light_blue_fill

    # Format all data cells
    for row_num in range(2, ws.max_row + 1):  # Skip header row
        for col_num in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_num, column=col_num)
            cell.border = data_border

            # Date and policy name are left-aligned
            if col_num in [1, 2]:
                cell.alignment = left_alignment
            else:
                cell.alignment = center_alignment

            # Apply currency formatting to commission values in month columns
            if col_num >= 4 and isinstance(cell.value, (int, float)):
                ws.cell(row=row_num, column=col_num).number_format = "$#,##0.00"

    # Add conditional formatting for NP values
    np_font = Font(name="Arial", size=11, bold=True, color="FF0000")  # Red font

    for row in ws.iter_rows(min_row=2):  # Skip header
        for cell in row[3:]:  # Month columns only
            if cell.value == "NP":
                cell.font = np_font

    # Format PMPM values with special formatting
    pmpm_font = Font(
        name="Arial", size=11, italic=True, color="333399"
    )  # Dark blue, italic

    for row in ws.iter_rows(min_row=2):  # Skip header
        for cell in row[3:]:  # Month columns only
            if isinstance(cell.value, str) and "PMPM" in cell.value:
                cell.font = pmpm_font


def create_conciliation_master(
    all_records: Dict[str, AllRecord], company: str, selected_month: int
):
    """Create the reconciliation master Excel file using all_records."""
    file_path = f"{company} Conciliation MASTER.xlsx"

    # Define headers
    headers = ["Effective date", "Member Name", "State"]
    months = [
        "ENE",
        "FEB",
        "MAR",
        "ABR",
        "MAY",
        "JUN",
        "JUL",
        "AGO",
        "SEP",
        "OCT",
        "NOV",
        "DIC",
    ]
    headers.extend(months)

    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Conciliation"

    # Add title with company name
    ws.merge_cells("A1:O1")
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = (
        f"{company} COMMISSION RECONCILIATION {datetime.datetime.now().year}"
    )
    title_cell.font = Font(name="Arial", size=14, bold=True)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = PatternFill(
        start_color="336699", end_color="336699", fill_type="solid"
    )
    title_cell.font = Font(name="Arial", size=14, bold=True, color="FFFFFF")

    # Start data from row 3
    ws.append([])  # Empty row after title
    ws.append(headers)

    # Fill the worksheet with data from all_records
    row_index = 4  # Starting row index (after headers)

    for member_name, record in all_records.items():
        # Format the effective date
        effective_date = parse_effective_date(record.effective_date)
        formatted_date = (
            effective_date.strftime("%m/%d/%Y") if effective_date else "Unknown"
        )

        # Initialize month data
        month_data = [""] * 12

        # Fill in the payment data for each month
        for payment in record.payments:
            if type(payment.payment_period) == "str":
                month, year = parse_payment_period(payment.payment_period)
            else:
                month, year = payment.payment_period.month, payment.payment_period.year
            if 0 <= month - 1 < 12:  # Ensure month index is valid
                if year < CURRENT_YEAR:
                    month_data[month - 1] += f"{year} "

                # Format the payment value
                if "pmpm" in payment.action.lower():
                    month_data[month - 1] += f"{int(payment.commission)} PMPM "
                else:
                    month_data[month - 1] += f"{payment.commission} "

        # Mark "NP" for no payment where applicable
        if effective_date:  # and record.in_client_list:
            statement_year = CURRENT_YEAR
            statement_month = STATEMENT_MONTH
            for i in range(12):
                month_date = datetime.datetime(statement_year, i + 1, 1)
                # Only mark as NP if the month has already started (not in the future)
                # and the effective date is on or before that month and no payment recorded
                if (
                    effective_date <= month_date
                    and (i + 1) <= statement_month
                    and month_data[i] == ""
                ):
                    month_data[i] = "NP"
                elif (
                    effective_date > month_date
                    and month_data[i] == ""
                    and (i + 1) <= statement_month
                ):
                    month_data[i] = "NE"

        # Create the row data
        row = [formatted_date, member_name, record.state]
        row.extend(month_data)
        ws.append(row)

        # Apply bold formatting to members not in client list
        if not record.in_client_list:
            ws.cell(row=row_index, column=1).font = Font(bold=True)
            ws.cell(row=row_index, column=2).font = Font(bold=True)
            ws.cell(row=row_index, column=3).font = Font(bold=True)

        row_index += 1

    # Apply beautiful formatting
    apply_excel_formatting(ws)

    # Add filters to header row
    ws.auto_filter.ref = (
        f"A3:O{ws.max_row-2}"  # Exclude title and summary row from filter
    )

    # Save workbook
    try:
        wb.save(file_path)
        print(f"File '{file_path}' created successfully!")
        return file_path
    except Exception as e:
        print(f"Error saving file '{file_path}': {e}")


def compare_names(statements_paths, client_list_path, company="BCBS"):
    """Match statement records with client list records using fuzzy matching."""
    # Read data
    client_records = read_client_list(client_list_path)
    statement_records = []
    for statement in statements_paths:
        statement_records.extend(read_bcbs(statement))

    # Filter for UHC clients
    client_list_company = [c for c in client_records if c.company == company]
    matched_rows = []
    unmatched_clients = []
    not_in_client_list = statement_records.copy()

    # Process each client and look for matches
    for client in client_list_company:
        # For each active member in the client record
        client_matched = False
        n_member_matched = 0
        for member_name in client.members_list:
            member_matched = False

            # Try to find a match in UHC records
            for statement_record in statement_records:
                # Use fuzzy matching to compare names
                ratio = fuzz.token_set_ratio(
                    member_name,
                    statement_record.member_name,
                    processor=utils.default_process,
                )

                # If good match found
                if ratio > 90:
                    matched_rows.append([statement_record, client, member_name])
                    member_matched = True
                    client_matched = True
                    n_member_matched += 1
                    # Remove from not_in_client_list
                    if statement_record in not_in_client_list:
                        not_in_client_list.remove(statement_record)

            # If this member wasn't matched, add to unmatched list
            if not member_matched:
                unmatched_clients.append([client, member_name])

    return matched_rows, unmatched_clients, not_in_client_list


def merge_records(matched_rows, unmatched_clients, not_in_client_list, company):
    """Merge the records lists into 1 as {"client_name: AllRecords"}"""
    all_records = {}
    for row in matched_rows:
        statement_record = row[0]
        client_record = row[1]
        member_name = row[2]
        effective_date = statement_record.effective_date
        state = client_record.state
        payment_period = statement_record.payment_period
        commission = statement_record.commission
        action = statement_record.action
        in_client_list = True
        payment = Payment(action, payment_period, commission)
        if member_name in all_records:
            all_records[member_name].payments.append(payment)
        else:
            all_records[member_name] = AllRecord(
                member_name, effective_date, state, company, in_client_list, [payment]
            )

    for row in unmatched_clients:  # mean no payments
        client_record = row[0]
        member_name = row[1]
        effective_date = client_record.effective_date
        state = client_record.state
        in_client_list = True
        if member_name not in all_records:  # Only add if not already in records
            all_records[member_name] = AllRecord(
                member_name, effective_date, state, company, in_client_list, []
            )

    for statement_record in not_in_client_list:
        member_name = statement_record.member_name
        effective_date = statement_record.effective_date
        state = statement_record.state
        payment_period = statement_record.payment_period
        commission = statement_record.commission
        action = statement_record.action
        in_client_list = False
        payment = Payment(action, payment_period, commission)
        if member_name in all_records:
            all_records[member_name].payments.append(payment)
        else:
            all_records[member_name] = AllRecord(
                member_name, effective_date, state, company, in_client_list, [payment]
            )

    return all_records


def main(
    statements_paths=statements_list,
    client_list_path=client_list_path,
    selected_month=STATEMENT_MONTH,
    company="BCBS",
):
    """Main function to run the reconciliation process."""

    print("Starting insurance reconciliation process...")

    # Compare names and get matches
    matched, unmatched, not_in_client_list = compare_names(
        statements_paths, client_list_path, company
    )
    # Print results
    print(f"Matched records: {len(matched)}")
    print(f"Unmatched clients: {len(unmatched)}")
    print(f"Records not in client list: {len(not_in_client_list)}")

    # Merge all records into a unified structure
    all_records = merge_records(matched, unmatched, not_in_client_list, company)
    print(f"Total unique members: {len(all_records)}")

    # Create reconciliation master file using the merged all_records
    output_path = create_conciliation_master(all_records, company, selected_month)

    print(f"Reconciliation process completed: {output_path}")
    return output_path


if __name__ == "__main__":
    main()
