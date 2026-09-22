from __future__ import annotations

import sys
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Iterable, Optional

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill # type: ignore
from openpyxl.utils import get_column_letter # type: ignore



# Output column names for attendance sheet and audit log.
OUTPUT_COLUMNS = ["Employee Name","Employee id","Email","Shift Timings","Year","Month","Week","Checked In Date","Checked Out Date","# of hours in Office","Date","Daily Total Hours"]
AUDIT_COLUMNS = ["Employee id", "Employee Name", "Issue", "Details", "Timestamp"]

# Required columns in shift roster
@dataclass(frozen=True, slots=True)
class ShiftRecord:
    emp_id: str
    name: str
    email: str
    shift_timing: str

#column names in monthly movement report
@dataclass(slots=True)
class VisitRecord:
    emp_code: str
    name: str
    check_in: datetime
    check_out: Optional[datetime] = None
    email: str = ""
    shift_timing: str = ""

    @property
    def year(self) -> int:
        return self.check_in.year

    @property
    def month_name(self) -> str:
        return self.check_in.strftime("%B")

    @property
    def week_of_month(self) -> int:
        return ((self.check_in.day - 1) // 7) + 1

    @property
    def hours_in_office(self) -> Optional[float]:
        if self.check_out is None:
            return None
        delta = self.check_out - self.check_in
        return round(delta.total_seconds() / 3600, 2)

    def as_row(self) -> list:
        return [self.name,self.emp_code,self.email,self.shift_timing,self.year,
        self.month_name,self.week_of_month,self.check_in,self.check_out,self.hours_in_office,
        self.check_in.date().isoformat(), None]

@dataclass(slots=True)
class AuditEntry:
    emp_code: str
    name: str
    issue: str
    detail: str
    timestamp: Optional[datetime] = None
    def as_row(self) -> list:
        return [self.emp_code, self.name, self.issue, self.detail, self.timestamp]


# Parsing and normalization helpers
def normalize_id(value) -> str:
    # while reading excel numbers are converted to float, so we need to remove the decimal part if it is .0
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    digits = "".join(ch for ch in text if ch.isdigit())
    return digits or text


# parsing date in the excel to python datetime
def parse_date(value) -> Optional[date]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.date()
    if isinstance(value, date):
        return value
    parsed = pd.to_datetime(str(value).strip(), dayfirst=True, errors="coerce")
    return None if pd.isna(parsed) else parsed.date()


#parsing the check in and out time to python format
def parse_time(value) -> Optional[time]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, time):
        return value
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.time()
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return None
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(text, fmt).time()
        except ValueError:
            continue
    parsed = pd.to_datetime(text, errors="coerce")
    return None if pd.isna(parsed) else parsed.time()


# strips every text to removes spaces.
def clean_text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text



# core engine
class AttendanceEngine:
    HEADER_PROBE = "empcode" #the column name we look for to detect the header row in the movement report
    MAX_HEADER_SCAN_ROWS = 15 #how many rows to scan for the header row in the movement report
    MAX_SHIFT_HOURS = 18 #taken for the maximum number of hours allowed in a single check-in/check-out session

    def __init__(
        self,
        movement_path: Path,
        shift_path: Optional[Path] = None,
        previous_movement_path: Optional[Path] = None,
    ):
        self.movement_path = Path(movement_path)
        self.shift_path = Path(shift_path) if shift_path else None
        self.previous_movement_path = (Path(previous_movement_path) if previous_movement_path else None)
        self.shift_lookup: dict[str, ShiftRecord] = {}
        self.audit: list[AuditEntry] = []
        self._reported_missing_roster: set[str] = set()

    def run(self) -> tuple[list[VisitRecord], list[AuditEntry]]:
        if self.shift_path:
            self.shift_lookup = self._load_shift_roster(self.shift_path)
        header, data = self._load_movement_table(self.movement_path)
        events = self._extract_events(header, data)
        current_dates = [timestamp for _, (_, ins, outs) in events.items() for timestamp in (*ins, *outs)]

        if self.previous_movement_path:
            previous_header, previous_data = self._load_movement_table(self.previous_movement_path)
            previous_last_date = self._last_transaction_date(previous_header, previous_data)
            previous_events = self._extract_events(previous_header, previous_data, {previous_last_date}) if previous_last_date else {}
            events = self._merge_event_pools(previous_events, events)

        visits: list[VisitRecord] = []
        for emp_code, (name, ins, outs) in events.items():
            visits.extend(self._reconcile_employee(emp_code, name, ins, outs))

        if self.previous_movement_path and current_dates:
            current_start = min(current_dates).date().replace(day=1)
            next_month = (current_start.replace(year=current_start.year + 1, month=1)if current_start.month == 12 else current_start.replace(month=current_start.month + 1))
            visits = [visit for visit in visits if current_start <= visit.check_in.date() < next_month
                or (visit.check_out is not None and current_start <= visit.check_out.date() < next_month)]

        self._enrich_with_roster(visits)
        visits.sort(key=lambda v: (v.name.lower(), v.check_in))
        return visits, self.audit

    @staticmethod
    def _merge_event_pools(first: dict[str, tuple[str, list[datetime], list[datetime]]],
        second: dict[str, tuple[str, list[datetime], list[datetime]]],) -> dict[str, tuple[str, list[datetime], list[datetime]]]:
        merged: dict[str, tuple[str, set[datetime], set[datetime]]] = {}
        for source in (first, second):
            for emp_code, (name, ins, outs) in source.items():
                current_name, current_ins, current_outs = merged.setdefault(emp_code, (name, set(), set()))
                if name and not current_name:
                    current_name = name
                current_ins.update(ins)
                current_outs.update(outs)
                merged[emp_code] = (current_name, current_ins, current_outs)
        return {emp_code: (name or emp_code, sorted(ins), sorted(outs)) for emp_code, (name, ins, outs) in merged.items()}

    # since the monthly movement report can have multiple rows with heading and actual rows 
    # are at different levels we are extracing exact row numner!!
    @classmethod
    def _detect_header_row(cls, raw: pd.DataFrame) -> int:
        scan = min(cls.MAX_HEADER_SCAN_ROWS, len(raw))
        for i in range(scan):
            row_values = {clean_text(v).lower() for v in raw.iloc[i].tolist()}
            if cls.HEADER_PROBE in row_values:
                return i
        raise ValueError(
            f"Could not find a header row containing '{cls.HEADER_PROBE}' "
            f"in the first {scan} rows of {raw.attrs.get('source', 'the file')}."
        )

    # calulating the header and data rows from the movement report, based on the detected header row
    def _load_movement_table(self, path: Path) -> tuple[list[str], pd.DataFrame]:
        if path.suffix.lower() == ".csv":
            raw = pd.read_csv(path, header=None, dtype=str, keep_default_na=True)
        else:
            raw = pd.read_excel(path, header=None, dtype=str, sheet_name=0)
        raw.attrs["source"] = path.name
        header_row = self._detect_header_row(raw)
        #header columns names
        header = [clean_text(v) for v in raw.iloc[header_row].tolist()]
        #actual data rows after the header row
        data = raw.iloc[header_row + 1 :].reset_index(drop=True)
        return header, data

    @staticmethod
    def _last_transaction_date(header: list[str], data: pd.DataFrame) -> Optional[date]:
        lower_header = [item.lower() for item in header]
        date_index = next(
            (lower_header.index(name) for name in ("transdate", "date") if name in lower_header),
            None,
        )
        if date_index is None:
            return None
        dates = [
            parsed
            for value in data.iloc[:, date_index]
            if (parsed := parse_date(value)) is not None
        ]
        return max(dates) if dates else None

    # generate a dictionary of employee id -> (name, list of check-ins, list of check-outs) from the movement report
    def _extract_events(
        self,
        header: list[str],
        data: pd.DataFrame,
        allowed_dates: Optional[set[date]] = None,
    ) -> dict[str, tuple[str, list[datetime], list[datetime]]]:
        # convert all header names to lowercase for case-insensitive matching
        lower_header = [h.lower() for h in header]

        def find_col(*names: str) -> int:
            for name in names:
                if name in lower_header:
                    return lower_header.index(name)
            raise ValueError(f"Movement report is missing a required column: {names}")

        col_emp = find_col("empcode", "employee id", "employee_id")
        col_name = find_col("employeename", "employee name")
        col_date = find_col("transdate", "date")
        #check in and check out columns are variable in number and can be anywhere in the report so we are collecting all the columns which have "in" or "out" in their header name
        swipe_cols = [(i, lower_header[i]) for i in range(len(header)) if lower_header[i] in ("in", "out")]
        # employee id -> (name, set of check-ins, set of check-outs)
        pools: dict[str, tuple[str, set[datetime], set[datetime]]] = {}

        # iterate over each row in the data and extract employee code, name, transaction date, and check-in/check-out timestamps
        for _, row in data.iterrows():
            raw_emp_id = row.iloc[col_emp] if col_emp < len(row) else None
            if raw_emp_id is None or clean_text(raw_emp_id) == "":
                continue
            emp_code = normalize_id(raw_emp_id)
            name = clean_text(row.iloc[col_name]) if col_name < len(row) else ""
            trans_date = parse_date(row.iloc[col_date]) if col_date < len(row) else None
            if trans_date is None:
                self.audit.append(AuditEntry(emp_code, name, "Unparsable row","TransDate could not be parsed; row skipped."))
                continue
            if allowed_dates is not None and trans_date not in allowed_dates:
                continue

            # for each employee, we maintain a pool of their check-in and check-out timestamps. 
            # If the employee already exists in the pool, we update their name if it's currently empty. 
            # We then iterate over the swipe columns to extract and parse the timestamps, 
            # adding them to the appropriate set (check-ins or check-outs) based on the column type.
            name_slot, ins, outs = pools.setdefault(emp_code, (name, set(), set()))
            if name and not name_slot:
                pools[emp_code] = (name, ins, outs)

            #parsing the check-in and check-out timestamps from the swipe columns and adding them to the respective sets
            for col_idx, kind in swipe_cols:
                if col_idx >= len(row):
                    continue
                t = parse_time(row.iloc[col_idx])
                if t is None:
                    continue
                stamp = datetime.combine(trans_date, t)
                (ins if kind == "in" else outs).add(stamp)

        return { emp_code: (name or emp_code, sorted(ins), sorted(outs)) for emp_code, (name, ins, outs) in pools.items()}

    #pairs each check-in with the earliest unused checkout that is later than the check-in and earlier than the next check-in, w
    # hile also enforcing the rules regarding allowed dates and maximum shift hours. If no valid checkout is found, it logs an audit entry and leaves the checkout blank.
    def _reconcile_employee(self, emp_code: str, name: str, ins: list[datetime], outs: list[datetime]) -> list[VisitRecord]:
        remaining_outs = list(outs)
        visits: list[VisitRecord] = []

        for idx, check_in in enumerate(ins):
            #calulating the next check-in timestamp if it exists, otherwise setting it to None
            next_check_in = ins[idx + 1] if idx + 1 < len(ins) else None
            checkout = self._find_checkout(check_in, next_check_in, remaining_outs, self.MAX_SHIFT_HOURS)
            if checkout is not None:
                remaining_outs.remove(checkout)
            else:
                self.audit.append(
                    AuditEntry(emp_code, name, "Orphan checkin",
                        f"Check-in at {check_in:%Y-%m-%d %H:%M:%S} has no valid "
                        f"checkout before the next check-in, so left blank.",check_in,))
            visits.append(VisitRecord(emp_code=emp_code, name=name,check_in=check_in, check_out=checkout))


        # the check out which are not paired with any checkin are left so auditing all of them.
        for orphan in remaining_outs:
            self.audit.append(
                AuditEntry(emp_code, name, "Orphan checkout",
                    f"Checkout at {orphan:%Y-%m-%d %H:%M:%S} has no preceding "
                    f"check-in in this report, so dropped from the attendance sheet.",orphan,))
        return visits
    
    #finds the earliest valid checkout timestamp for a given check-in, considering the next check-in and the remaining candidate checkouts. 
    # It checks if the checkout is after the check-in, before the next check-in (if it exists), falls within the allowed dates (the same day or the next day), 
    # and does not exceed the maximum shift hours. If a valid checkout is found, it returns that timestamp; otherwise, it returns None.
    @staticmethod
    def _find_checkout(check_in: datetime,next_check_in: Optional[datetime],candidates: Iterable[datetime],max_shift_hours: float,) -> Optional[datetime]:
        allowed_dates = {check_in.date(), check_in.date() + timedelta(days=1)}
        max_gap = timedelta(hours=max_shift_hours)
        #candidates -> remaining check-out timestamps for the employee, sorted in ascending order
        for out_dt in candidates:
            if out_dt <= check_in:
                continue
            if next_check_in is not None and out_dt >= next_check_in:
                break  # sorted ascending -> nothing further can qualify either
            if out_dt.date() not in allowed_dates:
                continue
            if out_dt - check_in > max_gap:
                continue  # plausible calendar-wise, but too long to be one visit
            return out_dt
        return None



    # reads the shift roster and returns a lookup table of employee id -> ShiftRecord
    @staticmethod
    def _load_shift_roster(path: Path) -> dict[str, ShiftRecord]:
        df = pd.read_excel(path, sheet_name=0)
        lookup: dict[str, ShiftRecord] = {}
        for idx, row in df.iterrows():
            emp_id_raw = row.get("Emplid")
            if pd.isna(emp_id_raw):
                continue
            name = clean_text(row.get("Legal Name"))
            #name is blank skip the record, as we cannot match it to any employee
            if not name:
                continue
            emp_id = normalize_id(emp_id_raw)
            current_shift = clean_text(row.get("Current Shift Timing"))
            base_shift = clean_text(row.get("Shit timing", row.get("Shift Timing", "")))
            lookup[emp_id] = ShiftRecord(emp_id=emp_id,name=name,email=clean_text(row.get("Email")),shift_timing=current_shift or base_shift,)
        return lookup

    # enriches the VisitRecord objects with additional information from the shift roster, if available.
    def _enrich_with_roster(self, visits: list[VisitRecord]) -> None:
        for visit in visits:
            record = self.shift_lookup.get(visit.emp_code)
            if record:
                visit.name = record.name or visit.name
                visit.email = record.email
                visit.shift_timing = record.shift_timing
            elif self.shift_path and visit.emp_code not in self._reported_missing_roster:
                self._reported_missing_roster.add(visit.emp_code)
                self.audit.append(AuditEntry(visit.emp_code, visit.name, "Not found in shift roster",
                    "Employee id from the movement report has no match in "
                    "the shift roster; name/email/shift left as-is.",))


# writes the attendance and audit log to an Excel file, applying formatting to the headers and date columns. 
# It uses pandas to create DataFrames from the VisitRecord and AuditEntry objects, then writes them to separate 
# sheets in the output Excel file. The _style_sheet method is responsible for styling the headers, freezing the top row, 
# and adjusting column widths based on content length.
class ReportWriter:
    DATE_FORMAT = "yyyy-mm-dd hh:mm:ss"
    HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
    HEADER_FONT = Font(bold=True, color="FFFFFF")

    @staticmethod
    def _shift_end_time(shift_timing: str) -> Optional[time]:
        parts = [part.strip() for part in shift_timing.split("-")]
        if len(parts) != 2:
            return None
        try:
            return datetime.strptime(parts[1], "%I:%M %p").time()
        except ValueError:
            return None

    @classmethod
    def _work_date(cls, check_in: datetime, shift_timing: str) -> date:
        parts = [part.strip() for part in shift_timing.split("-")]
        if len(parts) != 2:
            return check_in.date()
        try:
            shift_start = datetime.strptime(parts[0], "%I:%M %p").time()
            shift_end = cls._shift_end_time(shift_timing)
        except ValueError:
            return check_in.date()

        if shift_end is None or shift_end >= shift_start:
            return check_in.date()
        if check_in.time() <= shift_end:
            return check_in.date() - timedelta(days=1)
        return check_in.date()

    @classmethod
    def write(cls, visits: list[VisitRecord], audit: list[AuditEntry], output_path: Path) -> None:
        attendance_df = pd.DataFrame([v.as_row() for v in visits], columns=OUTPUT_COLUMNS)
        audit_df = pd.DataFrame([a.as_row() for a in audit], columns=AUDIT_COLUMNS)

        if not attendance_df.empty:
            attendance_df["Date"] = attendance_df.apply(
                lambda row: cls._work_date(row["Checked In Date"], row["Shift Timings"]).isoformat(),
                axis=1,
            )
            attendance_df["_employee_day_key"] = attendance_df.apply(
                lambda row: (str(row["Employee id"]), row["Date"]), axis=1
            )
            attendance_df["Daily Total Hours"] = None
            for key, group in attendance_df.groupby("_employee_day_key", sort=False, dropna=False):
                total_hours = group["# of hours in Office"].sum()
                first_row = group.index[0]
                attendance_df.at[first_row, "Daily Total Hours"] = total_hours
            attendance_df = attendance_df.drop(columns=["_employee_day_key"])

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            attendance_df.to_excel(writer, sheet_name="Attendance", index=False)
            audit_df.to_excel(writer, sheet_name="Audit Log", index=False)

            attendance_sheet = writer.book["Attendance"]
            cls._apply_daily_total_merges(attendance_sheet)
            cls._style_sheet(attendance_sheet, date_cols={8, 9})
            cls._style_sheet(writer.book["Audit Log"], date_cols={5})

    @classmethod
    def _apply_daily_total_merges(cls, sheet) -> None:
        emp_col = OUTPUT_COLUMNS.index("Employee id") + 1
        date_col = OUTPUT_COLUMNS.index("Date") + 1
        daily_total_col = OUTPUT_COLUMNS.index("Daily Total Hours") + 1

        last_row = sheet.max_row
        if last_row <= 1:
            return

        current_group = None
        current_start = None

        for row_idx in range(2, last_row + 1):
            emp_value = sheet.cell(row=row_idx, column=emp_col).value
            date_value = sheet.cell(row=row_idx, column=date_col).value
            group_key = (str(emp_value), str(date_value)) if emp_value is not None and date_value is not None else None

            if current_group is None:
                current_group = group_key
                current_start = row_idx
                continue

            if group_key == current_group:
                continue

            if current_start is not None:
                current_end = row_idx - 1
                if current_end >= current_start:
                    for col_idx in [date_col, daily_total_col]:
                        merged_range = f"{get_column_letter(col_idx)}{current_start}:{get_column_letter(col_idx)}{current_end}"
                        if sheet.merged_cells.ranges:
                            if merged_range not in [rng.coord for rng in sheet.merged_cells.ranges]:
                                sheet.merge_cells(merged_range)
                        else:
                            sheet.merge_cells(merged_range)
                        first_cell = sheet.cell(row=current_start, column=col_idx)
                        first_cell.alignment = Alignment(horizontal="center", vertical="center")

            current_group = group_key
            current_start = row_idx

        if current_start is not None:
            current_end = last_row
            if current_end >= current_start:
                for col_idx in [date_col, daily_total_col]:
                    merged_range = f"{get_column_letter(col_idx)}{current_start}:{get_column_letter(col_idx)}{current_end}"
                    if sheet.merged_cells.ranges:
                        if merged_range not in [rng.coord for rng in sheet.merged_cells.ranges]:
                            sheet.merge_cells(merged_range)
                    else:
                        sheet.merge_cells(merged_range)
                    first_cell = sheet.cell(row=current_start, column=col_idx)
                    first_cell.alignment = Alignment(horizontal="center", vertical="center")

    @classmethod
    def _style_sheet(cls, sheet, date_cols: set[int]) -> None:
        for cell in sheet[1]:
            cell.fill = cls.HEADER_FILL
            cell.font = cls.HEADER_FONT
            cell.alignment = Alignment(horizontal="center")
        sheet.freeze_panes = "A2"

        widths: dict[int, int] = {}
        for row in sheet.iter_rows(min_row=1):
            for cell in row:
                if cell.column in date_cols and isinstance(cell.value, datetime):
                    cell.number_format = cls.DATE_FORMAT
                text_len = len(str(cell.value)) if cell.value is not None else 0
                widths[cell.column] = max(widths.get(cell.column, 10), min(text_len + 2, 40))
        for col_idx, width in widths.items():
            sheet.column_dimensions[get_column_letter(col_idx)].width = width



# file input
def _prompt_path(message: str, required: bool) -> Optional[Path]:
    while True:
        raw = input(message).strip().strip('"').strip("'")
        if not raw:
            if required:
                print("  This file is required - please enter a path.")
                continue
            return None
        candidate = Path(raw).expanduser()
        if not candidate.exists():
            print(f"  Could not find '{candidate}'. Try again.")
            continue
        return candidate


def _default_output_name(visits: list[VisitRecord]) -> str:
    if visits:
        sample = visits[0].check_in
        return f"Attendance_Report_{sample:%B_%Y}.xlsx"
    return f"Attendance_Report_{datetime.now():%Y%m%d_%H%M%S}.xlsx"


def main() -> None:
    print("=" * 70)
    print("Attendance Reconciliation Engine")
    print("=" * 70)

    # check-in / check-out report and shift roster
    movement_path = _prompt_path("\nPath to this month's check-in / check-out report (.xlsx or .csv): ",required=True,)
    previous_movement_path = _prompt_path(
        "Path to last month's check-in / check-out report (optional, press Enter to skip): ",
        required=False,
    )
    # shift roster is optional, so we allow the user to skip it by pressing Enter
    shift_path = _prompt_path("Path to the shift roster file (optional, press Enter to skip): ",required=False,)

    # run the engine and get the visits and audit log
    engine = AttendanceEngine(movement_path, shift_path, previous_movement_path)
    visits, audit = engine.run()

    # prompt for output file name, defaulting to a name based on the first visit's check-in date or the current timestamp if no visits are found
    default_name = _default_output_name(visits)
    out_raw = input(f"Output file name (optional, press Enter for '{default_name}'): ").strip()
    output_name = out_raw or default_name
    if not output_name.lower().endswith(".xlsx"):
        output_name += ".xlsx"
    output_path = movement_path.parent / output_name

    # write the report to the specified output path
    ReportWriter.write(visits, audit, output_path)

    complete = sum(1 for v in visits if v.check_out is not None)
    incomplete = len(visits) - complete
    employees = len({v.emp_code for v in visits})

    print("\nDone.")
    print(f"  Employees processed : {employees}")
    print(f"  Visits (check-ins)  : {len(visits)}")
    print(f"  with a checkout     : {complete}")
    print(f"  without a checkout  : {incomplete}")
    print(f"  Audit log entries   : {len(audit)}")
    print(f"  Report written to   : {output_path}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nCancelled.")
