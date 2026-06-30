import argparse
from pathlib import Path

from fastmcp import FastMCP
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

mcp = FastMCP(name="Excel CRUD MCP")

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _sanitize(value) -> str:
    """Replace newlines and tabs with literal \\n and \\t."""
    if value is None:
        return ""
    return str(value).replace("\n", "\\n").replace("\t", "\\t")


def _open_workbook(file_path: str, read_only: bool = False):
    """Open an existing workbook. Raises ValueError if file not found."""
    path = Path(file_path)
    if not path.exists():
        raise ValueError(f"File not found: {file_path}")
    return load_workbook(path, read_only=read_only)


def _get_sheet(wb, sheet_name: str | None = None):
    """Get a sheet by name or first sheet. Raises ValueError if not found."""
    if sheet_name:
        if sheet_name not in wb.sheetnames:
            raise ValueError(
                f"Sheet '{sheet_name}' not found. Available sheets: {wb.sheetnames}"
            )
        return wb[sheet_name]
    # Default to first sheet
    return wb[wb.sheetnames[0]]


def _sheet_to_tsv(ws) -> str:
    """Convert an openpyxl worksheet to a TSV string."""
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return ""

    lines = []
    for row in rows:
        lines.append("\t".join(_sanitize(cell) for cell in row))
    return "\n".join(lines)


def _get_column_idx(ws, column_name: str) -> int:
    """Return the 1-based column index for a column header name."""
    header_row = list(next(ws.iter_rows(min_row=1, max_row=1, values_only=True)))
    for idx, cell in enumerate(header_row, start=1):
        if cell is not None and str(cell).strip() == column_name.strip():
            return idx
    raise ValueError(f"Column '{column_name}' not found in header row")


# =============================================
# CREATE tools
# =============================================


@mcp.tool
def add_row(file_path: str, row_data: str, sheet_name: str | None = None) -> str:
    """Add a new row at the end of the sheet.

    Args:
        file_path: Path to the Excel .xlsx file.
        row_data: Tab-separated values for the row (e.g. "val1\\tval2\\tval3").
        sheet_name: Sheet name (defaults to first sheet).
    """
    try:
        wb = _open_workbook(file_path)
        ws = _get_sheet(wb, sheet_name)
        values = row_data.split("\t")
        ws.append(values)
        wb.save(file_path)
        wb.close()
        return f"Row added successfully to sheet '{ws.title}'"
    except Exception as e:
        return f"Error: {e}"


@mcp.tool
def add_cell_data(
    file_path: str,
    row_number: int,
    column_name: str,
    value: str,
    sheet_name: str | None = None,
) -> str:
    """Write a value into a specific cell identified by column name and row number.

    Args:
        file_path: Path to the Excel .xlsx file.
        row_number: Row number (1 = first data row, header is row 1).
        column_name: Column header name.
        value: Value to write.
        sheet_name: Sheet name (defaults to first sheet).
    """
    try:
        wb = _open_workbook(file_path)
        ws = _get_sheet(wb, sheet_name)
        col_idx = _get_column_idx(ws, column_name)
        row_idx = row_number + 1  # +1 for header
        ws.cell(row=row_idx, column=col_idx, value=value)
        wb.save(file_path)
        wb.close()
        return f"Cell ({row_number}, '{column_name}') set to '{value}' in sheet '{ws.title}'"
    except Exception as e:
        return f"Error: {e}"


@mcp.tool
def add_column(
    file_path: str,
    column_name: str,
    default_value: str = "",
    sheet_name: str | None = None,
) -> str:
    """Add a new column with a header name and optional default value for existing rows.

    Args:
        file_path: Path to the Excel .xlsx file.
        column_name: Header name for the new column.
        default_value: Default value for existing rows (default: empty string).
        sheet_name: Sheet name (defaults to first sheet).
    """
    try:
        wb = _open_workbook(file_path)
        ws = _get_sheet(wb, sheet_name)
        new_col = (ws.max_column or 0) + 1
        # Write header
        ws.cell(row=1, column=new_col, value=column_name)
        # Fill default for existing data rows
        for r in range(2, (ws.max_row or 1) + 1):
            ws.cell(row=r, column=new_col, value=default_value)
        wb.save(file_path)
        wb.close()
        return f"Column '{column_name}' added to sheet '{ws.title}'"
    except Exception as e:
        return f"Error: {e}"


@mcp.tool
def add_full_data(
    file_path: str,
    data: str,
    sheet_name: str | None = None,
) -> str:
    """Replace all data in a sheet with TSV-formatted content.

    The first line of `data` is treated as column headers.
    Subsequent lines are data rows.

    Args:
        file_path: Path to the Excel .xlsx file.
        data: TSV-formatted data (header line + data lines, tab-separated).
        sheet_name: Sheet name (defaults to first sheet).
    """
    try:
        wb = _open_workbook(file_path)
        ws = _get_sheet(wb, sheet_name)
        # Clear existing content
        if ws.max_row:
            ws.delete_rows(1, ws.max_row)
        lines = data.strip().split("\n")
        for i, line in enumerate(lines):
            values = line.split("\t")
            for j, val in enumerate(values, start=1):
                ws.cell(row=i + 1, column=j, value=val)
        wb.save(file_path)
        wb.close()
        return f"Full data written to sheet '{ws.title}'"
    except Exception as e:
        return f"Error: {e}"


@mcp.tool
def create_sheet(file_path: str, sheet_name: str) -> str:
    """Create a new sheet in the workbook.

    Args:
        file_path: Path to the Excel .xlsx file.
        sheet_name: Name for the new sheet.
    """
    try:
        wb = _open_workbook(file_path)
        if sheet_name in wb.sheetnames:
            wb.close()
            return f"Error: Sheet '{sheet_name}' already exists"
        wb.create_sheet(title=sheet_name)
        wb.save(file_path)
        wb.close()
        return f"Sheet '{sheet_name}' created successfully"
    except Exception as e:
        return f"Error: {e}"


# =============================================
# READ tools
# =============================================


@mcp.tool
def read_excel(file_path: str, sheet_name: str | None = None) -> str:
    """Read an Excel sheet and return its content as TSV.

    Args:
        file_path: Path to the Excel .xlsx file.
        sheet_name: Sheet name (defaults to first sheet).
    """
    try:
        wb = _open_workbook(file_path, read_only=True)
        ws = _get_sheet(wb, sheet_name)
        result = _sheet_to_tsv(ws)
        wb.close()
        return result
    except Exception as e:
        return f"Error: {e}"


@mcp.tool
def read_rows_by_column(
    file_path: str,
    column_name: str,
    column_value: str,
    sheet_name: str | None = None,
) -> str:
    """Read rows where a specific column matches a value.

    Args:
        file_path: Path to the Excel .xlsx file.
        column_name: Column header name to filter on.
        column_value: Value to match in the column.
        sheet_name: Sheet name (defaults to first sheet).
    """
    try:
        wb = _open_workbook(file_path, read_only=True)
        ws = _get_sheet(wb, sheet_name)
        col_idx = _get_column_idx(ws, column_name)

        headers = [cell for cell in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
        matching_rows = []

        for row in ws.iter_rows(min_row=2, values_only=True):
            cell_value = row[col_idx - 1]
            if cell_value is not None and str(cell_value).strip() == column_value.strip():
                matching_rows.append(row)

        wb.close()

        if not matching_rows:
            return f"No rows found where '{column_name}' = '{column_value}'"

        result_lines = ["\t".join(_sanitize(h) for h in headers)]
        for row in matching_rows:
            result_lines.append("\t".join(_sanitize(c) for c in row))

        return "\n".join(result_lines)
    except Exception as e:
        return f"Error: {e}"


# =============================================
# UPDATE tools
# =============================================


@mcp.tool
def update_row(
    file_path: str,
    row_number: int,
    row_data: str,
    sheet_name: str | None = None,
) -> str:
    """Update an entire row with new values.

    Args:
        file_path: Path to the Excel .xlsx file.
        row_number: Row number (1 = first data row, header is row 1).
        row_data: Tab-separated new values for the row.
        sheet_name: Sheet name (defaults to first sheet).
    """
    try:
        wb = _open_workbook(file_path)
        ws = _get_sheet(wb, sheet_name)
        values = row_data.split("\t")
        row_idx = row_number + 1  # +1 for header
        for j, val in enumerate(values, start=1):
            ws.cell(row=row_idx, column=j, value=val)
        wb.save(file_path)
        wb.close()
        return f"Row {row_number} updated in sheet '{ws.title}'"
    except Exception as e:
        return f"Error: {e}"


@mcp.tool
def update_cell(
    file_path: str,
    row_number: int,
    column_name: str,
    value: str,
    sheet_name: str | None = None,
) -> str:
    """Update a single cell by column name and row number.

    Args:
        file_path: Path to the Excel .xlsx file.
        row_number: Row number (1 = first data row).
        column_name: Column header name.
        value: New value for the cell.
        sheet_name: Sheet name (defaults to first sheet).
    """
    try:
        wb = _open_workbook(file_path)
        ws = _get_sheet(wb, sheet_name)
        col_idx = _get_column_idx(ws, column_name)
        row_idx = row_number + 1
        ws.cell(row=row_idx, column=col_idx, value=value)
        wb.save(file_path)
        wb.close()
        return (
            f"Cell ({row_number}, '{column_name}') updated to '{value}'"
            f" in sheet '{ws.title}'"
        )
    except Exception as e:
        return f"Error: {e}"


@mcp.tool
def replace_full_data(
    file_path: str,
    new_data: str,
    sheet_name: str | None = None,
) -> str:
    """Replace all data in a sheet with new TSV-formatted content.

    This is an alias for add_full_data with clearer semantics for updates.

    Args:
        file_path: Path to the Excel .xlsx file.
        new_data: TSV-formatted data (first line = headers).
        sheet_name: Sheet name (defaults to first sheet).
    """
    return add_full_data(file_path, new_data, sheet_name)


# =============================================
# DELETE tools
# =============================================


@mcp.tool
def delete_sheet(file_path: str, sheet_name: str) -> str:
    """Delete a sheet from the workbook.

    Args:
        file_path: Path to the Excel .xlsx file.
        sheet_name: Name of the sheet to delete.
    """
    try:
        wb = _open_workbook(file_path)
        if sheet_name not in wb.sheetnames:
            wb.close()
            return f"Error: Sheet '{sheet_name}' not found"
        if len(wb.sheetnames) == 1:
            wb.close()
            return "Error: Cannot delete the only sheet in the workbook"
        del wb[sheet_name]
        wb.save(file_path)
        wb.close()
        return f"Sheet '{sheet_name}' deleted"
    except Exception as e:
        return f"Error: {e}"


@mcp.tool
def delete_row(
    file_path: str,
    row_number: int,
    sheet_name: str | None = None,
) -> str:
    """Delete a row from the sheet.

    Args:
        file_path: Path to the Excel .xlsx file.
        row_number: Row number to delete (1 = first data row, header is row 1).
        sheet_name: Sheet name (defaults to first sheet).
    """
    try:
        wb = _open_workbook(file_path)
        ws = _get_sheet(wb, sheet_name)
        row_idx = row_number + 1  # +1 for header
        if ws.max_row and row_idx > ws.max_row:
            wb.close()
            return f"Error: Row {row_number} does not exist in sheet '{ws.title}'"
        ws.delete_rows(row_idx)
        wb.save(file_path)
        wb.close()
        return f"Row {row_number} deleted from sheet '{ws.title}'"
    except Exception as e:
        return f"Error: {e}"


@mcp.tool
def delete_column(
    file_path: str,
    column_name: str,
    sheet_name: str | None = None,
) -> str:
    """Delete a column by its header name from the sheet.

    Args:
        file_path: Path to the Excel .xlsx file.
        column_name: Column header name to delete.
        sheet_name: Sheet name (defaults to first sheet).
    """
    try:
        wb = _open_workbook(file_path)
        ws = _get_sheet(wb, sheet_name)
        col_idx = _get_column_idx(ws, column_name)
        ws.delete_cols(col_idx)
        wb.save(file_path)
        wb.close()
        return f"Column '{column_name}' deleted from sheet '{ws.title}'"
    except Exception as e:
        return f"Error: {e}"


# =============================================
# Main
# =============================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Excel CRUD MCP Server")
    parser.add_argument(
        "--host",
        type=str,
        default="0.0.0.0",
        help="Host to bind the server to",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Port to bind the server to",
    )
    parser.add_argument(
        "--transport",
        type=str,
        default="streamable-http",
        help="Transport protocol (default: streamable-http)",
    )
    args = parser.parse_args()

    mcp.run(transport=args.transport, host=args.host, port=args.port)
