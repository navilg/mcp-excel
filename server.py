import argparse
import io
import pandas as pd
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("excel-reader")


def read_sheet_as_csv(file_path: str, sheet_name: str) -> str:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    buffer = io.StringIO()
    df.to_csv(buffer, index=False)
    return buffer.getvalue()


def read_all_sheets_as_csv(file_path: str) -> str:
    sheets = pd.read_excel(file_path, sheet_name=None)
    parts = []
    for name, df in sheets.items():
        buffer = io.StringIO()
        df.to_csv(buffer, index=False)
        parts.append(f"### Sheet: {name}\n{buffer.getvalue()}")
    return "\n".join(parts)


@mcp.tool()
def read_excel(file_path: str, sheet_name: str = "") -> str:
    """Read an Excel file and return its contents as CSV.

    Args:
        file_path: Absolute or relative path to the Excel file (.xlsx or .xls).
        sheet_name: Name of the sheet to read. If empty, all sheets are returned.
    """
    if sheet_name:
        return read_sheet_as_csv(file_path, sheet_name)
    return read_all_sheets_as_csv(file_path)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Excel MCP Server")
    parser.add_argument(
        "--transport",
        choices=["stdio", "http"],
        default="stdio",
        help="Transport to use: 'stdio' (default) or 'http' for Streamable HTTP",
    )
    parser.add_argument(
        "--host",
        default="0.0.0.0",
        help="Host to bind when using HTTP transport (default: 0.0.0.0)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Port to listen on when using HTTP transport (default: 8000)",
    )
    args = parser.parse_args()

    if args.transport == "http":
        mcp.run(transport="streamable-http", host=args.host, port=args.port)
    else:
        mcp.run(transport="stdio")
