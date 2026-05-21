import argparse
import base64
import io
import pandas as pd
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("excel-reader")


def _decode_excel_content(excel_content: str) -> io.BytesIO:
    try:
        raw_bytes = base64.b64decode(excel_content, validate=True)
    except Exception as exc:
        raise ValueError(
            "excel_content must be a valid base64-encoded Excel file payload"
        ) from exc
    return io.BytesIO(raw_bytes)


def read_sheet_as_csv(excel_content: str, sheet_name: str) -> str:
    source = _decode_excel_content(excel_content)
    df = pd.read_excel(source, sheet_name=sheet_name)
    buffer = io.StringIO()
    df.to_csv(buffer, index=False)
    return buffer.getvalue()


def read_all_sheets_as_csv(excel_content: str) -> str:
    source = _decode_excel_content(excel_content)
    sheets = pd.read_excel(source, sheet_name=None)
    parts = []
    for name, df in sheets.items():
        buffer = io.StringIO()
        df.to_csv(buffer, index=False)
        parts.append(f"### Sheet: {name}\n{buffer.getvalue()}")
    return "\n".join(parts)


@mcp.tool()
def read_excel(excel_content: str, sheet_name: str = "") -> str:
    """Read Excel content and return its contents as CSV.

    Args:
        excel_content: Base64-encoded raw content of the Excel file (.xlsx or .xls).
        sheet_name: Name of the sheet to read. If empty, all sheets are returned.
    """
    if sheet_name:
        return read_sheet_as_csv(excel_content, sheet_name)
    return read_all_sheets_as_csv(excel_content)


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
        # FastMCP.run signature varies across versions.
        try:
            mcp.run(transport="streamable-http", host=args.host, port=args.port)
        except TypeError:
            mcp.run(transport="streamable-http")
    else:
        mcp.run(transport="stdio")
