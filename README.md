# MCP Excel – CRUD Operations on Excel Files

An MCP (Model Context Protocol) server that provides tools to perform **CRUD operations** on Excel (`.xlsx`) files using [FastMCP](https://github.com/jlowin/fastmcp) and [openpyxl](https://openpyxl.readthedocs.io/).

## Tools

### Utility
| Tool | Description |
|------|-------------|
| `encode_base64` | Encode plain text to base64 (use to encode file content for other tools) |
| `decode_base64` | Decode a base64 string back to plain text (use to inspect write tool output) |

### Create (C)
| Tool | Description |
|------|-------------|
| `add_row` | Append a tab-separated row of values to the sheet |
| `add_cell_data` | Write a value into a specific cell by column name and row number |
| `add_column` | Add a new column with a header name and optional default value |
| `add_full_data` | Replace all sheet data with TSV-formatted content (headers + rows) |
| `create_sheet` | Create a new empty sheet in the workbook |

### Read (R)
| Tool | Description |
|------|-------------|
| `read_excel` | Read an entire sheet and return its content as TSV |
| `read_rows_by_column` | Find rows where a specific column matches a value |

### Update (U)
| Tool | Description |
|------|-------------|
| `update_row` | Replace the values of an entire row |
| `update_cell` | Update a single cell by column name and row number |
| `replace_full_data` | Replace all data in a sheet (alias for `add_full_data`) |

### Delete (D)
| Tool | Description |
|------|-------------|
| `delete_sheet` | Delete a sheet from the workbook |
| `delete_row` | Delete a row from the sheet |
| `delete_column` | Delete a column by its header name |

## Common Parameters

- **`file_content`** – Base64-encoded content of the `.xlsx` file. Use `encode_base64` to convert a file's text content to base64 first.
- **`sheet_name`** – (Optional) Name of the target sheet. Defaults to the **first sheet** in the workbook if omitted.

## What Each Tool Returns

| Tool type | Returns |
|-----------|---------|
| **Utility** (`encode_base64`, `decode_base64`) | Plain text |
| **Read** (e.g. `read_excel`) | **TSV text** – the sheet contents |
| **Write** (Create / Update / Delete) | **Base64 encoded workbook** – the modified file |
| **Error** | Plain text starting with `Error: ` |

## Data Flow (Remote Setup)

```
User attaches Excel file in chat
       ↓
LLM calls encode_base64 with file content → gets base64 string
       ↓
LLM passes base64 to any CRUD tool as file_content
       ↓
Read tool → returns TSV text for display
Write tool → returns base64 of modified workbook
       ↓
LLM can use decode_base64 to inspect the output, or
pass the base64 directly to the user to save as .xlsx
```

## Usage

Start the server:

```bash
pip install -r requirements.txt
python server.py
```

For HTTP transport (default):

```bash
python server.py --transport streamable-http --host 0.0.0.0 --port 8000
```

For stdio transport (for use with MCP clients like Claude Desktop):

```bash
python server.py --transport stdio
```

## Output Format Details

- **TSV** – Read operations return data as tab-separated values
- **Plaintext** – Error messages always start with `Error: `
- Special characters `\n` (newline) and `\t` (tab) in cell values are escaped to literal `\n` and `\t` to prevent broken formatting

## Docker

```bash
docker build -t mcp-excel .
docker run -p 8000:8000 mcp-excel
```