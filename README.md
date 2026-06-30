# MCP Excel – CRUD Operations on Excel Files

An MCP (Model Context Protocol) server that provides tools to perform **CRUD operations** on Excel (`.xlsx`) files using [FastMCP](https://github.com/jlowin/fastmcp) and [openpyxl](https://openpyxl.readthedocs.io/).

## Tools

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

- **`file_path`** – Path to the `.xlsx` file on disk. When you attach an Excel file in chat, the file path is automatically passed.
- **`sheet_name`** – (Optional) Name of the target sheet. Defaults to the **first sheet** in the workbook if omitted.

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

## Output Format

- **TSV** – Read operations return data as tab-separated values
- **Plaintext** – Success/error messages are returned as plain text
- Special characters `\n` (newline) and `\t` (tab) in cell values are escaped to literal `\n` and `\t`

## Docker

```bash
docker build -t mcp-excel .
docker run -p 8000:8000 mcp-excel
```
