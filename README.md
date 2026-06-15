# mcp-excel

MCP server to read Excel data as CSV.

## Servers

This repository contains two MCP servers:

### 1. Excel Reader (`server.py`)

Read Excel (.xlsx / .xls) files and return contents as CSV.

**Tools:**
- `read_excel(excel_content, sheet_name="")` — Read a base64-encoded Excel file. Returns all sheets (or a specific sheet) as CSV.

**Usage:**

```bash
# stdio transport (default)
python server.py

# Streamable HTTP transport
python server.py --transport http --host 0.0.0.0 --port 8000
```

---

### 2. Utility Server (`utility_server.py`)

General-purpose MCP server with datetime, IP, and Wikipedia tools.

**Tools:**

| Tool | Description |
|------|-------------|
| `get_current_datetime()` | Returns the current UTC date/time in ISO 8601, Unix timestamp, date, time, day of week, and day of year |
| `get_current_ip()` | Returns the server's public IP address with optional geo-location and ISP details |
| `read_wikipedia_page(page_title)` | Reads and returns the extract of a Wikipedia page by title |
| `search_wikipedia(query, limit=5)` | Searches Wikipedia and returns matching page titles and snippets |

**Usage:**

```bash
# stdio transport (default)
python utility_server.py

# Streamable HTTP transport
python utility_server.py --transport http --host 0.0.0.0 --port 8001
```

---

## Docker

### Excel server

```bash
docker build -t mcp-excel -f Dockerfile .
docker run -p 8000:8000 mcp-excel --transport http --host 0.0.0.0 --port 8000
```

### Utility server

```bash
docker build -t mcp-utility -f Dockerfile.utility .
docker run -p 8001:8001 mcp-utility --transport http --host 0.0.0.0 --port 8001
```

### Both servers via docker-compose

```bash
docker-compose up --build
```

The Excel server runs on **port 8000** and the Utility server runs on **port 8001**.

### Transport options

Both servers support two transports:
- **stdio** — Standard input/output (default, for local MCP clients)
- **http** — Streamable HTTP transport for remote connections

When using HTTP transport with MCP clients, configure the endpoint URL as:
- `http://<host>:8000/mcp` (Excel server)
- `http://<host>:8001/mcp` (Utility server)
