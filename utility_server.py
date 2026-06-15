import argparse
import json
from datetime import datetime, timezone

from mcp.server.fastmcp import FastMCP
import httpx

mcp = FastMCP("utility-server")

WIKIPEDIA_API_URL = "https://en.wikipedia.org/w/api.php"


@mcp.tool()
def get_current_datetime() -> str:
    """Get the current date and time information.

    Returns:
        Current date and time in multiple formats including ISO format,
        Unix timestamp, date, and time components.
    """
    now = datetime.now(timezone.utc)
    return json.dumps(
        {
            "iso_8601": now.isoformat(),
            "unix_timestamp": int(now.timestamp()),
            "date": now.strftime("%Y-%m-%d"),
            "time": now.strftime("%H:%M:%S"),
            "timezone": "UTC",
            "day_of_week": now.strftime("%A"),
            "day_of_year": now.timetuple().tm_yday,
        },
        indent=2,
    )


@mcp.tool()
def get_current_ip() -> str:
    """Get the current public IP address of the server.

    Returns:
        The public IP address and related location/ISP information.
    """
    try:
        response = httpx.get("https://api.ipify.org?format=json", timeout=10)
        response.raise_for_status()
        data = response.json()

        detail_resp = httpx.get(
            f"https://ipapi.co/{data['ip']}/json/", timeout=10
        )
        if detail_resp.status_code == 200:
            detail = detail_resp.json()
            return json.dumps(
                {
                    "ip": data["ip"],
                    "location": {
                        "city": detail.get("city"),
                        "region": detail.get("region"),
                        "country": detail.get("country_name"),
                        "latitude": detail.get("latitude"),
                        "longitude": detail.get("longitude"),
                    },
                    "isp": detail.get("org"),
                },
                indent=2,
            )

        return json.dumps({"ip": data["ip"]}, indent=2)
    except Exception as e:
        return json.dumps(
            {"error": f"Failed to get IP: {str(e)}"}, indent=2
        )


@mcp.tool()
def read_wikipedia_page(page_title: str) -> str:
    """Read the content of a Wikipedia page.

    Args:
        page_title: The title of the Wikipedia page to read
                    (e.g., "Python (programming language)").

    Returns:
        The page extract/summary content from Wikipedia.
    """
    params = {
        "action": "query",
        "format": "json",
        "titles": page_title,
        "prop": "extracts",
        "explaintext": True,
        "exlimit": 1,
    }
    try:
        resp = httpx.get(WIKIPEDIA_API_URL, params=params, timeout=15)
        resp.raise_for_status()
        data = resp.json()

        pages = data.get("query", {}).get("pages", {})
        for page_id, page_data in pages.items():
            if page_id == "-1":
                return json.dumps(
                    {"error": f"Page '{page_title}' not found on Wikipedia"},
                    indent=2,
                )
            return json.dumps(
                {
                    "title": page_data.get("title"),
                    "page_id": page_id,
                    "extract": page_data.get("extract", ""),
                },
                indent=2,
                ensure_ascii=False,
            )

        return json.dumps({"error": f"Page '{page_title}' not found"}, indent=2)
    except Exception as e:
        return json.dumps(
            {"error": f"Failed to read Wikipedia page: {str(e)}"},
            indent=2,
        )


@mcp.tool()
def search_wikipedia(query: str, limit: int = 5) -> str:
    """Search Wikipedia for pages matching the given query.

    Args:
        query: The search query string.
        limit: Maximum number of search results to return
               (default: 5, max: 20).

    Returns:
        A list of search results with titles and descriptions.
    """
    limit = min(max(1, limit), 20)
    params = {
        "action": "query",
        "format": "json",
        "list": "search",
        "srsearch": query,
        "srlimit": limit,
        "srprop": "snippet",
    }
    try:
        resp = httpx.get(WIKIPEDIA_API_URL, params=params, timeout=15)
        resp.raise_for_status()
        data = resp.json()

        results = []
        for item in data.get("query", {}).get("search", []):
            results.append(
                {
                    "title": item.get("title"),
                    "page_id": item.get("pageid"),
                    "snippet": item.get("snippet", ""),
                }
            )

        return json.dumps(
            {
                "query": query,
                "total_results": data.get("query", {})
                .get("searchinfo", {})
                .get("totalhits", 0),
                "results": results,
            },
            indent=2,
            ensure_ascii=False,
        )
    except Exception as e:
        return json.dumps(
            {"error": f"Failed to search Wikipedia: {str(e)}"},
            indent=2,
        )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Utility MCP Server - datetime, IP, Wikipedia"
    )
    parser.add_argument(
        "--transport",
        choices=["stdio", "http"],
        default="stdio",
        help="Transport: 'stdio' (default) or 'http' for Streamable HTTP",
    )
    parser.add_argument(
        "--host",
        default="0.0.0.0",
        help="Host to bind when using HTTP transport (default: 0.0.0.0)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8001,
        help="Port for HTTP transport (default: 8001)",
    )
    args = parser.parse_args()

    if args.transport == "http":
        try:
            mcp.run(
                transport="streamable-http",
                host=args.host,
                port=args.port,
            )
        except TypeError:
            mcp.run(transport="streamable-http")
    else:
        mcp.run(transport="stdio")
