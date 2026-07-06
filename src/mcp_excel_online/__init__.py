from .server import serve
from mcp_excel_online.core.args import args


def main() -> None:
    """MCP Excel Online server entry point"""
    serve(args.transport)


if __name__ == "__main__":
    main()
