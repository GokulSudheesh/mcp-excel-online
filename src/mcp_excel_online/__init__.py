from .server import serve


def main():
    """MCP Excel Online Server entry point"""
    import argparse
    import asyncio

    parser = argparse.ArgumentParser(
        description="give an agent the ability to handle Excel Online operations")
    parser.add_argument("--transport", default="sse")
    args = parser.parse_args()
    asyncio.run(serve(args.transport))


if __name__ == "__main__":
    main()
