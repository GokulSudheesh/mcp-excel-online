import argparse
import ast
import logging
import os
from typing import get_args, Optional
from mcp_excel_online.core.models.mcp import Transport
from mcp_excel_online.core.models.graph_api import PermissionType

_TOOLS_DIR = os.path.normpath(os.path.join(
    os.path.dirname(__file__), "..", "tools"))


def _discover_tools() -> list[str]:
    """Parse tool source files via AST — no imports, no circular dependency."""
    names: list[str] = []
    for filename in sorted(os.listdir(_TOOLS_DIR)):
        if not filename.endswith(".py") or filename.startswith("_"):
            continue
        with open(os.path.join(_TOOLS_DIR, filename)) as fh:
            tree = ast.parse(fh.read(), filename=filename)
        for node in ast.walk(tree):
            if isinstance(node, (ast.AsyncFunctionDef, ast.FunctionDef)):
                if not node.name.startswith("_"):
                    names.append(node.name)
    return names


class AppArgs(argparse.Namespace):
    transport: Transport
    graph_permission: PermissionType
    include_tools: Optional[list[str]]

    AVAILABLE_TOOLS: list[str] = _discover_tools()

    @classmethod
    def _parse_include_tools(cls, value: str) -> list[str]:
        selected = [t.strip() for t in value.split(",") if t.strip()]
        invalid = [t for t in selected if t not in cls.AVAILABLE_TOOLS]
        if invalid:
            raise argparse.ArgumentTypeError(
                f"Unknown tool(s): {', '.join(invalid)}. "
                f"Available tools: {', '.join(cls.AVAILABLE_TOOLS)}"
            )
        return selected

    @classmethod
    def parse(cls) -> "AppArgs":
        logging.info(f"Available tools: {', '.join(cls.AVAILABLE_TOOLS)}")

        parser = argparse.ArgumentParser(description="MCP Excel Online server")
        parser.add_argument(
            "--transport",
            default="sse",
            choices=list(get_args(Transport.__value__)),
        )
        parser.add_argument(
            "--graph-permission",
            dest="graph_permission",
            default="delegated",
            choices=list(get_args(PermissionType.__value__)),
            help=(
                "'delegated' — user sign-in via device code (default); "
                "'application' — service principal via client secret "
                "(APP_TENANT_ID + APP_CLIENT_SECRET required)"
            ),
        )
        parser.add_argument(
            "--include-tools",
            dest="include_tools",
            default=None,
            type=cls._parse_include_tools,
            metavar="TOOL1,TOOL2,...",
            help=(
                "Comma-separated list of tools to enable. "
                f"Available Tools: {', '.join(cls.AVAILABLE_TOOLS)}. "
                "If omitted, all tools are enabled."
            ),
        )
        return parser.parse_args(namespace=cls())


args: AppArgs = AppArgs.parse()
