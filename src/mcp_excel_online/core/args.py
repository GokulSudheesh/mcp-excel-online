import argparse
from typing import get_args

from mcp_excel_online.core.models.mcp import Transport
from mcp_excel_online.core.models.graph_api import PermissionType


class Args(argparse.Namespace):
    transport: Transport
    graph_permission: PermissionType


_parser = argparse.ArgumentParser(
    description="MCP Excel Online server"
)
_parser.add_argument(
    "--transport",
    default="sse",
    choices=list(get_args(Transport.__value__)),
)
_parser.add_argument(
    "--graph-permission",
    dest="graph_permission",
    default="delegated",
    choices=list(get_args(PermissionType.__value__)),
    help=(
        "'delegated' — user sign-in via device code (default); "
        "'application' — service principal via client secret (APP_TENANT_ID + APP_CLIENT_SECRET required)"
    ),
)

args: Args = _parser.parse_args(namespace=Args())
