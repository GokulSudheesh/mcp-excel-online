from typing import Literal, Optional

from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from dataclasses import dataclass
from mcp_excel_online.core.graph_sdk.service_client import ServiceClient


@dataclass
class SpreadsheetContext:
    """Context for Microsoft Graph API service"""
    graph_client: ServiceClient
    folder_id: Optional[str] = None


ToolContext = Context[ServerSession, SpreadsheetContext]

type Transport = Literal['stdio', 'sse', 'streamable-http']
