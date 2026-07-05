from typing import Literal, Optional

from mcp.server.fastmcp import Context
from mcp.server.session import ServerSession
from dataclasses import dataclass
from msgraph.graph_service_client import GraphServiceClient


@dataclass
class SpreadsheetContext:
    """Context for Microsoft Graph API service"""
    graph_client: GraphServiceClient
    folder_id: Optional[str] = None


ToolContext = Context[ServerSession, SpreadsheetContext]

type Transport = Literal['stdio', 'sse', 'streamable-http']
