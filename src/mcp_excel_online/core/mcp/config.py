from mcp.server.fastmcp import FastMCP
from contextlib import asynccontextmanager
from collections.abc import AsyncIterator

from mcp_excel_online.core.config import Settings
from mcp_excel_online.core.graph_sdk.client_manager import GraphClientManager
from mcp_excel_online.core.models.mcp import SpreadsheetContext


@asynccontextmanager
async def workbook_lifespan(server: FastMCP) -> AsyncIterator[SpreadsheetContext]:
    """Manage Microsoft Graph API connection lifecycle"""
    manager = await GraphClientManager.get_instance()
    try:
        # Provide the client in the context
        yield SpreadsheetContext(
            graph_client=manager.client,
            folder_id=Settings.DRIVE_ID
        )
    finally:
        # No explicit cleanup needed for Microsoft Graph API
        pass

mcp = FastMCP[SpreadsheetContext](
    name="Excel Online",
    lifespan=workbook_lifespan,
    port=Settings.PORT,
    host=Settings.HOST,
)
