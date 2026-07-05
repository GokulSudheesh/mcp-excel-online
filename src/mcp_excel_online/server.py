from mcp_excel_online.core.graph_sdk.client_manager import GraphClientManager
from mcp_excel_online.core.models.mcp import Transport
from mcp_excel_online.core.mcp.config import mcp
import mcp_excel_online.tools.workbook  # noqa: F401
import mcp_excel_online.tools.worksheet  # noqa: F401


def serve(transport: Transport = "sse") -> None:
    mcp.run(transport=transport)
