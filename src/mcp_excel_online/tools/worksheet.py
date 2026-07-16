from typing import Any, Dict, List, Optional
from mcp.types import ToolAnnotations
from pydantic import Field
from mcp_excel_online.core.models.mcp import ToolContext

from mcp_excel_online.core.mcp.config import mcp
from mcp_excel_online.core.models.tools import WORKBOOK_ID_FIELD, SHEET_NAME_FIELD


@mcp.tool(
    annotations=ToolAnnotations(
        title="Get Worksheet Data",
        readOnlyHint=True,
    ),
)
async def get_worksheet_data(workbook_id: str = WORKBOOK_ID_FIELD,
                             sheet_name: str = SHEET_NAME_FIELD,
                             range: Optional[str] = None,
                             ctx: ToolContext = None) -> List[List[Any]] | None:
    """
    Get data from a specific worksheet and range.

    Args:
        workbook_id: The ID of the workbook.
        sheet_name: The name of the sheet
        range: Optional cell range in A1 notation (e.g., 'A1:C10'). If not provided, gets all data.

    Returns:
        A 2D array of the sheet formulas.
    """
    client = ctx.request_context.lifespan_context.graph_client
    workbook = await client.get_workbook_by_id(workbook_id)
    if not workbook:
        return None

    worksheet = workbook.worksheets.by_workbook_worksheet_id(sheet_name)
    if range:
        data = await worksheet.range_with_address(range).get()
    else:
        data = await worksheet.used_range.get()
    data = data.additional_data.get("values") if data else None
    return data


@mcp.tool(
    annotations=ToolAnnotations(
        title="Get Worksheet Formulas",
        readOnlyHint=True,
    ),
)
async def get_worksheet_formulas(workbook_id: str = WORKBOOK_ID_FIELD,
                                 sheet_name: str = SHEET_NAME_FIELD,
                                 range: Optional[str] = None,
                                 ctx: ToolContext = None) -> List[List[Any]] | None:
    """
    Get formulas from a specific worksheet and range.

    Args:
        workbook_id: The ID of the workbook.
        sheet_name: The name of the sheet
        range: Optional cell range in A1 notation (e.g., 'A1:C10'). If not provided, gets all formulas from the sheet.

    Returns:
        A 2D array of the sheet formulas.
    """
    client = ctx.request_context.lifespan_context.graph_client
    workbook = await client.get_workbook_by_id(workbook_id)
    if not workbook:
        return None

    worksheet = workbook.worksheets.by_workbook_worksheet_id(sheet_name)
    if range:
        data = await worksheet.range_with_address(range).get()
    else:
        data = await worksheet.used_range.get()
    data = data.additional_data.get("formulas") if data else None
    return data


@mcp.tool(
    annotations=ToolAnnotations(
        title="Update Worksheet Data",
        destructiveHint=True,
    ),
)
async def update_worksheet_data(workbook_id: str = WORKBOOK_ID_FIELD,
                                sheet_name: str = SHEET_NAME_FIELD,
                                range: str = Field(
                                    description="Cell range in A1 notation (e.g., 'A1:C10')"),
                                data: List[List[str]] = Field(
                                    description="A 2D array of values to update the specified range with."),
                                ctx: ToolContext = None) -> Dict[str, Any]:
    """
    Update data in a specific worksheet and range.

    Args:
        workbook_id: The ID of the workbook.
        sheet_name: The name of the sheet
        range: Cell range in A1 notation (e.g., 'A1:C10')
        data: A 2D array of values to update the specified range with.
    """
    client = ctx.request_context.lifespan_context.graph_client
    workbook = await client.get_workbook_by_id(workbook_id)
    if not workbook:
        return {"success": False, "error": "Workbook not found"}
    body = {"values": data}
    result = await client.patch_worksheet_data_in_range(workbook=workbook, sheet_name=sheet_name,
                                                        range=range, body=body)

    return {
        "success": True if result is not None else False,
        "updated_range": result.address if result else None,
        "updated_values": result.additional_data.get("values") if result and result.additional_data else None
    }
