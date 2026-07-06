from typing import Any, Dict, List
from icecream import ic
from mcp.types import ToolAnnotations
from msgraph.generated.models.workbook_worksheet import WorkbookWorksheet
from mcp_excel_online.core.models.mcp import ToolContext, Transport

from mcp_excel_online.core.mcp.config import mcp


@mcp.tool(
    annotations=ToolAnnotations(
        title="List Sheets",
        readOnlyHint=True,
    ),
)
async def list_sheets(workbook_id: str,
                      ctx: ToolContext = None) -> List[str] | None:
    """
    List all sheets in a workbook.

    Args:
        workbook_id: The ID of the workbook.

    Returns:
        List of sheet names
    """
    client = ctx.request_context.lifespan_context.graph_client
    workbook = await client.get_workbook_by_id(workbook_id)
    if not workbook:
        return None
    sheets = await workbook.worksheets.get()
    sheet_names = [
        sheet.name for sheet in sheets.value] if sheets and sheets.value else []
    return sheet_names


@mcp.tool(
    annotations=ToolAnnotations(
        title="Rename Worksheet",
        destructiveHint=True,
    ),
)
async def rename_sheet(workbook_id: str, sheet_name: str,
                       new_name: str, ctx: ToolContext = None) -> Dict[str, Any] | None:
    """Rename a worksheet.

    Args:
        workbook_id: The ID of the workbook.
        sheet_name: Current name (or id) of the sheet.
        new_name: The desired new name.
    """
    client = ctx.request_context.lifespan_context.graph_client
    workbook = await client.get_workbook_by_id(workbook_id)
    if not workbook:
        return None
    worksheet = workbook.worksheets.by_workbook_worksheet_id(sheet_name)
    body = WorkbookWorksheet()
    body.name = new_name
    result = await worksheet.patch(body)
    return {
        "success": True,
        "updated_name": result.name if result else None,
    }


@mcp.tool(
    annotations=ToolAnnotations(
        title="Create Worksheet",
        destructiveHint=True,
    ),
)
async def create_sheet(workbook_id: str, sheet_name: str,
                       ctx: ToolContext = None) -> Dict[str, Any] | None:
    """Create a new worksheet.

    Args:
        workbook_id: The ID of the workbook.
        sheet_name: Name for the new sheet.

    Returns:
        The created WorkbookWorksheet object.
    """
    client = ctx.request_context.lifespan_context.graph_client
    workbook = await client.get_workbook_by_id(workbook_id)
    if not workbook:
        return None

    body = WorkbookWorksheet()
    body.name = sheet_name
    result = await workbook.worksheets.post(body)
    return {
        "success": True,
        "sheet_name": result.name if result else None,
        "sheet_id": result.id if result else None,
    }


@mcp.tool(
    annotations=ToolAnnotations(
        title="Delete Worksheet",
        destructiveHint=True,
    ),
)
async def delete_sheet(workbook_id: str, sheet_name: str, ctx: ToolContext = None) -> Dict[str, Any] | None:
    """Delete a worksheet.

    Args:
        workbook_id: The ID of the workbook.
        sheet_name: Name (or id) of the sheet to delete.
    """
    client = ctx.request_context.lifespan_context.graph_client
    workbook = await client.get_workbook_by_id(workbook_id)
    if not workbook:
        return None

    await workbook.worksheets.by_workbook_worksheet_id(sheet_name).delete()
    return {
        "success": True
    }


@mcp.tool(
    annotations=ToolAnnotations(
        title="Copy Worksheet",
        destructiveHint=True,
    ),
)
async def copy_sheet(src_workbook_id: str, src_sheet_name: str, dst_workbook_id: str, dst_sheet_name: str, ctx: ToolContext = None) -> Dict[str, Any] | None:
    """Copy a worksheet.

    Args:
        src_workbook_id: The ID of the source workbook.
        src_sheet_name: Name (or id) of the source sheet to copy.
        dst_workbook_id: The ID of the destination workbook.
        dst_sheet_name: Name for the new sheet in the destination workbook.
    """
    client = ctx.request_context.lifespan_context.graph_client
    src_workbook = await client.get_workbook_by_id(src_workbook_id)
    dst_workbook = await client.get_workbook_by_id(dst_workbook_id)
    if (not src_workbook or not dst_workbook):
        return {"success": False, "error": "Source or destination workbook not found"}

    src_worksheet = src_workbook.worksheets.by_workbook_worksheet_id(
        src_sheet_name)
    src_workbook_range = await src_worksheet.used_range.get()
    if not src_workbook_range or not src_workbook_range.additional_data:
        return {"success": False, "error": "Source worksheet not found"}

    # First create a new worksheet in the destination workbook with the desired name
    body = WorkbookWorksheet()
    body.name = dst_sheet_name
    created_worksheet = await dst_workbook.worksheets.post(body)
    # Now copy the data from the source worksheet to the new destination worksheet
    range = src_workbook_range.address.split(
        "!")[1] if "!" in src_workbook_range.address else src_workbook_range.address
    ic(range, src_workbook_range)
    if not range:
        return {"success": False, "error": "Source worksheet has no data to copy"}
    body = {
        "formulas": src_workbook_range.additional_data.get("formulas"),
        "formulasLocal": src_workbook_range.additional_data.get("formulasLocal"),
        "formulasR1C1": src_workbook_range.additional_data.get("formulasR1C1"),
        "numberFormat": src_workbook_range.additional_data.get("numberFormat"),
        "text": src_workbook_range.additional_data.get("text"),
        "valueTypes": src_workbook_range.additional_data.get("valueTypes"),
        "values": src_workbook_range.additional_data.get("values"),
    }
    result = await client.patch_worksheet_data_in_range(workbook=dst_workbook, sheet_name=dst_sheet_name, range=range, body=body)
    if not result:
        return {"success": False, "error": "Failed to copy data to destination worksheet"}

    return {
        "success": True,
        "new_sheet_name": created_worksheet.name if created_worksheet else None,
        "new_sheet_id": created_worksheet.id if created_worksheet else None,
    }
