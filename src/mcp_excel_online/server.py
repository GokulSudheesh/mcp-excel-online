import json
from typing import Any, Dict, List, Literal, Optional

from icecream import ic
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
from msgraph.generated.drives.item.items.item.workbook.workbook_request_builder import WorkbookRequestBuilder
from contextlib import asynccontextmanager
from collections.abc import AsyncIterator
from msgraph.graph_service_client import GraphServiceClient
from msgraph.generated.models.workbook_range import WorkbookRange
from msgraph.generated.models.workbook_worksheet import WorkbookWorksheet

from kiota_abstractions.method import Method
from kiota_abstractions.request_information import RequestInformation

from mcp_excel_online.core.config import Settings
from mcp_excel_online.core.graph_sdk.client_manager import GraphClientManager
from mcp_excel_online.core.models.mcp import SpreadsheetContext, ToolContext, Transport


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


async def get_workbook_by_id(workbook_id: str, ctx: ToolContext) -> WorkbookRequestBuilder | None:
    """Get the workbook from a specific drive item by its ID."""
    client = ctx.request_context.lifespan_context.graph_client
    drive_item = client.drives.by_drive_id(
        Settings.DRIVE_ID).items.by_drive_item_id(workbook_id)
    workbook = drive_item.workbook if drive_item else None
    return workbook


mcp = FastMCP(
    name="Excel Online",
    lifespan=workbook_lifespan,
)


@mcp.tool(
    annotations=ToolAnnotations(
        title="Get Worksheet Data",
        readOnlyHint=True,
    ),
)
async def get_worksheet_data(workbook_id: str,
                             sheet_name: str,
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
    workbook = await get_workbook_by_id(workbook_id, ctx)
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
async def get_worksheet_formulas(workbook_id: str,
                                 sheet_name: str,
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
    workbook = await get_workbook_by_id(workbook_id, ctx)
    if not workbook:
        return None

    worksheet = workbook.worksheets.by_workbook_worksheet_id(sheet_name)
    if range:
        data = await worksheet.range_with_address(range).get()
    else:
        data = await worksheet.used_range.get()
    data = data.additional_data.get("formulas") if data else None
    return data


async def patch_worksheet_data_in_range(
        *,
        client: GraphServiceClient,
        workbook: WorkbookRequestBuilder,
        sheet_name: str,
        range: str,
        body: Dict) -> WorkbookRange:
    worksheet = workbook.worksheets.by_workbook_worksheet_id(sheet_name)
    range_builder = worksheet.range_with_address(range)

    # RangeWithAddressRequestBuilder only generates get(); build the PATCH manually
    request_info = RequestInformation(
        Method.PATCH, range_builder.url_template, range_builder.path_parameters)
    request_info.headers.try_add("Content-Type", "application/json")

    request_info.content = json.dumps(body).encode("utf-8")
    result = await client.request_adapter.send_async(request_info, WorkbookRange, {})
    return result


@mcp.tool(
    annotations=ToolAnnotations(
        title="Update Worksheet Data",
        destructiveHint=True,
    ),
)
async def update_worksheet_data(workbook_id: str,
                                sheet_name: str,
                                range: str, data: List[List[str]],
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
    workbook = await get_workbook_by_id(workbook_id, ctx)
    if not workbook:
        return {"success": False, "error": "Workbook not found"}
    body = {"values": data}
    result = await patch_worksheet_data_in_range(client=client, workbook=workbook, sheet_name=sheet_name,
                                                 range=range, body=body)

    return {
        "success": True if result is not None else False,
        "updated_range": result.address if result else None,
        "updated_values": result.additional_data.get("values") if result and result.additional_data else None
    }


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
    workbook = await get_workbook_by_id(workbook_id, ctx)
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
    workbook = await get_workbook_by_id(workbook_id, ctx)
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
    workbook = await get_workbook_by_id(workbook_id, ctx)
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
    workbook = await get_workbook_by_id(workbook_id, ctx)
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
    src_workbook = await get_workbook_by_id(src_workbook_id, ctx)
    dst_workbook = await get_workbook_by_id(dst_workbook_id, ctx)
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
    result = await patch_worksheet_data_in_range(client=client, workbook=dst_workbook, sheet_name=dst_sheet_name, range=range, body=body)
    if not result:
        return {"success": False, "error": "Failed to copy data to destination worksheet"}

    return {
        "success": True,
        "new_sheet_name": created_worksheet.name if created_worksheet else None,
        "new_sheet_id": created_worksheet.id if created_worksheet else None,
    }


def serve(transport: Transport = "sse"):
    mcp.run(transport=transport)
