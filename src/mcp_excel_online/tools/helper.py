import json
from typing import Dict

from msgraph.generated.drives.item.items.item.workbook.workbook_request_builder import WorkbookRequestBuilder
from msgraph.graph_service_client import GraphServiceClient
from msgraph.generated.models.workbook_range import WorkbookRange

from kiota_abstractions.method import Method
from kiota_abstractions.request_information import RequestInformation

from mcp_excel_online.core.config import Settings
from mcp_excel_online.core.models.mcp import ToolContext


async def get_workbook_by_id(workbook_id: str, ctx: ToolContext) -> WorkbookRequestBuilder | None:
    """Get the workbook from a specific drive item by its ID."""
    client = ctx.request_context.lifespan_context.graph_client
    drive_item = client.drives.by_drive_id(
        Settings.DRIVE_ID).items.by_drive_item_id(workbook_id)
    workbook = drive_item.workbook if drive_item else None
    return workbook


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
