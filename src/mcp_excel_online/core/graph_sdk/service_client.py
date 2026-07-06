import json
from typing import Dict

from msgraph.generated.drives.item.items.item.workbook.workbook_request_builder import WorkbookRequestBuilder
from msgraph.generated.models.drive_item_collection_response import DriveItemCollectionResponse
from msgraph.generated.models.workbook_range import WorkbookRange
from kiota_abstractions.method import Method
from kiota_abstractions.request_information import RequestInformation

from mcp_excel_online.core.config import Settings
from msgraph.graph_service_client import GraphServiceClient


class ServiceClient(GraphServiceClient):
    """
    A service client for interacting with Microsoft Graph API.
    """

    def __init__(self, **kwargs):
        """
        Initializes a new instance of the ServiceClient class.

        :param request_adapter: The request adapter to use for making HTTP requests.
        """
        super().__init__(**kwargs)

    async def get_workbook_by_id(self, workbook_id: str) -> WorkbookRequestBuilder | None:
        """Get the workbook from a specific drive item by its ID."""
        drive_item = self.drives.by_drive_id(
            Settings.DRIVE_ID).items.by_drive_item_id(workbook_id)
        workbook = drive_item.workbook if drive_item else None
        return workbook

    async def patch_worksheet_data_in_range(
            self,
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
        result = await self.request_adapter.send_async(request_info, WorkbookRange, {})
        return result

    async def get_drive_children(self, url: str) -> DriveItemCollectionResponse | None:
        """GET any OneDrive /children endpoint and return a list of DriveItems."""
        request_info = RequestInformation(
            Method.GET, f"{self.request_adapter.base_url}{url}", {})
        request_info.headers.try_add("Accept", "application/json")
        result = await self.request_adapter.send_async(request_info, DriveItemCollectionResponse, {})
        return result
