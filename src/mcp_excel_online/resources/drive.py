import json
from kiota_serialization_json.json_serialization_writer import JsonSerializationWriter
from pydantic import Field

from mcp_excel_online.core.config import Settings
from mcp_excel_online.core.mcp.config import mcp
from mcp_excel_online.tools.helper import get_drive_children


@mcp.resource(
    "drive://root/children",
    name="Drive Root Items",
    description="List all items at the root of the configured OneDrive.",
    mime_type="application/json",
)
async def list_drive_items() -> str:
    """GET /me/drive/root/children"""
    context = mcp.get_context()
    client = context.request_context.lifespan_context.graph_client
    data = await get_drive_children(client, "/me/drive/root/children")
    writer = JsonSerializationWriter()
    data.serialize(writer)
    json_data = json.loads(writer.get_serialized_content())
    return json.dumps(json_data, indent=2)


@mcp.resource(
    uri="drive://drives/{drive_id}/root/{relative_path}/children",
    name="Folder Items",
    description="List the children based on the Path relative to the root. Pass an empty string for drive_id to use the default drive.",
    mime_type="application/json",
)
async def list_folder_items(drive_id: str = Field(description="The ID of the drive to use."),
                            relative_path: str = Field(description="The path relative to the root.")) -> str:
    """GET /drives/{drive-id}/root:/{path-relative-to-root}:/children"""
    context = mcp.get_context()
    client = context.request_context.lifespan_context.graph_client
    resolved_drive_id = drive_id or Settings.DRIVE_ID
    data = await get_drive_children(client, f"/drives/{resolved_drive_id}/root:/{relative_path}:/children")
    writer = JsonSerializationWriter()
    data.serialize(writer)
    json_data = json.loads(writer.get_serialized_content())
    return json.dumps(json_data, indent=2)
