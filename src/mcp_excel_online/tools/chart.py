import json
from typing import List
from mcp.types import ToolAnnotations
from pydantic import Field
from mcp_excel_online.core.models.chart import Chart, ChartImageResponse, ChartList, ChartSeriesItem, ChartSeriesResponse, CreateChartSeriesRequest, Fitting_Mode, UpdateChartRequest, CreateChartRequest
from mcp_excel_online.core.models.mcp import ToolContext
from mcp_excel_online.core.mcp.config import mcp
from icecream import ic
from msgraph.generated.models.workbook_chart import WorkbookChart
from msgraph.generated.models.workbook_chart_series import WorkbookChartSeries
from kiota_serialization_json.json_serialization_writer import JsonSerializationWriter
from msgraph.generated.models.workbook_chart import WorkbookChart
from msgraph.generated.drives.item.items.item.workbook.worksheets.item.charts.add.add_post_request_body import AddPostRequestBody

from mcp_excel_online.core.models.tools import WORKBOOK_ID_FIELD, SHEET_NAME_FIELD


CHART_NAME_FIELD = Field(description="The name / id of the chart.")


@mcp.tool(
    annotations=ToolAnnotations(
        title="List Chart Collection",
        readOnlyHint=True,
    ),
)
async def list_chart_collection(workbook_id: str = WORKBOOK_ID_FIELD,
                                sheet_name: str = SHEET_NAME_FIELD,
                                ctx: ToolContext = None) -> List[Chart] | None:
    """
    Retrieve a list of chart objects.

    Args:
        workbook_id: The ID of the workbook.
        sheet_name: The name of the sheet.

    Returns:
        List of chart objects in the specified worksheet.
    """
    client = ctx.request_context.lifespan_context.graph_client
    workbook = await client.get_workbook_by_id(workbook_id)
    if not workbook:
        return None
    worksheet = workbook.worksheets.by_workbook_worksheet_id(sheet_name)
    data = await worksheet.charts.get()
    writer = JsonSerializationWriter()
    data.serialize(writer)
    json_data = json.loads(writer.get_serialized_content())
    ic(json_data)
    chart_list = ChartList.model_validate(json_data)
    return chart_list.value if chart_list and chart_list.value else None


@mcp.tool(
    annotations=ToolAnnotations(
        title="Get Chart",
        readOnlyHint=True,
    ),
)
async def get_chart(workbook_id: str = WORKBOOK_ID_FIELD,
                    sheet_name: str = SHEET_NAME_FIELD,
                    chart_name: str = CHART_NAME_FIELD,
                    ctx: ToolContext = None) -> Chart | None:
    """
    Retrieve a specific chart object.

    Args:
        workbook_id: The ID of the workbook.
        sheet_name: The name of the sheet.
        chart_name: The name / id of the chart.

    Returns:
        The specified chart object in the worksheet.
    """
    client = ctx.request_context.lifespan_context.graph_client
    workbook = await client.get_workbook_by_id(workbook_id)
    if not workbook:
        return None
    worksheet = workbook.worksheets.by_workbook_worksheet_id(sheet_name)
    data = await worksheet.charts.by_workbook_chart_id(chart_name).get()
    writer = JsonSerializationWriter()
    data.serialize(writer)
    json_data = json.loads(writer.get_serialized_content())
    ic(json_data)
    return Chart.model_validate(json_data)


@mcp.tool(
    annotations=ToolAnnotations(
        title="Create Chart",
        destructiveHint=True,
    ),
)
async def create_chart(workbook_id: str = WORKBOOK_ID_FIELD,
                       sheet_name: str = SHEET_NAME_FIELD,
                       data: CreateChartRequest = Field(
        description="An instance of CreateChartRequest containing the following properties:\
                        type: The type of chart to create. Valid values include:\n\
                        'ColumnClustered', 'ColumnStacked', 'ColumnStacked100',\n\
                        'BarClustered', 'BarStacked', 'BarStacked100',\n\
                        'LineStacked', 'LineStacked100', 'LineMarkers',\n\
                        'LineMarkersStacked', 'LineMarkersStacked100', 'PieOfPie', 'Pie'.\n\
                        source_data: The range of cells that contain the data for the chart, in A1 notation (e.g., 'A1:C10').\n\
                        series_by: Optional. Specifies the way columns or rows are used as data series on the chart. Valid values are: 'Auto' (default), 'Columns', or 'Rows'."),
        ctx: ToolContext = None) -> Chart | None:
    """
    Create a new chart object.

    Args:
        workbook_id: The ID of the workbook.
        sheet_name: The name of the sheet.
        data: An instance of CreateChartRequest containing the following properties:
            type: The type of chart to create. Valid values include:
                "ColumnClustered", "ColumnStacked", "ColumnStacked100",
                "BarClustered", "BarStacked", "BarStacked100",
                "LineStacked", "LineStacked100", "LineMarkers",
                "LineMarkersStacked", "LineMarkersStacked100", "PieOfPie", "Pie".
            source_data: The range of cells that contain the data for the chart, in A1 notation (e.g., 'A1:C10').
            series_by: Optional. Specifies the way columns or rows are used as data series on the chart. Valid values are:
                "Auto" (default), "Columns", or "Rows".

    Returns:
        The created chart object in the worksheet.
    """
    client = ctx.request_context.lifespan_context.graph_client
    workbook = await client.get_workbook_by_id(workbook_id)
    if not workbook:
        return None
    worksheet = workbook.worksheets.by_workbook_worksheet_id(sheet_name)
    request_body = AddPostRequestBody(type=data.type,
                                      additional_data={
                                          "name": data.name,
                                          "sourceData": data.source_data},
                                      series_by=data.series_by or "Auto",)
    result = await worksheet.charts.add.post(request_body)
    writer = JsonSerializationWriter()
    result.serialize(writer)
    json_data = json.loads(writer.get_serialized_content())
    return Chart.model_validate(json_data)


@mcp.tool(
    annotations=ToolAnnotations(
        title="Update Chart",
        destructiveHint=True,
    ),
)
async def update_chart(workbook_id: str = WORKBOOK_ID_FIELD,
                       sheet_name: str = SHEET_NAME_FIELD,
                       chart_name: str = CHART_NAME_FIELD,
                       data: UpdateChartRequest = Field(
                           description="An instance of UpdateChartRequest containing the properties to update. The properties include:\
                               name:    Represents the name of a chart object.\n\
                               height:  Represents the height, in points, of the chart object.\n\
                               width:   Represents the width, in points, of the chart object.\n\
                               left:    The distance, in points, from the left side of the chart to the worksheet origin.\n\
                               top:     Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart)."),
                       ctx: ToolContext = None) -> Chart | None:
    """
    Update a specific chart object.
    Args:
        workbook_id: The ID of the workbook.
        sheet_name: The name of the sheet.
        chart_name: The name / id of the chart.
        data: An instance of UpdateChartRequest containing the properties to update. The properties include:
            name:    Represents the name of a chart object.
            height:  Represents the height, in points, of the chart object.
            width:   Represents the width, in points, of the chart object.
            left:    The distance, in points, from the left side of the chart to the worksheet origin.
            top:     Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).

    Returns:
        The updated chart object in the worksheet.
    """
    client = ctx.request_context.lifespan_context.graph_client
    workbook = await client.get_workbook_by_id(workbook_id)
    if not workbook:
        return None
    worksheet = workbook.worksheets.by_workbook_worksheet_id(sheet_name)
    request_body = WorkbookChart(**data.model_dump(exclude_unset=True))
    result = await worksheet.charts.by_workbook_chart_id(chart_name).patch(request_body)
    writer = JsonSerializationWriter()
    result.serialize(writer)
    json_data = json.loads(writer.get_serialized_content())
    return Chart.model_validate(json_data)


@mcp.tool(
    annotations=ToolAnnotations(
        title="Delete Chart",
        destructiveHint=True,
    ),
)
async def delete_chart(workbook_id: str = WORKBOOK_ID_FIELD,
                       sheet_name: str = SHEET_NAME_FIELD,
                       chart_name: str = CHART_NAME_FIELD,
                       ctx: ToolContext = None) -> bool:
    """
    Delete a specific chart object.

    Args:
        workbook_id: The ID of the workbook.
        sheet_name: The name of the sheet.
        chart_name: The name / id of the chart.

    Returns:
        True if the chart was successfully deleted, False otherwise.
    """
    client = ctx.request_context.lifespan_context.graph_client
    workbook = await client.get_workbook_by_id(workbook_id)
    if not workbook:
        return False
    worksheet = workbook.worksheets.by_workbook_worksheet_id(sheet_name)
    await worksheet.charts.by_workbook_chart_id(chart_name).delete()
    return True


@mcp.tool(
    annotations=ToolAnnotations(
        title="Get Chart Image",
        readOnlyHint=True,
    ),
)
async def get_chart_image(workbook_id: str = WORKBOOK_ID_FIELD,
                          sheet_name: str = SHEET_NAME_FIELD,
                          chart_name: str = CHART_NAME_FIELD,
                          width: int = Field(
                              description="The desired width of the resulting image."),
                          height: int = Field(
                              description="The desired height of the resulting image."),
                          fitting_mode: Fitting_Mode = Field(
                              description="The method used to scale the chart to the specified dimensions (if both height and width are set). The possible values are: Fit, FitAndCenter, Fill."),
                          ctx: ToolContext = None) -> ChartImageResponse | None:
    """
    Retrieve the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.

    Args:
        workbook_id: The ID of the workbook.
        sheet_name: The name of the sheet.
        chart_name: The name / id of the chart.

    Returns:
        The image of the specified chart object as a base64-encoded string, or None if the chart does not exist.
    """
    client = ctx.request_context.lifespan_context.graph_client
    workbook = await client.get_workbook_by_id(workbook_id)
    if not workbook:
        return None
    worksheet = workbook.worksheets.by_workbook_worksheet_id(sheet_name)
    img_request_builder = worksheet.charts.by_workbook_chart_id(chart_name).image_with_width_with_height_with_fitting_mode(
        width=width,
        height=height,
        fitting_mode=fitting_mode
    )
    img_request_builder.url_template = "{+baseurl}/drives/{drive%2Did}/items/{driveItem%2Did}/workbook/worksheets/{workbookWorksheet%2Did}/charts/{workbookChart%2Did}/image(width={width}%2Cheight={height}%2CfittingMode='{fittingMode}')"
    image_data = await img_request_builder.get()
    if not image_data or not image_data.value:
        return None
    return ChartImageResponse(base64_image=image_data.value)


@mcp.tool(
    annotations=ToolAnnotations(
        title="Get Chart Series",
        readOnlyHint=True,
    ),
)
async def get_chart_series(workbook_id: str = WORKBOOK_ID_FIELD,
                           sheet_name: str = SHEET_NAME_FIELD,
                           chart_name: str = CHART_NAME_FIELD,
                           ctx: ToolContext = None) -> ChartSeriesResponse | None:
    """
    Retrieve the series of a specific chart object.

    Args:
        workbook_id: The ID of the workbook.
        sheet_name: The name of the sheet.
        chart_name: The name / id of the chart.

    Returns:
        A list of series objects in the specified chart, or None if the chart does not exist.
    """
    client = ctx.request_context.lifespan_context.graph_client
    workbook = await client.get_workbook_by_id(workbook_id)
    if not workbook:
        return None
    worksheet = workbook.worksheets.by_workbook_worksheet_id(sheet_name)
    series_data = await worksheet.charts.by_workbook_chart_id(chart_name).series.get()
    writer = JsonSerializationWriter()
    series_data.serialize(writer)
    json_data = json.loads(writer.get_serialized_content())
    return ChartSeriesResponse.model_validate(json_data)
