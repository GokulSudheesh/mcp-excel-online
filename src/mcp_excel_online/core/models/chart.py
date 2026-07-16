from pydantic import BaseModel, Field
from typing import List, Literal, Optional

from mcp_excel_online.core.models.tools import GenericValueResponse


class Chart(BaseModel):
    id: str
    odata_context: Optional[str] = Field(None, alias="@odata.context")
    odata_id: Optional[str] = Field(None, alias="@odata.id")
    height: int
    left: float
    name: str
    top: float
    width: int

    class Config:
        populate_by_name = True


class ChartList(BaseModel):
    odata_context: Optional[str] = Field(None, alias="@odata.context")
    value: list[Chart]

    class Config:
        populate_by_name = True


class UpdateChartRequest(BaseModel):
    name: Optional[str] = Field(
        default=None, description="Represents the name of a chart object.")
    height: Optional[float] = Field(
        default=None, description="Represents the height, in points, of the chart object.")
    width: Optional[float] = Field(
        default=None, description="Represents the width, in points, of the chart object.")
    left: Optional[float] = Field(
        default=None, description="The distance, in points, from the left side of the chart to the worksheet origin.")
    top: Optional[float] = Field(
        default=None, description="Represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).")


Chart_Type = Literal["ColumnClustered", "ColumnStacked", "ColumnStacked100", "BarClustered", "BarStacked", "BarStacked100",
                     "LineStacked", "LineStacked100", "LineMarkers", "LineMarkersStacked", "LineMarkersStacked100", "Pie", "PieOfPie"]

Series_By = Literal["Auto", "Columns", "Rows"]


class CreateChartRequest(BaseModel):
    type: Chart_Type = Field(
        ..., description="The type of chart to create.")
    source_data: str = Field(
        ..., description="The range of cells that contain the data for the chart, in A1 notation (e.g., 'A1:C10').")
    series_by: Optional[Series_By] = Field(
        default=None, description="Specifies the way columns or rows are used as data series on the chart.")


Fitting_Mode = Literal["Fit", "FitAndCenter", "Fill"]


class ChartImageResponse(BaseModel):
    base64_image: str = Field(
        ..., description="The base64-encoded image data of the chart.")


class ChartSeriesItem(BaseModel):
    name: str = Field(..., description="The name of the series.")


class ChartSeriesResponse(GenericValueResponse[List[ChartSeriesItem]]):
    ...


class CreateChartSeriesRequest(BaseModel):
    name: str = Field(..., description="The name of the new series to create.")
