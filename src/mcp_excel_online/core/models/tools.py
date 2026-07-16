from typing import Generic, TypeVar

from pydantic import BaseModel, Field

T = TypeVar("T")


class GenericValueResponse(BaseModel, Generic[T]):
    value: T = Field(
        description="The value of the response, which can be of any generic type."
    )


WORKBOOK_ID_FIELD = Field(description="The ID of the workbook.")
SHEET_NAME_FIELD = Field(description="The name of the sheet.")
