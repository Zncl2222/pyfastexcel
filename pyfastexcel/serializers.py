from pydantic import BaseModel, field_serializer
from typing import Optional

from ._typing import SetPanesSelection
from .utils import Selection


class PanesSerializer(BaseModel):
    selection: Optional[SetPanesSelection | list[Selection] | Selection] = None

    @field_serializer('selection')
    @classmethod
    def serialize_selection(
        cls,
        selection: Optional[SetPanesSelection | list[Selection] | Selection],
    ) -> list[dict[str, int]]:
        if selection is None:
            selection = []
        elif not isinstance(selection, list):
            selection = [selection]

        selection = [item.to_dict() for item in selection if isinstance(item, Selection)]

        return selection
