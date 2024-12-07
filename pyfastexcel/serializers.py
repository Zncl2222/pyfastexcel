from __future__ import annotations

from typing import Any, Optional

from pydantic import BaseModel, field_serializer, model_serializer

from ._typing import CommentTextStructure, SetPanesSelection
from .utils import CommentText, Selection, _validate_cell_reference


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


class CommentSerializer(BaseModel):
    text: CommentTextStructure | CommentText | list[CommentText]

    @field_serializer('text')
    @classmethod
    def serialize_text(
        cls,
        text: CommentTextStructure | CommentText | list[CommentText],
    ) -> list[dict[str, str]]:
        text = (
            [text]
            if isinstance(text, (str, CommentText))
            else text
            if isinstance(text, list)
            else [text]
        )
        if all(isinstance(item, (dict, str)) for item in text):
            for idx, item in enumerate(text):
                if isinstance(item, str):
                    text[idx] = {'text': item}
                else:
                    if 'text' not in item:
                        raise ValueError('Comment text should contain the key "text".')
                    text[idx] = {
                        k[0].upper() + k[1:] if k != 'text' else k: v for k, v in item.items()
                    }
        elif all(isinstance(item, CommentText) for item in text):
            text = [t.to_dict() for t in text]
        return text


class DataValidationSerializer(BaseModel):
    set_range: Optional[list[int | float]]
    input_msg: Optional[list[str]]
    drop_list: Optional[list[str | int | float] | str]
    error_msg: Optional[list[str]]

    @model_serializer(mode='plain')
    def model_serialize(self) -> dict[str, Any]:
        drop_list_key = 'drop_list'
        if isinstance(self.drop_list, str):
            if ':' not in self.drop_list:
                raise ValueError(
                    'Invalid drop list. Sequential Reference'
                    'Drop list should be in the format "A1:B2".',
                )
            drop_list_split = self.drop_list.split(':')
            _validate_cell_reference(drop_list_split[0])
            _validate_cell_reference(drop_list_split[1])
            drop_list_key = 'sqref_drop_list'
        elif self.drop_list is not None:
            self.drop_list = [str(x) for x in self.drop_list]

        dv = {}
        if self.set_range is not None:
            if not isinstance(self.set_range, list) or len(self.set_range) != 2:
                raise ValueError('Set range should be a list of two elements. Like [1, 10].')
            dv['set_range'] = self.set_range
        if self.input_msg is not None:
            if not isinstance(self.input_msg, list) or len(self.input_msg) != 2:
                raise ValueError(
                    'Input message should be a list of two elements. Like ["Title", "Body"].',
                )
            dv['input_title'] = self.input_msg[0]
            dv['input_body'] = self.input_msg[1]
        if self.drop_list is not None:
            dv[drop_list_key] = self.drop_list
        if self.error_msg is not None:
            dv['error_title'] = self.error_msg[0]
            dv['error_body'] = self.error_msg[1]

        return dv
