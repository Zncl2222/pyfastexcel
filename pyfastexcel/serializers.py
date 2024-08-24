from __future__ import annotations

from pydantic import BaseModel, field_serializer
from typing import Optional

from ._typing import CommentTextStructure, SetPanesSelection
from .utils import CommentText, Selection


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
