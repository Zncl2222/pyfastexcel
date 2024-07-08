from typing import Literal, TypedDict, Optional


class CommentTextDict(TypedDict, total=False):
    text: str
    size: Optional[int]
    name: Optional[str]
    bold: Optional[bool]
    italic: Optional[bool]
    underline: Optional[Literal['single', 'double']]
    strike: Optional[bool]
    vertAlign: Optional[str]
    color: Optional[str]


CommentTextStructure = str | list[str] | CommentTextDict | list[CommentTextDict]
