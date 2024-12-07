from __future__ import annotations

from typing import List, Literal, Optional, Protocol, TypeVar, Union

from typing_extensions import TypedDict


class CommentTextDict(TypedDict, total=False):
    text: str
    size: Optional[int]
    name: Optional[str]
    bold: Optional[bool]
    italic: Optional[bool]
    underline: Optional[Literal['single', 'double']]
    strike: Optional[bool]
    vert_align: Optional[str]
    color: Optional[str]


class SelectionDict(TypedDict, total=False):
    sq_ref: str
    active_cell: str
    pane: str


class Writable(Protocol):
    def write(self, content: str) -> None: ...


Self = TypeVar('Self')
CommentTextStructure = Union[str, List[str], CommentTextDict, List[CommentTextDict]]
SetPanesSelection = List[SelectionDict]
