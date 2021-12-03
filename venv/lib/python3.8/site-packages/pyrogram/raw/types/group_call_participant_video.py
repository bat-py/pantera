#  Pyrogram - Telegram MTProto API Client Library for Python
#  Copyright (C) 2017-2021 Dan <https://github.com/delivrance>
#
#  This file is part of Pyrogram.
#
#  Pyrogram is free software: you can redistribute it and/or modify
#  it under the terms of the GNU Lesser General Public License as published
#  by the Free Software Foundation, either version 3 of the License, or
#  (at your option) any later version.
#
#  Pyrogram is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU Lesser General Public License for more details.
#
#  You should have received a copy of the GNU Lesser General Public License
#  along with Pyrogram.  If not, see <http://www.gnu.org/licenses/>.

from io import BytesIO

from pyrogram.raw.core.primitives import Int, Long, Int128, Int256, Bool, Bytes, String, Double, Vector
from pyrogram.raw.core import TLObject
from pyrogram import raw
from typing import List, Union, Any

# # # # # # # # # # # # # # # # # # # # # # # #
#               !!! WARNING !!!               #
#          This is a generated file!          #
# All changes made in this file will be lost! #
# # # # # # # # # # # # # # # # # # # # # # # #


class GroupCallParticipantVideo(TLObject):  # type: ignore
    """This object is a constructor of the base type :obj:`~pyrogram.raw.base.GroupCallParticipantVideo`.

    Details:
        - Layer: ``133``
        - ID: ``0x67753ac8``

    Parameters:
        endpoint: ``str``
        source_groups: List of :obj:`GroupCallParticipantVideoSourceGroup <pyrogram.raw.base.GroupCallParticipantVideoSourceGroup>`
        paused (optional): ``bool``
        audio_source (optional): ``int`` ``32-bit``
    """

    __slots__: List[str] = ["endpoint", "source_groups", "paused", "audio_source"]

    ID = 0x67753ac8
    QUALNAME = "types.GroupCallParticipantVideo"

    def __init__(self, *, endpoint: str, source_groups: List["raw.base.GroupCallParticipantVideoSourceGroup"], paused: Union[None, bool] = None, audio_source: Union[None, int] = None) -> None:
        self.endpoint = endpoint  # string
        self.source_groups = source_groups  # Vector<GroupCallParticipantVideoSourceGroup>
        self.paused = paused  # flags.0?true
        self.audio_source = audio_source  # flags.1?int

    @staticmethod
    def read(data: BytesIO, *args: Any) -> "GroupCallParticipantVideo":
        flags = Int.read(data)
        
        paused = True if flags & (1 << 0) else False
        endpoint = String.read(data)
        
        source_groups = TLObject.read(data)
        
        audio_source = Int.read(data) if flags & (1 << 1) else None
        return GroupCallParticipantVideo(endpoint=endpoint, source_groups=source_groups, paused=paused, audio_source=audio_source)

    def write(self) -> bytes:
        data = BytesIO()
        data.write(Int(self.ID, False))

        flags = 0
        flags |= (1 << 0) if self.paused else 0
        flags |= (1 << 1) if self.audio_source is not None else 0
        data.write(Int(flags))
        
        data.write(String(self.endpoint))
        
        data.write(Vector(self.source_groups))
        
        if self.audio_source is not None:
            data.write(Int(self.audio_source))
        
        return data.getvalue()
