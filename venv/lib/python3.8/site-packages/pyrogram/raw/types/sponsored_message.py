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


class SponsoredMessage(TLObject):  # type: ignore
    """This object is a constructor of the base type :obj:`~pyrogram.raw.base.SponsoredMessage`.

    Details:
        - Layer: ``133``
        - ID: ``0x2a3c381f``

    Parameters:
        random_id: ``bytes``
        from_id: :obj:`Peer <pyrogram.raw.base.Peer>`
        message: ``str``
        start_param (optional): ``str``
        entities (optional): List of :obj:`MessageEntity <pyrogram.raw.base.MessageEntity>`
    """

    __slots__: List[str] = ["random_id", "from_id", "message", "start_param", "entities"]

    ID = 0x2a3c381f
    QUALNAME = "types.SponsoredMessage"

    def __init__(self, *, random_id: bytes, from_id: "raw.base.Peer", message: str, start_param: Union[None, str] = None, entities: Union[None, List["raw.base.MessageEntity"]] = None) -> None:
        self.random_id = random_id  # bytes
        self.from_id = from_id  # Peer
        self.message = message  # string
        self.start_param = start_param  # flags.0?string
        self.entities = entities  # flags.1?Vector<MessageEntity>

    @staticmethod
    def read(data: BytesIO, *args: Any) -> "SponsoredMessage":
        flags = Int.read(data)
        
        random_id = Bytes.read(data)
        
        from_id = TLObject.read(data)
        
        start_param = String.read(data) if flags & (1 << 0) else None
        message = String.read(data)
        
        entities = TLObject.read(data) if flags & (1 << 1) else []
        
        return SponsoredMessage(random_id=random_id, from_id=from_id, message=message, start_param=start_param, entities=entities)

    def write(self) -> bytes:
        data = BytesIO()
        data.write(Int(self.ID, False))

        flags = 0
        flags |= (1 << 0) if self.start_param is not None else 0
        flags |= (1 << 1) if self.entities is not None else 0
        data.write(Int(flags))
        
        data.write(Bytes(self.random_id))
        
        data.write(self.from_id.write())
        
        if self.start_param is not None:
            data.write(String(self.start_param))
        
        data.write(String(self.message))
        
        if self.entities is not None:
            data.write(Vector(self.entities))
        
        return data.getvalue()
