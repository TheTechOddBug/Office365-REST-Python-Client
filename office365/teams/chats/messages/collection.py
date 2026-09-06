from office365.delta_collection import DeltaCollection
from office365.teams.chats.messages.body import ChatMessageBody
from office365.teams.chats.messages.message import ChatMessage


class ChatMessageCollection(DeltaCollection[ChatMessage]):
    """Chat message's collection"""

    def __init__(self, context, resource_path=None):
        super().__init__(context, ChatMessage, resource_path)

    def add(self, content: str) -> ChatMessage:
        return super().add(body=ChatMessageBody(content=content))
