import asyncio
from typing import Optional
from azure.identity import DeviceCodeCredential
import logging
from mcp_excel_online.core.config import Settings
from mcp_excel_online.core.graph_sdk.service_client import ServiceClient


class GraphClientManager:
    """Singleton wrapper around the authenticated ServiceClient."""
    SCOPES = ['https://graph.microsoft.com/.default']
    _instance: Optional["GraphClientManager"] = None
    _lock = asyncio.Lock()

    def __init__(self):
        # Guard against accidental direct instantiation
        if GraphClientManager._instance is not None:
            raise RuntimeError(
                "Use GraphClientManager.get_instance() instead of constructing directly.")
        self._client: Optional[ServiceClient] = None
        self._credential = DeviceCodeCredential(
            client_id=Settings.CLIENT_ID,
            tenant_id="consumers",
        )

    @classmethod
    async def get_instance(cls) -> "GraphClientManager":
        """Get the singleton, creating and authenticating it on first call."""
        async with cls._lock:
            if cls._instance is None:
                cls._instance = cls()
                await cls._instance._authenticate()
            return cls._instance

    async def _authenticate(self) -> None:
        self._client = ServiceClient(
            credentials=self._credential, scopes=self.SCOPES)
        profile_info = await self._client.me.get()
        logging.info(
            f"Authenticated as: {profile_info.display_name} ({profile_info.user_principal_name})"
        )

    @property
    def client(self) -> ServiceClient:
        if self._client is None:
            raise RuntimeError(
                "GraphClientManager not authenticated yet — call get_instance() first.")
        return self._client
