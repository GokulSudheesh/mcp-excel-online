import asyncio
from typing import Optional, Union
from azure.identity import DeviceCodeCredential
from azure.identity.aio import ClientSecretCredential
from azure.core.credentials import TokenCredential
from azure.core.credentials_async import AsyncTokenCredential
import logging
from mcp_excel_online.core.args import args
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
        self._credential: Optional[Union[TokenCredential,
                                         AsyncTokenCredential]] = self.get_credential()

    @classmethod
    async def get_instance(cls) -> "GraphClientManager":
        """Get the singleton, creating and authenticating it on first call."""
        async with cls._lock:
            if cls._instance is None:
                cls._instance = cls()
                await cls._instance._authenticate()
            return cls._instance

    def get_credential(self) -> Optional[Union[TokenCredential, AsyncTokenCredential]]:
        if args.graph_permission == "application":
            logging.info(
                "Using application credential (ClientSecretCredential)")
            return ClientSecretCredential(
                tenant_id=Settings.TENANT_ID,
                client_id=Settings.CLIENT_ID,
                client_secret=Settings.CLIENT_SECRET,
            )
        elif args.graph_permission == "delegated":
            logging.info("Using delegated credential (DeviceCodeCredential)")
            return DeviceCodeCredential(
                client_id=Settings.CLIENT_ID,
                tenant_id="consumers",
            )
        else:
            raise ValueError(
                f"Invalid graph_permission: {args.graph_permission}")

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
