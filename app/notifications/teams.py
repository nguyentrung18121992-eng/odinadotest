"""
Proactive Teams message via Bot Framework Connector (Microsoft app id + secret).
Mirrors the Node helper `TeamsOutreach` / `scripts/test-teams-send.js`.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import quote

import httpx
from msal import ConfidentialClientApplication

CONNECTOR_SCOPE = "https://api.botframework.com/.default"


def _load_project_dotenv() -> None:
    try:
        from dotenv import load_dotenv
    except ImportError:
        return
    root = Path(__file__).resolve().parent.parent.parent
    load_dotenv(root / ".env")


@dataclass
class SendResult:
    success: bool
    error: str | None = None


class TeamsSender:
    """Quick test helper — run from backend container with env vars set."""

    def __init__(self) -> None:
        _load_project_dotenv()
        self.app_id = os.environ["MicrosoftAppId"]
        self.app_password = os.environ["MicrosoftAppPassword"]
        tenant = (
            os.environ.get("MicrosoftAppTenantId")
            or os.environ.get("MICROSOFT_APP_TENANT_ID")
            or os.environ.get("AZURE_TENANT_ID")
            or "botframework.com"
        )
        self._authority = f"https://login.microsoftonline.com/{tenant}"
        base = os.environ.get(
            "TEAMS_SERVICE_URL", "https://smba.trafficmanager.net/teams/"
        )
        self._service_url = base.rstrip("/") + "/"

    def _access_token(self) -> str:
        app = ConfidentialClientApplication(
            self.app_id,
            authority=self._authority,
            client_credential=self.app_password,
        )
        result: dict[str, Any] = app.acquire_token_for_client(
            scopes=[CONNECTOR_SCOPE]
        )
        token = result.get("access_token")
        if not token:
            desc = result.get("error_description") or str(result)
            raise RuntimeError(desc)
        return str(token)

    async def send_to_bot(self, chat_id: str, message: str) -> SendResult:
        try:
            token = self._access_token()
            cid = quote(chat_id, safe="")
            url = f"{self._service_url}v3/conversations/{cid}/activities"
            payload = {
                "type": "message",
                "text": message,
                "textFormat": "markdown",
            }
            async with httpx.AsyncClient(timeout=30.0) as client:
                resp = await client.post(
                    url,
                    headers={
                        "Authorization": f"Bearer {token}",
                        "Content-Type": "application/json",
                    },
                    json=payload,
                )
            if resp.status_code >= 300:
                return SendResult(
                    success=False,
                    error=f"{resp.status_code} {resp.text}",
                )
            return SendResult(success=True)
        except Exception as e:
            return SendResult(success=False, error=str(e))
