import os
from pathlib import Path
from urllib.parse import urlsplit
from uuid import UUID

from dotenv import dotenv_values
from pydantic import BaseModel, SecretStr

from app.errors import DemoError

ROOT = Path(__file__).resolve().parent.parent


class Settings(BaseModel):
    tenant_id: str = ""
    client_id: str = ""
    ai_endpoint: str = ""
    ai_deployment: str = ""
    ai_auth: str = "entra"
    ai_key: SecretStr = SecretStr("")

    @classmethod
    def from_env(cls) -> "Settings":
        env = {**dotenv_values(ROOT / ".env"), **os.environ}
        return cls(
            tenant_id=env.get("ENTRA_TENANT_ID") or "",
            client_id=env.get("ENTRA_CLIENT_ID") or "",
            ai_endpoint=env.get("AZURE_OPENAI_ENDPOINT") or "",
            ai_deployment=env.get("AZURE_OPENAI_DEPLOYMENT") or "",
            ai_auth=env.get("AZURE_OPENAI_AUTH") or "entra",
            ai_key=SecretStr(env.get("AZURE_OPENAI_API_KEY") or ""),
        )

    def graph_problem(self) -> str | None:
        if not self.tenant_id or not self.client_id:
            return "Set ENTRA_TENANT_ID and ENTRA_CLIENT_ID in .env, then restart."
        try:
            UUID(self.tenant_id)
            UUID(self.client_id)
        except ValueError:
            return "ENTRA_TENANT_ID and ENTRA_CLIENT_ID must be GUIDs."
        return None

    def ai_problem(self) -> str | None:
        if not self.ai_endpoint or "YOUR-RESOURCE" in self.ai_endpoint or not self.ai_deployment:
            return "Set AZURE_OPENAI_ENDPOINT and AZURE_OPENAI_DEPLOYMENT in .env, then restart."
        try:
            url = urlsplit(self.ai_endpoint)
            host = url.hostname or ""
            port = url.port
        except ValueError:
            return "Use a valid Azure HTTPS resource endpoint."
        path = url.path.rstrip("/")
        if path == "/openai/v1/responses":
            return (
                "AZURE_OPENAI_ENDPOINT is an operation URL. Remove the trailing responses "
                "so it ends in /openai/v1/, then restart."
            )
        if path.startswith("/api/projects/"):
            return (
                "This app uses the resource OpenAI endpoint, not the Foundry project endpoint. "
                "Copy the resource's OpenAI endpoint ending in /openai/v1/, then restart."
            )
        if (
            url.scheme != "https"
            or not any(host.endswith(suffix) for suffix in (
                ".openai.azure.com", ".services.ai.azure.com", ".cognitiveservices.azure.com"
            ))
            or url.username
            or url.password
            or url.query
            or url.fragment
            or port not in (None, 443)
            or path not in ("", "/openai/v1")
        ):
            return "Use your Azure HTTPS resource endpoint or its /openai/v1/ URL."
        if self.ai_auth not in ("entra", "key"):
            return "AZURE_OPENAI_AUTH must be entra or key."
        if self.ai_auth == "key" and not self.ai_key.get_secret_value():
            return "Key authentication requires AZURE_OPENAI_API_KEY in .env."
        return None

    def ai_base_url(self) -> str:
        problem = self.ai_problem()
        if problem:
            raise DemoError(problem, 503)
        base = self.ai_endpoint.rstrip("/")
        return base + "/" if base.endswith("/openai/v1") else base + "/openai/v1/"
