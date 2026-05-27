import os

import boto3
from botocore.exceptions import BotoCoreError, ClientError
from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.server import TransportSecuritySettings
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import JSONResponse
from starlette.routing import Route

from ingestion.retriever import SlideRetriever

_EXPECTED_TOKEN = os.environ.get("MCP_SHARED_SECRET")


class AuthMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request, call_next):
        if request.url.path in ("/health", "/secrets-check"):
            return await call_next(request)
        # Fail closed if the secret was never injected into the environment.
        if not _EXPECTED_TOKEN:
            return JSONResponse({"error": "server misconfigured"}, status_code=503)
        auth = request.headers.get("authorization", "")
        if not auth.startswith("Bearer ") or auth[7:] != _EXPECTED_TOKEN:
            return JSONResponse({"error": "unauthorized"}, status_code=401)
        return await call_next(request)

mcp = FastMCP(
    "Sales MCP",
    transport_security=TransportSecuritySettings(enable_dns_rebinding_protection=False),
)

_retriever: SlideRetriever | None = None

def _get_retriever() -> SlideRetriever:
    global _retriever
    if _retriever is None:
        _retriever = SlideRetriever()
    return _retriever

@mcp.tool()
def search_decks(query: str, k: int = 5) -> list[dict]:
    """Search Fortune sales decks by semantic similarity. Returns the k most relevant slides."""
    return _get_retriever().search(query, k=k)

@mcp.tool()
def hello(name: str) -> str:
    """Sanity check tool."""
    return f"Hello, {name}!"

async def health(request):
    return JSONResponse({"status": "ok"})

# Temporary to test IAM deployment and permissions. Will delete after confirming the secret check.
async def secrets_check(request):
    secret_name = "fortune-sales-mcp/claude-api-key"
    try:
        client = boto3.client("secretsmanager", region_name="us-east-1")
        client.get_secret_value(SecretId=secret_name)
        return JSONResponse({"status": "ok", "secret": secret_name})
    except ClientError as e:
        return JSONResponse({"status": "error", "detail": str(e)}, status_code=500)
    except BotoCoreError as e:
        return JSONResponse({"status": "error", "detail": str(e)}, status_code=500)

app = mcp.streamable_http_app()

app.routes.append(Route("/health", health))
app.routes.append(Route("/secrets-check", secrets_check))
app.add_middleware(AuthMiddleware)

if __name__ == "__main__":
    mcp.run()