import os

from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.server import TransportSecuritySettings
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import JSONResponse
from starlette.routing import Route

_EXPECTED_TOKEN = os.environ.get("MCP_SHARED_SECRET")


class AuthMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request, call_next):
        if request.url.path == "/health":
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

# Stub data so you have something to query
DECKS = [
    {"id": "deck1", "title": "Tech Q1 Pitch", "industry": "tech"},
    {"id": "deck2", "title": "Healthcare Q2", "industry": "healthcare"},
]

@mcp.tool()
def search_historical_decks(industry: str) -> list[dict]:
    """Find historical decks by industry."""
    return [d for d in DECKS if d["industry"] == industry]

@mcp.tool()
def hello(name: str) -> str:
    """Sanity check tool."""
    return f"Hello, {name}!"

# Plain HTTP endpoint for Lightsail's health checks.
async def health(request):
    return JSONResponse({"status": "ok"})

app = mcp.streamable_http_app()

app.routes.append(Route("/health", health))
app.add_middleware(AuthMiddleware)

if __name__ == "__main__":
    mcp.run()