import logging
import os
import time

import msal
import requests

from ingestion.models import Tags

logger = logging.getLogger(__name__)

_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_SCOPE = ["https://graph.microsoft.com/.default"]


class GraphClient:
    def __init__(
        self,
        tenant_id: str | None = None,
        client_id: str | None = None,
        client_secret: str | None = None,
    ):
        self._tenant_id = tenant_id or os.environ["GRAPH_TENANT_ID"]
        self._client_id = client_id or os.environ["GRAPH_CLIENT_ID"]
        self._client_secret = client_secret or os.environ["GRAPH_CLIENT_SECRET"]

        self._app = msal.ConfidentialClientApplication(
            self._client_id,
            authority=f"https://login.microsoftonline.com/{self._tenant_id}",
            client_credential=self._client_secret,
        )
        self._session = requests.Session()

    def _token(self) -> str:
        # check token cache first; only calls AAD if expired
        result = self._app.acquire_token_silent(_SCOPE, account=None)
        if not result:
            result = self._app.acquire_token_for_client(scopes=_SCOPE)
        if "access_token" not in result:
            raise RuntimeError(f"MSAL auth failed: {result.get('error_description')}")
        return result["access_token"]

    def _get(self, url: str, params: dict | None = None, max_retries: int = 5) -> dict:
        headers = {"Authorization": f"Bearer {self._token()}"}
        delay = 1.0
        for attempt in range(max_retries):
            resp = self._session.get(url, headers=headers, params=params)
            if resp.status_code == 429:
                wait = float(resp.headers.get("Retry-After", delay))
                logger.warning(
                    "rate limited by Graph API — waiting %.1fs (attempt %d)", wait, attempt + 1
                )
                time.sleep(wait)
                delay = min(delay * 2, 60)
                continue
            resp.raise_for_status()
            return resp.json()
        raise RuntimeError(f"Graph API request failed after {max_retries} retries: {url}")

    def _paginate(self, url: str, params: dict | None = None):
        while url:
            page = self._get(url, params=params)
            yield from page.get("value", [])
            url = page.get("@odata.nextLink")
            params = None  # nextLink already encodes all params

    def _folder_children(self, item_id: str) -> list[dict]:
        site_id = os.environ["SHAREPOINT_SITE_ID"]
        url = f"{_GRAPH_BASE}/sites/{site_id}/drive/items/{item_id}/children"
        return list(self._paginate(url))

    def _walk_folder(self, item_id: str, industry: str, sub_industry: str, results: list) -> None:
        for item in self._folder_children(item_id):
            if "folder" in item:
                # second-level folder = sub-industry; go one level deeper only
                if not sub_industry:
                    self._walk_folder(item["id"], industry, item["name"], results)
            elif "file" in item and item["name"].lower().endswith(".pptx"):
                results.append({**item, "_industry": industry, "_sub_industry": sub_industry})

    def download_deck(self, item_id: str) -> bytes:
        site_id = os.environ["SHAREPOINT_SITE_ID"]
        url = f"{_GRAPH_BASE}/sites/{site_id}/drive/items/{item_id}/content"
        headers = {"Authorization": f"Bearer {self._token()}"}
        resp = self._session.get(url, headers=headers, allow_redirects=True)
        resp.raise_for_status()
        return resp.content

    def extract_tags(self, deck_item: dict) -> Tags:
        sub_industry = deck_item.get("_sub_industry", "")
        return Tags(
            industry=deck_item["_industry"],
            sub_industry=sub_industry,
            deck_type="template" if sub_industry == "General Deck Shells" else "",
        )

    def list_decks(self) -> list[dict]:
        site_id = os.environ["SHAREPOINT_SITE_ID"]
        root_url = f"{_GRAPH_BASE}/sites/{site_id}/drive/root/children"
        results = []
        for item in self._paginate(root_url):
            if "folder" in item and "GTM Current" in item["name"]:
                # GTM Current subfolders = industry verticals (source of truth)
                for industry_folder in self._folder_children(item["id"]):
                    if "folder" in industry_folder:
                        self._walk_folder(
                            industry_folder["id"],
                            industry=industry_folder["name"],
                            sub_industry="",
                            results=results,
                        )
        return results
