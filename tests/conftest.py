from unittest.mock import patch
import requests
import pytest


@pytest.fixture(autouse=True, scope="session")
def mock_moex():
    with patch.object(requests, "get") as mock_requests_get:
        def side_effect(url: str):
            response = None
            if url.startswith("https://iss.moex.com/iss/engines/stock/markets/bonds/boardgroups/58/securities.json"):
                raise requests.exceptions.RequestException()

            return response

        mock_requests_get.side_effect = side_effect
        yield