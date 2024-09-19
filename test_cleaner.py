# test_cleaner.py
import os
import sys
import asyncio
import aiohttp
import pytest
from unittest.mock import patch, MagicMock, mock_open
from cleaner import CalendarCleaner

@pytest.fixture
def cleaner(monkeypatch):
    monkeypatch.setenv('CLIENT_ID', 'test_client_id')
    monkeypatch.setenv('TENANT_ID', 'test_tenant_id')
    
    with patch('msal.PublicClientApplication') as MockClientApp:
        mock_app = MockClientApp.return_value
        mock_app.get_accounts.return_value = [{'id': 'test_account'}]
        mock_app.acquire_token_silent.return_value = {'access_token': 'test_token'}
        mock_app.initiate_device_flow.return_value = {'user_code': 'test_code', 'message': 'test_message'}
        mock_app.acquire_token_by_device_flow.return_value = {'access_token': 'test_token'}
        return CalendarCleaner('test_client_id', 'test_tenant_id')

def test_load_cache_exists(cleaner, monkeypatch):
    monkeypatch.setattr(os.path, 'exists', lambda x: True)
    with patch('builtins.open', mock_open(read_data='{}')) as mock_file:
        cache = cleaner.load_cache()
        assert cache is not None
        mock_file.assert_called_once_with('token_cache.bin', 'r')

def test_load_cache_not_exists(cleaner, monkeypatch):
    monkeypatch.setattr(os.path, 'exists', lambda x: False)
    cache = cleaner.load_cache()
    assert cache is not None

def test_save_cache(cleaner):
    with patch('builtins.open', mock_open()) as mock_file:
        cleaner.cache.has_state_changed = True
        cleaner.save_cache()
        mock_file.assert_called_once_with('token_cache.bin', 'w')

def test_clean_cache(cleaner, monkeypatch):
    monkeypatch.setattr(os.path, 'exists', lambda x: True)
    with patch('os.remove') as mock_remove:
        cleaner.clean_cache()
        mock_remove.assert_called_once_with('token_cache.bin')

def test_acquire_token_silent(cleaner):
    token = cleaner.acquire_token()
    assert token == 'test_token'

def test_acquire_token_device_flow(cleaner, monkeypatch):
    with patch.object(cleaner.app, 'get_accounts', return_value=[]), \
         patch('sys.stdout', new_callable=MagicMock) as mock_stdout:
        token = cleaner.acquire_token()
        assert token == 'test_token'
        mock_stdout.flush.assert_called_once()

def test_fetch_events(cleaner):
    with patch('requests.get') as mock_get:
        mock_response = MagicMock()
        mock_response.json.return_value = {'value': [], '@odata.nextLink': None}
        mock_get.return_value = mock_response

        events = cleaner.fetch_events('2023-01-01T00:00:00Z', '2023-01-02T00:00:00Z')
        assert events == []
        mock_get.assert_called()
