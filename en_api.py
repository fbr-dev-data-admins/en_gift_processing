"""
Engaging Networks API Module

Handles authentication and bulk data export from Engaging Networks API.
"""

import requests
import csv
import io
from typing import List, Optional
import time


def authenticate(token: str, start_date: str, end_date: str) -> List[List[str]]:
    """
    Authenticate with Engaging Networks and retrieve bulk transaction data.
    
    Args:
        token: EN API authentication token
        start_date: Start date in MMDDYYYY format
        end_date: End date in MMDDYYYY format
        
    Returns:
        List of rows (first row is headers, subsequent rows are data)
    """
    url = "https://us.engagingnetworks.app/ea-dataservice/export.service"
    
    querystring = {
        "token": token,
        "startDate": start_date,
        "endDate": end_date
    }
    
    headers = {
        "Accept": "text/html; charset=UTF-8, text/xml; charset=UTF-8, text/csv; charset=UTF-8"
    }
    
    try:
        response = requests.get(url, headers=headers, params=querystring, timeout=300)
        response.raise_for_status()
        
        # Parse CSV response using csv module for proper handling
        csv_content = io.StringIO(response.text)
        reader = csv.reader(csv_content)
        rows = list(reader)
        
        return rows
        
    except requests.exceptions.RequestException as e:
        raise Exception(f"EN API request failed: {e}")


def get_supporter_data(token: str, supporter_id: str) -> Optional[dict]:
    """
    Get detailed supporter data by ID.
    
    Args:
        token: EN API authentication token
        supporter_id: EN Supporter ID
        
    Returns:
        Supporter data dictionary or None
    """
    base_url = "https://e-activist.com/ea-dataservice/supporter.service"
    
    params = {
        'token': token,
        'supporterId': supporter_id
    }
    
    try:
        response = requests.get(base_url, params=params, timeout=60)
        response.raise_for_status()
        
        # Parse response (format depends on EN API version)
        return response.json() if response.text else None
        
    except requests.exceptions.RequestException as e:
        print(f"Error fetching supporter {supporter_id}: {e}")
        return None


def get_transaction_details(token: str, transaction_id: str) -> Optional[dict]:
    """
    Get detailed transaction data by ID.
    
    Args:
        token: EN API authentication token
        transaction_id: EN Transaction ID
        
    Returns:
        Transaction data dictionary or None
    """
    base_url = "https://e-activist.com/ea-dataservice/transaction.service"
    
    params = {
        'token': token,
        'transactionId': transaction_id
    }
    
    try:
        response = requests.get(base_url, params=params, timeout=60)
        response.raise_for_status()
        
        return response.json() if response.text else None
        
    except requests.exceptions.RequestException as e:
        print(f"Error fetching transaction {transaction_id}: {e}")
        return None


class ENBulkAPI:
    """
    Enhanced Engaging Networks Bulk API client with pagination and retry logic.
    """
    
    BASE_URL = "https://us.engagingnetworks.app/ea-dataservice"
    
    def __init__(self, token: str):
        """
        Initialize the EN API client.
        
        Args:
            token: EN API authentication token
        """
        self.token = token
        self.session = requests.Session()
        self.session.headers.update({
            'Content-Type': 'application/json'
        })
    
    def export_transactions(
        self,
        start_date: str,
        end_date: str,
        campaign_types: List[str] = None,
        max_retries: int = 3
    ) -> List[dict]:
        """
        Export transactions with retry logic.
        
        Args:
            start_date: Start date in MMDDYYYY format
            end_date: End date in MMDDYYYY format
            campaign_types: Optional list of campaign types to filter
            max_retries: Maximum number of retry attempts
            
        Returns:
            List of transaction dictionaries
        """
        for attempt in range(max_retries):
            try:
                rows = authenticate(self.token, start_date, end_date)
                
                if not rows:
                    return []
                
                # Convert to list of dicts
                headers = rows[0]
                transactions = []
                
                for row in rows[1:]:
                    if len(row) == len(headers):
                        transaction = dict(zip(headers, row))
                        
                        # Filter by campaign type if specified
                        if campaign_types:
                            if transaction.get('Campaign Type') in campaign_types:
                                transactions.append(transaction)
                        else:
                            transactions.append(transaction)
                
                return transactions
                
            except Exception as e:
                if attempt < max_retries - 1:
                    wait_time = (attempt + 1) * 5  # Exponential backoff
                    print(f"Attempt {attempt + 1} failed, retrying in {wait_time}s: {e}")
                    time.sleep(wait_time)
                else:
                    raise
        
        return []
    
    def get_tribute_data(self, start_date: str, end_date: str) -> List[dict]:
        """
        Get tribute (FIM/PFIM) data for gift reference matching.
        
        Args:
            start_date: Start date in MMDDYYYY format
            end_date: End date in MMDDYYYY format
            
        Returns:
            List of tribute transactions
        """
        return self.export_transactions(
            start_date,
            end_date,
            campaign_types=['FIM', 'PFIM']
        )
