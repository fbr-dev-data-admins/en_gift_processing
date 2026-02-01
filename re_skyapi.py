"""
Raiser's Edge Sky API Integration Module

Handles authentication and API calls to Blackbaud's RE NXT Sky API
for constituent lookups, gift queries, and solicitor management.
"""

import requests
import json
import os
from datetime import datetime, timedelta
from typing import Optional, Dict, List, Any
from urllib.parse import urlencode


class RESkyAPI:
    """
    Raiser's Edge NXT Sky API Client
    
    Handles OAuth 2.0 authentication flow and API interactions with
    Blackbaud's Sky API for Raiser's Edge NXT.
    """
    
    # API Endpoints
    AUTH_URL = "https://oauth2.sky.blackbaud.com/authorization"
    TOKEN_URL = "https://oauth2.sky.blackbaud.com/token"
    API_BASE_URL = "https://api.sky.blackbaud.com"
    
    def __init__(
        self,
        client_id: str,
        client_secret: str,
        redirect_uri: str,
        subscription_key: str,
        token_file: str = "config/re_tokens.json"
    ):
        """
        Initialize the RE Sky API client.
        
        Args:
            client_id: OAuth application client ID from Blackbaud
            client_secret: OAuth application client secret
            redirect_uri: Registered redirect URI for OAuth flow
            subscription_key: Blackbaud API subscription key (Bb-Api-Subscription-Key)
            token_file: Path to store/load OAuth tokens
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.redirect_uri = redirect_uri
        self.subscription_key = subscription_key
        self.token_file = token_file
        
        self.access_token = None
        self.refresh_token = None
        self.token_expires_at = None
        
        # Load existing tokens if available
        self._load_tokens()
    
    def _load_tokens(self):
        """Load tokens from file if they exist"""
        if os.path.exists(self.token_file):
            try:
                with open(self.token_file, 'r') as f:
                    tokens = json.load(f)
                    self.access_token = tokens.get('access_token')
                    self.refresh_token = tokens.get('refresh_token')
                    expires_at = tokens.get('expires_at')
                    if expires_at:
                        self.token_expires_at = datetime.fromisoformat(expires_at)
            except Exception as e:
                print(f"Error loading tokens: {e}")
    
    def _save_tokens(self):
        """Save tokens to file"""
        os.makedirs(os.path.dirname(self.token_file), exist_ok=True)
        tokens = {
            'access_token': self.access_token,
            'refresh_token': self.refresh_token,
            'expires_at': self.token_expires_at.isoformat() if self.token_expires_at else None
        }
        with open(self.token_file, 'w') as f:
            json.dump(tokens, f)
    
    def get_authorization_url(self, state: str = None) -> str:
        """
        Generate the OAuth authorization URL for user consent.
        
        Args:
            state: Optional state parameter for CSRF protection
            
        Returns:
            Authorization URL to redirect user to
        """
        params = {
            'client_id': self.client_id,
            'response_type': 'code',
            'redirect_uri': self.redirect_uri,
        }
        if state:
            params['state'] = state
        
        return f"{self.AUTH_URL}?{urlencode(params)}"
    
    def exchange_code_for_token(self, authorization_code: str) -> bool:
        """
        Exchange authorization code for access/refresh tokens.
        
        Args:
            authorization_code: Code received from OAuth callback
            
        Returns:
            True if successful, False otherwise
        """
        data = {
            'grant_type': 'authorization_code',
            'code': authorization_code,
            'redirect_uri': self.redirect_uri,
            'client_id': self.client_id,
            'client_secret': self.client_secret,
        }
        
        try:
            response = requests.post(self.TOKEN_URL, data=data)
            response.raise_for_status()
            
            token_data = response.json()
            self.access_token = token_data['access_token']
            self.refresh_token = token_data['refresh_token']
            expires_in = token_data.get('expires_in', 3600)
            self.token_expires_at = datetime.now() + timedelta(seconds=expires_in)
            
            self._save_tokens()
            return True
            
        except Exception as e:
            print(f"Error exchanging code for token: {e}")
            return False
    
    def _refresh_access_token(self) -> bool:
        """
        Refresh the access token using the refresh token.
        
        Returns:
            True if successful, False otherwise
        """
        if not self.refresh_token:
            return False
        
        data = {
            'grant_type': 'refresh_token',
            'refresh_token': self.refresh_token,
            'client_id': self.client_id,
            'client_secret': self.client_secret,
        }
        
        try:
            response = requests.post(self.TOKEN_URL, data=data)
            response.raise_for_status()
            
            token_data = response.json()
            self.access_token = token_data['access_token']
            self.refresh_token = token_data.get('refresh_token', self.refresh_token)
            expires_in = token_data.get('expires_in', 3600)
            self.token_expires_at = datetime.now() + timedelta(seconds=expires_in)
            
            self._save_tokens()
            return True
            
        except Exception as e:
            print(f"Error refreshing token: {e}")
            return False
    
    def is_authenticated(self) -> bool:
        """Check if we have valid authentication"""
        if not self.access_token:
            return False
        
        # Check if token is expired or about to expire (within 5 minutes)
        if self.token_expires_at:
            if datetime.now() >= self.token_expires_at - timedelta(minutes=5):
                # Try to refresh
                return self._refresh_access_token()
        
        return True
    
    def _get_headers(self) -> Dict[str, str]:
        """Get headers for API requests"""
        return {
            'Authorization': f'Bearer {self.access_token}',
            'Bb-Api-Subscription-Key': self.subscription_key,
            'Content-Type': 'application/json',
        }
    
    def _make_request(
        self, 
        method: str, 
        endpoint: str, 
        params: dict = None, 
        data: dict = None
    ) -> Optional[Dict[str, Any]]:
        """
        Make an authenticated API request.
        
        Args:
            method: HTTP method (GET, POST, PATCH, etc.)
            endpoint: API endpoint (without base URL)
            params: Query parameters
            data: Request body data
            
        Returns:
            Response JSON or None if error
        """
        if not self.is_authenticated():
            raise Exception("Not authenticated. Please authenticate first.")
        
        url = f"{self.API_BASE_URL}{endpoint}"
        
        try:
            response = requests.request(
                method=method,
                url=url,
                headers=self._get_headers(),
                params=params,
                json=data
            )
            response.raise_for_status()
            
            if response.text:
                return response.json()
            return {}
            
        except requests.exceptions.HTTPError as e:
            print(f"HTTP Error: {e}")
            print(f"Response: {e.response.text if e.response else 'No response'}")
            raise
        except Exception as e:
            print(f"Request error: {e}")
            raise
    
    # ---------- CONSTITUENT ENDPOINTS ----------
    
    def search_constituents(
        self, 
        search_text: str = None,
        email: str = None,
        constituent_id: str = None
    ) -> List[Dict[str, Any]]:
        """
        Search for constituents.
        
        Args:
            search_text: General search text (name, etc.)
            email: Email address to search
            constituent_id: Specific constituent ID
            
        Returns:
            List of matching constituents
        """
        params = {'limit': 100}
        
        if search_text:
            params['search_text'] = search_text
        
        result = self._make_request('GET', '/constituent/v1/constituents', params=params)
        
        constituents = result.get('value', []) if result else []
        
        # Filter by email if specified
        if email and constituents:
            email_lower = email.lower()
            constituents = [
                c for c in constituents 
                if any(e.get('address', '').lower() == email_lower 
                      for e in c.get('email', {}).get('addresses', []))
            ]
        
        return constituents
    
    def get_constituent(self, constituent_id: str) -> Optional[Dict[str, Any]]:
        """
        Get a specific constituent by ID.
        
        Args:
            constituent_id: RE Constituent ID
            
        Returns:
            Constituent data or None
        """
        return self._make_request('GET', f'/constituent/v1/constituents/{constituent_id}')
    
    def get_constituent_by_lookup_id(self, lookup_id: str) -> Optional[Dict[str, Any]]:
        """
        Get constituent by lookup ID (external system ID).
        
        Args:
            lookup_id: External lookup ID
            
        Returns:
            Constituent data or None
        """
        params = {'lookup_id': lookup_id}
        result = self._make_request('GET', '/constituent/v1/constituents', params=params)
        
        if result and result.get('value'):
            return result['value'][0]
        return None
    
    # ---------- GIFT ENDPOINTS ----------
    
    def get_constituent_gifts(
        self, 
        constituent_id: str,
        gift_date: datetime = None,
        date_range_days: int = 1,
        year: int = None,
        month: int = None,
        days: List[int] = None
    ) -> List[Dict[str, Any]]:
        """
        Get gifts for a constituent, optionally filtered by date.
        
        Args:
            constituent_id: RE Constituent ID
            gift_date: Specific date to filter gifts
            date_range_days: Number of days around gift_date to include
            year: Year to filter (used with month/days)
            month: Month to filter (used with year/days)
            days: List of days of month to check (used with year/month)
            
        Returns:
            List of gifts with 'date' and 'amount' keys
        """
        params = {'limit': 500}
        
        # If year/month/days provided, construct date range
        if year and month and days:
            # Get gifts for the entire month, then filter
            import calendar
            days_in_month = calendar.monthrange(year, month)[1]
            start_date = f"{year}-{month:02d}-01"
            end_date = f"{year}-{month:02d}-{days_in_month:02d}"
            params['date_added'] = f"gte {start_date}"
        elif gift_date:
            start_date = gift_date - timedelta(days=date_range_days)
            end_date = gift_date + timedelta(days=date_range_days)
            params['date_added'] = f"gte {start_date.strftime('%Y-%m-%d')}"
        
        result = self._make_request(
            'GET', 
            f'/gift/v1/constituents/{constituent_id}/gifts',
            params=params
        )
        
        gifts = result.get('value', []) if result else []
        
        # Filter by specific days if year/month/days provided
        if year and month and days and gifts:
            target_dates = [f"{year}-{month:02d}-{d:02d}" for d in days]
            filtered_gifts = []
            for g in gifts:
                gift_date_str = g.get('date', '')[:10]
                if gift_date_str in target_dates:
                    # Exclude soft credits (check gift type)
                    gift_type = g.get('type', '')
                    if 'soft' not in gift_type.lower():
                        filtered_gifts.append({
                            'date': gift_date_str,
                            'amount': g.get('amount', {}).get('value', 0)
                        })
            return filtered_gifts
        
        # Filter by exact date if specified
        if gift_date and gifts:
            target_date = gift_date.date()
            gifts = [
                {'date': g.get('date', '')[:10], 'amount': g.get('amount', {}).get('value', 0)}
                for g in gifts 
                if datetime.fromisoformat(g.get('date', '')[:10]).date() == target_date
            ]
        
        return gifts
    
    def get_gift(self, gift_id: str) -> Optional[Dict[str, Any]]:
        """
        Get a specific gift by ID.
        
        Args:
            gift_id: RE Gift ID
            
        Returns:
            Gift data or None
        """
        return self._make_request('GET', f'/gift/v1/gifts/{gift_id}')
    
    # ---------- SOLICITOR MANAGEMENT ----------
    
    def mark_as_solicitor(self, constituent_id: str) -> bool:
        """
        Mark a constituent as a solicitor by adding the solicitor constituent code.
        
        Args:
            constituent_id: RE Constituent ID to mark as solicitor
            
        Returns:
            True if successful
        """
        # First, get existing constituent codes
        codes_result = self._make_request(
            'GET',
            f'/constituent/v1/constituents/{constituent_id}/constituentcodes'
        )
        
        existing_codes = codes_result.get('value', []) if codes_result else []
        
        # Check if already has solicitor code
        has_solicitor = any(
            c.get('description', '').lower() == 'solicitor' 
            for c in existing_codes
        )
        
        if has_solicitor:
            return True  # Already marked
        
        # Add solicitor constituent code
        # Note: The actual code table entry ID may vary by organization
        # This assumes 'Solicitor' is a valid constituent code in your system
        data = {
            'constituent_id': constituent_id,
            'description': 'Solicitor'
        }
        
        try:
            self._make_request(
                'POST',
                '/constituent/v1/constituentcodes',
                data=data
            )
            return True
        except Exception as e:
            print(f"Error marking as solicitor: {e}")
            return False
    
    def add_constituent_attribute(
        self,
        constituent_id: str,
        category: str,
        value: str,
        date: datetime = None
    ) -> bool:
        """
        Add an attribute to a constituent (for monthly donor fields, etc.).
        
        Args:
            constituent_id: RE Constituent ID
            category: Attribute category name
            value: Attribute value
            date: Optional date for the attribute
            
        Returns:
            True if successful
        """
        data = {
            'constituent_id': constituent_id,
            'category': category,
            'value': value
        }
        
        if date:
            data['date'] = date.strftime('%Y-%m-%d')
        
        try:
            self._make_request(
                'POST',
                '/constituent/v1/constituents/customfields',
                data=data
            )
            return True
        except Exception as e:
            print(f"Error adding attribute: {e}")
            return False
    
    # ---------- P2P MATCHING HELPERS ----------
    
    def find_p2p_solicitor(
        self,
        system_record_id: str = None,
        email: str = None,
        name: str = None
    ) -> Optional[Dict[str, Any]]:
        """
        Find a P2P solicitor by various identifiers.
        
        First tries System Record ID, then email, then name search.
        
        Args:
            system_record_id: RE System Record ID (lookup ID)
            email: Email address
            name: Full name for search
            
        Returns:
            Matched constituent with match type, or None
        """
        result = None
        matched_on = None
        
        # Try System Record ID first
        if system_record_id:
            result = self.get_constituent_by_lookup_id(system_record_id)
            if result:
                matched_on = 'System Record ID'
        
        # Try email if no match
        if not result and email:
            constituents = self.search_constituents(email=email)
            if constituents:
                result = constituents[0]
                matched_on = 'Email'
        
        # Try name search if still no match
        if not result and name:
            constituents = self.search_constituents(search_text=name)
            if constituents:
                result = constituents[0]
                matched_on = 'Name'
        
        if result:
            return {
                'id': result.get('id'),
                'name': result.get('name', ''),
                'email': result.get('email', {}).get('address', ''),
                'matched_on': matched_on,
                'raw': result
            }
        
        return None


# Example usage and testing
if __name__ == "__main__":
    # This is for testing - actual credentials should come from secrets
    import os
    
    api = RESkyAPI(
        client_id=os.getenv('RE_CLIENT_ID', ''),
        client_secret=os.getenv('RE_CLIENT_SECRET', ''),
        redirect_uri=os.getenv('RE_REDIRECT_URI', 'http://localhost:8501/callback'),
        subscription_key=os.getenv('RE_SUBSCRIPTION_KEY', '')
    )
    
    if not api.is_authenticated():
        print("Not authenticated. Get authorization URL:")
        print(api.get_authorization_url())
    else:
        print("Authenticated!")
        # Test search
        results = api.search_constituents(search_text="John Smith")
        print(f"Found {len(results)} constituents")
