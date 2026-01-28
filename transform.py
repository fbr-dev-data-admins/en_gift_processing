"""
Gift Transformation Logic Module

Handles all data transformation rules for converting Engaging Networks
export data to Raiser's Edge import format.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import re
from typing import Tuple, List, Dict, Optional


class GiftTransformer:
    """Main transformation class for EN to RE gift data conversion"""
    
    # Campaign types that should continue processing
    VALID_SUCCESS_TYPES = ['FCS', 'FBS', 'FCR', 'FBR', 'PFCS', 'PFBS', 'PFCR', 'PFBR']
    VALID_PENDING_TYPES = ['FBS', 'FBR', 'PFBS', 'PFBR']
    TRIBUTE_TYPES = ['FIM', 'PFIM']
    P2P_TYPES = ['PFTC', 'PFCS', 'PFCR', 'PFBS', 'PFBR']
    RECURRING_TYPES = ['FCR', 'FBR', 'PFCR', 'PFBR']
    
    # Fundraising page ID mappings
    FUNDRAISING_PAGE_MAPPING = {
        '1064': 'Denver - Food Bank of the Rockies Virtual Food Drives',
        '1068': 'Western Slope - Food Bank of the Rockies Virtual Food Drives',
        '1067': 'Wyoming - Food Bank of the Rockies Virtual Food Drives'
    }
    
    def __init__(self):
        self.exceptions = []
        self.p2p_pending = []
    
    def transform(
        self,
        df: pd.DataFrame,
        mapping_config: dict,
        p2p_config: dict,
        tribute_df: Optional[pd.DataFrame] = None,
        re_api=None
    ) -> Tuple[pd.DataFrame, pd.DataFrame, List[dict]]:
        """
        Main transformation method
        
        Returns:
            Tuple of (processed_df, exceptions_df, p2p_pending_list)
        """
        self.exceptions = []
        self.p2p_pending = []
        
        # Create a copy to avoid modifying original
        df = df.copy()
        
        # Initialize output columns
        output_df = pd.DataFrame()
        
        # Filter by campaign type and status
        df = self._filter_by_status(df)
        
        if len(df) == 0:
            return pd.DataFrame(), pd.DataFrame(self.exceptions), self.p2p_pending
        
        # Process each transformation rule
        output_df['Branch'] = df.apply(self._get_branch, axis=1)
        output_df['EN Donation Form Name'] = df['Campaign ID'].fillna('') if 'Campaign ID' in df.columns else ''
        
        # Apply mapping config for Campaign, Appeal ID, Fund ID, Package
        mapping_results = df.apply(
            lambda row: self._apply_mapping(row, mapping_config), 
            axis=1, 
            result_type='expand'
        )
        output_df['Campaign'] = mapping_results[0]
        output_df['Appeal ID'] = mapping_results[1]
        output_df['Fund ID'] = mapping_results[2]
        output_df['Package'] = mapping_results[3]
        
        # Gift Subtype
        output_df['Gift Subtype'] = df.apply(self._get_gift_subtype, axis=1)
        
        # Donation Type
        output_df['Donation Type'] = df['Campaign Type'].apply(
            lambda x: 'Recurring' if x in self.RECURRING_TYPES else ''
        )
        
        # Monthly Donor fields
        monthly_fields = df.apply(
            lambda row: self._get_monthly_donor_fields(row, re_api), 
            axis=1, 
            result_type='expand'
        )
        monthly_cols = [
            'Monthly Donor Status Description', 'Monthly Donor Status Date',
            'Monthly Donor Anniversary Description', 'Monthly Donor Anniversary Date',
            'Monthly Donor Annual Statement Type', 'Monthly Donor Channel',
            'Monthly Donor Payment Method', 'Monthly Donor Region', 'Gifts Last Month'
        ]
        for i, col in enumerate(monthly_cols):
            output_df[col] = monthly_fields[i] if i < len(monthly_fields.columns) else ''
        
        # P2P fields
        p2p_fields = df.apply(
            lambda row: self._get_p2p_fields(row, p2p_config), 
            axis=1, 
            result_type='expand'
        )
        output_df['EN Fundraising Page ID'] = p2p_fields[0]
        output_df['EN Fundraising Page Name'] = p2p_fields[1]
        output_df['EN Campaign ID'] = p2p_fields[2]
        output_df['EN Campaign Name'] = p2p_fields[3]
        output_df['Gift Solicitor'] = p2p_fields[4]
        
        # Direct field mappings
        output_df['EN Transaction ID'] = df['EN Transaction ID'].fillna('') if 'EN Transaction ID' in df.columns else ''
        output_df['Gift Amount'] = df['Campaign Data 4'].fillna('') if 'Campaign Data 4' in df.columns else ''
        output_df['Gift Date'] = df['Campaign Date'].fillna('') if 'Campaign Date' in df.columns else ''
        output_df['GL Post Date'] = df['Campaign Date'].fillna('') if 'Campaign Date' in df.columns else ''
        output_df['Stripe Transaction ID'] = df['Campaign Data 2'].fillna('') if 'Campaign Data 2' in df.columns else ''
        output_df['Engaging Networks ID'] = df['Supporter ID'].fillna('') if 'Supporter ID' in df.columns else ''
        
        # Handle columns that might have different names
        if 'Company/Org Name' in df.columns:
            output_df['Org Name'] = df['Company/Org Name'].fillna('')
        elif 'Company Name' in df.columns:
            output_df['Org Name'] = df['Company Name'].fillna('')
        else:
            output_df['Org Name'] = ''
        
        if 'Raisers Edge Constituent ID' in df.columns:
            output_df['Constituent ID'] = df['Raisers Edge Constituent ID'].fillna('')
        elif 'RE Constituent ID' in df.columns:
            output_df['Constituent ID'] = df['RE Constituent ID'].fillna('')
        else:
            output_df['Constituent ID'] = ''
        
        output_df['First Name'] = df['First Name'].fillna('') if 'First Name' in df.columns else ''
        output_df['Nickname'] = df['First Name'].fillna('') if 'First Name' in df.columns else ''
        output_df['Middle Name'] = df['Middle Name'].fillna('') if 'Middle Name' in df.columns else ''
        output_df['Last Name'] = df['Last Name'].fillna('') if 'Last Name' in df.columns else ''
        
        # Spouse name parsing
        spouse_fields = df.apply(
            lambda row: self._parse_spouse_name(row), 
            axis=1, 
            result_type='expand'
        )
        output_df['Spouse First Name'] = spouse_fields[0]
        output_df['Spouse Middle Name'] = spouse_fields[1]
        output_df['Spouse Last Name'] = spouse_fields[2]
        
        # Addressee/Salutation formulas (as strings for Excel)
        output_df['Addressee'] = df.apply(
            lambda row: self._get_addressee_formula(row, output_df.loc[row.name] if row.name in output_df.index else {}), 
            axis=1
        )
        output_df['Spouse Addressee'] = output_df.apply(
            lambda row: row['Addressee'] if row['Spouse Last Name'] else '', 
            axis=1
        )
        output_df['Salutation'] = df.apply(
            lambda row: self._get_salutation_formula(row, output_df.loc[row.name] if row.name in output_df.index else {}), 
            axis=1
        )
        output_df['Spouse Salutation'] = output_df.apply(
            lambda row: row['Salutation'] if row['Spouse Last Name'] else '', 
            axis=1
        )
        
        # Address fields
        if 'Address 1' in df.columns:
            addr1 = df['Address 1'].fillna('')
        elif 'Address' in df.columns:
            addr1 = df['Address'].fillna('')
        else:
            addr1 = pd.Series([''] * len(df), index=df.index)
        
        if 'Address 2' in df.columns:
            addr2 = df['Address 2'].fillna('')
        else:
            addr2 = pd.Series([''] * len(df), index=df.index)
        
        output_df['Address'] = (addr1.astype(str) + ' ' + addr2.astype(str)).str.strip()
        output_df['City'] = df['City'].fillna('') if 'City' in df.columns else ''
        output_df['State'] = df['State'].fillna('') if 'State' in df.columns else ''
        
        if 'ZIP' in df.columns:
            output_df['ZIP'] = df['ZIP'].fillna('')
        elif 'Postal Code' in df.columns:
            output_df['ZIP'] = df['Postal Code'].fillna('')
        else:
            output_df['ZIP'] = ''
        
        if 'Country' in df.columns:
            output_df['Country'] = df['Country'].apply(
                lambda x: 'United States' if str(x).upper() == 'US' else x
            ).fillna('')
        else:
            output_df['Country'] = ''
        
        output_df['E-mail'] = df['Email'].fillna('') if 'Email' in df.columns else ''
        
        # Mobile phone - remove (+1)
        if 'Mobile Phone' in df.columns:
            output_df['Cell'] = df['Mobile Phone'].fillna('').astype(str).str.replace(r'^\(\+1\)', '', regex=True).str.strip()
        elif 'Phone' in df.columns:
            output_df['Cell'] = df['Phone'].fillna('').astype(str).str.replace(r'^\(\+1\)', '', regex=True).str.strip()
        else:
            output_df['Cell'] = ''
        
        # Gift Reference (from tribute matching)
        if tribute_df is not None and len(tribute_df) > 0:
            output_df['Gift Reference'] = df.apply(
                lambda row: self._get_gift_reference(row, tribute_df), 
                axis=1
            )
        else:
            output_df['Gift Reference'] = ''
        
        # Requests no email (from QCB campaign type)
        output_df['Requests no email?'] = df.apply(self._get_no_email, axis=1)
        
        # Key Indicator - all "I" initially
        output_df['Key Indicator'] = 'I'
        
        # Mark records with Org Name as "O"
        output_df.loc[output_df['Org Name'].notna() & (output_df['Org Name'] != ''), 'Key Indicator'] = 'O'
        
        # Credit Type (usually empty, set based on business rules)
        output_df['Credit Type'] = ''
        
        # Create exceptions dataframe
        exceptions_df = pd.DataFrame(self.exceptions) if self.exceptions else pd.DataFrame()
        
        return output_df, exceptions_df, self.p2p_pending
    
    def _filter_by_status(self, df: pd.DataFrame) -> pd.DataFrame:
        """Filter records by campaign type and status"""
        if 'Campaign Type' not in df.columns or 'Campaign Status' not in df.columns:
            return df
        
        valid_mask = pd.Series([False] * len(df), index=df.index)
        
        # Success status for FCS, FBS, FCR, FBR, PFCS, PFBS, PFCR, PFBR
        success_mask = (
            df['Campaign Type'].isin(self.VALID_SUCCESS_TYPES) & 
            (df['Campaign Status'] == 'success')
        )
        valid_mask = valid_mask | success_mask
        
        # Pending status for FBS, FBR, PFBS, PFBR (bank transactions)
        pending_mask = (
            df['Campaign Type'].isin(self.VALID_PENDING_TYPES) & 
            (df['Campaign Status'] == 'pending')
        )
        valid_mask = valid_mask | pending_mask
        
        # Track exceptions (reject or change status)
        exception_mask = (
            df['Campaign Type'].isin(self.VALID_SUCCESS_TYPES) & 
            ((df['Campaign Status'] == 'reject') | df['Campaign Status'].str.contains('change', case=False, na=False))
        )
        
        for idx in df[exception_mask].index:
            self.exceptions.append({
                'EN Transaction ID': df.loc[idx, 'EN Transaction ID'],
                'Campaign Type': df.loc[idx, 'Campaign Type'],
                'Campaign Status': df.loc[idx, 'Campaign Status'],
                'Campaign ID': df.loc[idx, 'Campaign ID'] if 'Campaign ID' in df.columns else '',
                'Reason': 'Rejected or Changed Status'
            })
        
        # Filter to valid records first
        filtered_df = df[valid_mask].copy()
        
        # Now filter out form names starting with D.8., Y.8., or S.8. and add to exceptions
        if 'Campaign ID' in filtered_df.columns:
            excluded_form_mask = filtered_df['Campaign ID'].str.startswith(('D.8.', 'Y.8.', 'S.8.'), na=False)
            
            for idx in filtered_df[excluded_form_mask].index:
                self.exceptions.append({
                    'EN Transaction ID': filtered_df.loc[idx, 'EN Transaction ID'] if 'EN Transaction ID' in filtered_df.columns else '',
                    'Campaign Type': filtered_df.loc[idx, 'Campaign Type'],
                    'Campaign Status': filtered_df.loc[idx, 'Campaign Status'],
                    'Campaign ID': filtered_df.loc[idx, 'Campaign ID'],
                    'Reason': 'Excluded Form Name (D.8./Y.8./S.8.)'
                })
            
            # Remove excluded forms from the filtered dataframe
            filtered_df = filtered_df[~excluded_form_mask].copy()
        
        return filtered_df
    
    def _get_branch(self, row: pd.Series) -> str:
        """Determine branch from Campaign ID prefix"""
        campaign_id = str(row.get('Campaign ID', ''))
        
        if campaign_id.startswith('D.'):
            return 'Main'
        elif campaign_id.startswith('S.'):
            return 'WSlope'
        elif campaign_id.startswith('Y.'):
            return 'Wyoming'
        return ''
    
    def _apply_mapping(self, row: pd.Series, mapping_config: dict) -> tuple:
        """Apply mapping configuration for Campaign, Appeal ID, Fund ID, Package"""
        form_name = str(row.get('Campaign ID', ''))
        campaign_id = str(row.get('Campaign ID', ''))
        
        fy_designation = mapping_config.get('fiscal_year_designation', '')
        forms = mapping_config.get('forms', {})
        match_forms = mapping_config.get('MATCH', [])
        
        campaign = ''
        appeal_id = ''
        fund_id = ''
        package = ''
        
        if form_name in forms:
            form_config = forms[form_name]
            appeal_id = form_config.get('appeal', '')
            fund_id = form_config.get('fund', '')
            
            # Prepend fiscal year designation unless appeal starts with CCDEN
            if appeal_id and not appeal_id.startswith('CCDEN'):
                appeal_id = fy_designation + appeal_id
            
            # Package is MATCH only for forms in the MATCH list
            if form_name in match_forms:
                package = 'MATCH'
        
        return (campaign, appeal_id, fund_id, package)
    
    def _get_gift_subtype(self, row: pd.Series) -> str:
        """Determine gift subtype from Campaign Data 12 and Campaign ID"""
        data_12 = str(row.get('Campaign Data 12', '')).lower()
        campaign_id = str(row.get('Campaign ID', ''))
        is_wyoming = campaign_id.startswith('Y.')
        
        if data_12.startswith('stripe'):
            return 'Stripe-ENWY' if is_wyoming else 'Stripe-ENCO'
        elif data_12.startswith('iats'):
            return 'IATS'
        elif data_12.startswith('blackbaud'):
            return 'BBMS-WY' if is_wyoming else 'BBMS'
        elif 'paypal' in data_12:
            return 'PayPal'
        return ''
    
    def _get_monthly_donor_fields(self, row: pd.Series, re_api=None) -> tuple:
        """Get all monthly donor fields for recurring gifts"""
        campaign_type = str(row.get('Campaign Type', ''))
        
        if campaign_type not in self.RECURRING_TYPES:
            return ('', '', '', '', '', '', '', '', '')
        
        campaign_date_str = str(row.get('Campaign Date', ''))
        data_16_str = str(row.get('Campaign Data 16', ''))
        
        # Parse dates
        try:
            campaign_date = pd.to_datetime(campaign_date_str)
        except:
            campaign_date = None
        
        # Parse Campaign Data 16 (dd/mm/yyyy format)
        try:
            data_16_date = pd.to_datetime(data_16_str, format='%d/%m/%Y')
        except:
            data_16_date = None
        
        # Determine if this is a new recurring gift (dates match)
        is_new_recurring = False
        if campaign_date and data_16_date:
            is_new_recurring = (campaign_date.date() == data_16_date.date())
        
        # Branch for region
        branch = self._get_branch(row)
        region_map = {'Main': 'Denver', 'WSlope': 'Western Slope', 'Wyoming': 'Wyoming'}
        region = region_map.get(branch, '')
        
        # Payment method
        if campaign_type in ['FCR', 'PFCR']:
            payment_method = 'Credit Card/Electronic'
        elif campaign_type in ['FBR', 'PFBR']:
            payment_method = 'ACH'
        else:
            payment_method = ''
        
        if is_new_recurring and campaign_date:
            status = 'Active'
            status_date = campaign_date.strftime('%Y-%m-%d')
            anniversary_desc = campaign_date.strftime('%B')
            anniversary_date = campaign_date.strftime('%Y-%m-%d')
            statement_type = 'Emailed'
            channel = 'Digital -- Recurring'
            gifts_last_month = ''
        else:
            # Need to check for gifts in previous month
            gifts_last_month = 'CHECK'  # Default to CHECK, actual lookup would use RE API
            
            if re_api and campaign_date:
                # Look up gifts from previous month
                re_id = row.get('System Record ID', row.get('RE System Record ID', ''))
                if re_id:
                    gifts = self._lookup_previous_month_gifts(re_api, re_id, campaign_date)
                    if gifts:
                        gifts_last_month = '\n'.join([f"{g['date']} - ${g['amount']}" for g in gifts])
            
            status = ''
            status_date = ''
            anniversary_desc = ''
            anniversary_date = ''
            statement_type = ''
            channel = ''
        
        return (status, status_date, anniversary_desc, anniversary_date, 
                statement_type, channel, payment_method, region, gifts_last_month)
    
    def _lookup_previous_month_gifts(self, re_api, constituent_id: str, current_date: datetime) -> list:
        """Look up gifts from previous month for recurring gift verification"""
        # Calculate previous month date range
        prev_month = current_date - relativedelta(months=1)
        day = current_date.day
        
        # Handle end of month edge cases
        gifts = []
        days_to_check = [day]
        
        # For gifts on 31st, also check 30th, 28th, 29th
        if day >= 29:
            days_to_check.extend([28, 29, 30])
        elif day == 30:
            days_to_check.extend([28, 29])
        
        try:
            for check_day in set(days_to_check):
                try:
                    check_date = prev_month.replace(day=check_day)
                    # This would call the RE API to get gifts
                    # result = re_api.get_constituent_gifts(constituent_id, check_date)
                    # gifts.extend(result)
                except ValueError:
                    # Day doesn't exist in that month
                    continue
        except Exception:
            pass
        
        return gifts
    
    def _get_p2p_fields(self, row: pd.Series, p2p_config: dict) -> tuple:
        """Get P2P fundraising fields"""
        campaign_type = str(row.get('Campaign Type', ''))
        
        if campaign_type not in ['PFCS', 'PFCR', 'PFBS', 'PFBR']:
            return ('', '', '', '', '')
        
        data_15 = str(row.get('Campaign Data 15', ''))
        campaign_number = str(row.get('Campaign Number', ''))
        
        # Get fundraising page name from mapping
        page_name = self.FUNDRAISING_PAGE_MAPPING.get(data_15, '')
        
        # Look up in P2P config
        campaign_name = ''
        solicitor = ''
        
        if campaign_number in p2p_config:
            p2p_entry = p2p_config[campaign_number]
            campaign_name = p2p_entry.get('EN Campaign Name', '')
            solicitor = p2p_entry.get('Solicitor', '')
        else:
            # Add to pending list for manual matching
            self.p2p_pending.append({
                'campaign_number': campaign_number,
                'campaign_type': campaign_type,
                'campaign_data_6': row.get('Campaign Data 6', ''),
                'campaign_data_7': row.get('Campaign Data 7', ''),
                'campaign_data_10': row.get('Campaign Data 10', ''),
                'campaign_data_11': row.get('Campaign Data 11', ''),
                'system_record_id': row.get('System Record ID', row.get('RE System Record ID', '')),
                're_match': None  # Would be populated by RE API lookup
            })
        
        return (data_15, page_name, campaign_number, campaign_name, solicitor)
    
    def _parse_spouse_name(self, row: pd.Series) -> tuple:
        """Parse spouse name from Partner Name field"""
        partner_name = str(row.get('Partner Name', ''))
        first_name = str(row.get('First Name', ''))
        last_name = str(row.get('Last Name', ''))
        
        spouse_first = ''
        spouse_middle = ''
        spouse_last = ''
        
        # Check if partner name contains #
        if '#' not in partner_name or not partner_name.strip():
            return (spouse_first, spouse_middle, spouse_last)
        
        # Remove the # and parse
        partner_name = partner_name.replace('#', '').strip()
        
        if not partner_name:
            return (spouse_first, spouse_middle, spouse_last)
        
        parts = partner_name.split()
        
        if len(parts) >= 1:
            spouse_first = parts[0]
        
        if len(parts) >= 2:
            # Check if second part is a middle initial
            if len(parts[1]) == 1 or (len(parts[1]) == 2 and parts[1].endswith('.')):
                spouse_middle = parts[1].rstrip('.')
                if len(parts) >= 3:
                    spouse_last = ' '.join(parts[2:])
                else:
                    spouse_last = last_name
            else:
                spouse_last = ' '.join(parts[1:])
        else:
            spouse_last = last_name
        
        # If spouse name matches primary name, clear all fields
        if spouse_first == first_name and spouse_last == last_name:
            return ('', '', '')
        
        return (spouse_first, spouse_middle, spouse_last)
    
    def _get_addressee_formula(self, row: pd.Series, output_row: dict) -> str:
        """Generate addressee value (simplified - actual Excel would use formula)"""
        last_name = str(row.get('Last Name', ''))
        spouse_last = str(output_row.get('Spouse Last Name', ''))
        
        if spouse_last and last_name == spouse_last:
            return '49'  # Code for same last name
        elif spouse_last:
            return '48'  # Code for different last names
        return ''
    
    def _get_salutation_formula(self, row: pd.Series, output_row: dict) -> str:
        """Generate salutation value (simplified - actual Excel would use formula)"""
        spouse_last = str(output_row.get('Spouse Last Name', ''))
        
        if not spouse_last:
            return '35'  # Individual salutation code
        return '46'  # Couple salutation code
    
    def _get_gift_reference(self, row: pd.Series, tribute_df: pd.DataFrame) -> str:
        """Match FIM/PFIM records to get gift reference"""
        transaction_id = str(row.get('EN Transaction ID', ''))
        
        if tribute_df is None or len(tribute_df) == 0:
            return ''
        
        # Convert to string for matching
        tribute_df = tribute_df.copy()
        tribute_df['EN Transaction ID'] = tribute_df['EN Transaction ID'].astype(str).str.strip()
        
        match = tribute_df[tribute_df['EN Transaction ID'] == transaction_id.strip()]
        
        if len(match) > 0:
            data_9 = str(match.iloc[0].get('Campaign Data 9', '')).lower()
            data_11 = str(match.iloc[0].get('Campaign Data 11', ''))
            return f"{data_9} {data_11}"
        
        return ''
    
    def _get_no_email(self, row: pd.Series) -> str:
        """Determine no email preference from QCB campaign type"""
        campaign_type = str(row.get('Campaign Type', ''))
        campaign_status = str(row.get('Campaign Status', ''))
        
        if campaign_type == 'QCB':
            return 'FALSE' if campaign_status == 'Y' else 'TRUE'
        return ''
