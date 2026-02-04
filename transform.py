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
    P2P_CREATOR_TYPE = 'PFTC'  # P2P page creator
    P2P_GIFT_TYPES = ['PFCS', 'PFCR', 'PFBS', 'PFBR']
    RECURRING_TYPES = ['FCR', 'FBR', 'PFCR', 'PFBR']
    
    # Fundraising page ID mappings
    FUNDRAISING_PAGE_MAPPING = {
        '1064': 'Denver - Food Bank of the Rockies Virtual Food Drives',
        '1068': 'Western Slope - Food Bank of the Rockies Virtual Food Drives',
        '1067': 'Wyoming - Food Bank of the Rockies Virtual Food Drives'
    }
    
    # Column order for output
    COLUMN_ORDER = [
        'Key Indicator', 'Fund ID', 'Engaging Networks ID', 'Constituent ID', 'Org Name',
        'First Name', 'Nickname', 'Middle Name', 'Last Name', 'Spouse First Name',
        'Spouse Nickname', 'Spouse Middle Name', 'Spouse Last Name', 'Address', 'City', 'State', 'ZIP',
        'Country', 'E-mail', 'Cell', 'EN Transaction ID', 'Gift Date', 'GL Post Date',
        'Gift Amount', 'Donation Type', 'Gifts Last Month', 'EN Donation Form Name', 'Branch',
        'Gift Reference', 'Campaign', 'Appeal ID', 'Package', 'Gift Subtype', 'Addressee',
        'Spouse Addressee', 'Salutation', 'Spouse Salutation', 'Monthly Donor Status Description',
        'Monthly Donor Status Date', 'Monthly Donor Anniversary Description', 'Monthly Donor Anniversary Date',
        'Monthly Donor Annual Statement Type', 'Monthly Donor Channel', 'Monthly Donor Payment Method',
        'Monthly Donor Region', 'EN Campaign ID', 'EN Campaign Name', 'Solicitor',
        'EN Fundraising Page ID', 'EN Fundraising Page Name', 'Credit Type', 'Stripe Transaction ID',
        'Requests no email?'
    ]
    
    # Columns to fill for QCB (opt-in/out) records - biographical info only
    QCB_COLUMNS = [
        'Engaging Networks ID', 'Constituent ID', 'Org Name', 'First Name', 'Nickname', 'Middle Name', 'Last Name',
        'Spouse First Name', 'Spouse Nickname', 'Spouse Middle Name', 'Spouse Last Name',
        'Address', 'City', 'State', 'ZIP', 'Country', 'E-mail', 'Cell',
        'Addressee', 'Spouse Addressee', 'Salutation', 'Spouse Salutation', 'Requests no email?'
    ]
    
    def _clean_id(self, val) -> str:
        """Clean an ID value - remove decimal places from floats, handle NaN"""
        if val is None:
            return ''
        # Handle pandas NA values
        try:
            if pd.isna(val):
                return ''
        except (ValueError, TypeError):
            pass
        
        val_str = str(val).strip()
        if val_str.lower() == 'nan' or val_str == '':
            return ''
        
        # Remove .0 decimal from float representation
        if val_str.endswith('.0'):
            val_str = val_str[:-2]
        
        # Also handle cases like 202668.0 by trying to convert to int
        try:
            float_val = float(val_str)
            if float_val == int(float_val):
                return str(int(float_val))
        except (ValueError, TypeError):
            pass
        
        return val_str
    
    def __init__(self):
        self.exceptions = []
        self.p2p_pending = []
        self.debug_log = []  # For debugging RE API calls
        self.p2p_config_updates = {}  # Track P2P config updates made during transform
    
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
        self.debug_log = []  # Reset debug log
        self.p2p_config_updates = {}  # Reset P2P config updates
        
        # Create a copy to avoid modifying original
        df = df.copy()
        
        # Build QCB lookup for "Requests no email?" field
        # Only look at QCB records with Campaign ID = "General Opt-In"
        # Campaign Status = "Y" means FALSE (opted in), else TRUE (opted out)
        qcb_lookup = {}
        if 'Campaign Type' in df.columns and 'Supporter ID' in df.columns and 'Campaign ID' in df.columns:
            qcb_df = df[(df['Campaign Type'] == 'QCB') & (df['Campaign ID'] == 'General Opt-In')].copy()
            for idx, row in qcb_df.iterrows():
                supporter_id = str(row.get('Supporter ID', '')).strip()
                if supporter_id:
                    status = str(row.get('Campaign Status', ''))
                    # Y = opted in (FALSE = does not request no email)
                    # Anything else = opted out (TRUE = requests no email)
                    qcb_lookup[supporter_id] = 'FALSE' if status == 'Y' else 'TRUE'
        
        # Step 1: Handle PFTC (P2P page creators) - these need matching before other processing
        if 'Campaign Type' in df.columns:
            pftc_df = df[df['Campaign Type'] == self.P2P_CREATOR_TYPE].copy()
            for idx, row in pftc_df.iterrows():
                self._handle_pftc_record(row, p2p_config, re_api, row_number=idx)
        
        # Step 2: Filter by campaign type and status
        df = self._filter_by_status(df)
        
        if len(df) == 0:
            return pd.DataFrame(), pd.DataFrame(self.exceptions), self.p2p_pending
        
        # Build tribute lookup from FIM/PFIM records for Gift Reference
        tribute_lookup = {}
        if tribute_df is not None and len(tribute_df) > 0:
            for idx, row in tribute_df.iterrows():
                txn_id = str(row.get('EN Transaction ID', '')).strip()
                if txn_id:
                    # Reference = Campaign Data 9 (lowercase) + " " + Campaign Data 11
                    data_9 = str(row.get('Campaign Data 9', '')).lower()
                    data_11 = str(row.get('Campaign Data 11', ''))
                    tribute_lookup[txn_id] = f"{data_9} {data_11}".strip()
        
        # Initialize output dataframe
        output_df = pd.DataFrame(index=df.index)
        
        # Get fiscal year designation from mapping
        fy_designation = mapping_config.get('fiscal_year_designation', '')
        forms_config = mapping_config.get('forms', {})
        match_forms = mapping_config.get('MATCH', [])
        
        # === TRANSFORMATIONS ===
        
        # Branch: Based on Campaign ID prefix
        output_df['Branch'] = df.apply(self._get_branch, axis=1)
        
        # EN Donation Form Name: Write exact Campaign ID
        output_df['EN Donation Form Name'] = self._safe_column(df, 'Campaign ID')
        
        # Campaign, Appeal ID, Fund ID, Package from mapping.json
        for idx, row in df.iterrows():
            form_name = str(row.get('Campaign ID', ''))
            if form_name in forms_config:
                form_config = forms_config[form_name]
                appeal = form_config.get('appeal', '')
                fund = form_config.get('fund', '')
                
                # Campaign = "FY" + fiscal_year_designation (e.g., "FY" + "26" = "FY26"), EXCEPT CCDEN forms get "CCDEN"
                if form_name.startswith('CCDEN') or (appeal and appeal.startswith('CCDEN')):
                    campaign = 'CCDEN'
                else:
                    campaign = 'FY' + fy_designation if fy_designation else ''  # e.g., "FY26"
                
                # Prepend fiscal year designation to Appeal EXCEPT for appeals starting with CCDEN
                if appeal and not appeal.startswith('CCDEN'):
                    appeal = fy_designation + appeal
                
                output_df.loc[idx, 'Campaign'] = campaign
                output_df.loc[idx, 'Appeal ID'] = appeal
                output_df.loc[idx, 'Fund ID'] = fund
                output_df.loc[idx, 'Package'] = 'MATCH' if form_name in match_forms else ''
            else:
                output_df.loc[idx, 'Campaign'] = ''
                output_df.loc[idx, 'Appeal ID'] = ''
                output_df.loc[idx, 'Fund ID'] = ''
                output_df.loc[idx, 'Package'] = ''
        
        # Gift Subtype: Based on Campaign Data 12 and Campaign ID
        output_df['Gift Subtype'] = df.apply(self._get_gift_subtype, axis=1)
        
        # Donation Type: "Recurring" for recurring campaign types
        output_df['Donation Type'] = df['Campaign Type'].apply(
            lambda x: 'Recurring' if x in self.RECURRING_TYPES else ''
        ) if 'Campaign Type' in df.columns else ''
        
        # Monthly Donor fields
        for idx, row in df.iterrows():
            monthly_fields = self._get_monthly_donor_fields(row, re_api)
            output_df.loc[idx, 'Monthly Donor Status Description'] = monthly_fields[0]
            output_df.loc[idx, 'Monthly Donor Status Date'] = monthly_fields[1]
            output_df.loc[idx, 'Monthly Donor Anniversary Description'] = monthly_fields[2]
            output_df.loc[idx, 'Monthly Donor Anniversary Date'] = monthly_fields[3]
            output_df.loc[idx, 'Monthly Donor Annual Statement Type'] = monthly_fields[4]
            output_df.loc[idx, 'Monthly Donor Channel'] = monthly_fields[5]
            output_df.loc[idx, 'Monthly Donor Payment Method'] = monthly_fields[6]
            output_df.loc[idx, 'Monthly Donor Region'] = monthly_fields[7]
            output_df.loc[idx, 'Gifts Last Month'] = monthly_fields[8]
        
        # P2P Fields (for PFCS, PFCR, PFBS, PFBR)
        for idx, row in df.iterrows():
            p2p_fields = self._get_p2p_fields(row, p2p_config, row_number=idx)
            output_df.loc[idx, 'EN Fundraising Page ID'] = p2p_fields[0]
            output_df.loc[idx, 'EN Fundraising Page Name'] = p2p_fields[1]
            output_df.loc[idx, 'EN Campaign ID'] = p2p_fields[2]
            output_df.loc[idx, 'EN Campaign Name'] = p2p_fields[3]
            output_df.loc[idx, 'Solicitor'] = p2p_fields[4]
        
        # Direct field mappings
        output_df['EN Transaction ID'] = self._safe_column(df, 'EN Transaction ID')
        output_df['Gift Amount'] = self._safe_column(df, 'Campaign Data 4')
        output_df['Gift Date'] = self._safe_column(df, 'Campaign Date')
        output_df['GL Post Date'] = self._safe_column(df, 'Campaign Date')
        output_df['Stripe Transaction ID'] = self._safe_column(df, 'Campaign Data 2')
        
        # Engaging Networks ID - clean to remove .0 from float values
        en_id_col = self._safe_column(df, 'Supporter ID')
        output_df['Engaging Networks ID'] = en_id_col.apply(self._clean_id)
        
        # Org Name (working column, removed in final export)
        output_df['Org Name'] = self._safe_column(df, 'Company/Org Name', 
                                  fallback_col='Company Name')
        
        # Constituent ID - try multiple column name variations
        # NOTE: Must use Raisers Edge Constituent ID, NOT LO Contact ID
        # Explicitly check column names to avoid confusion with LO Contact ID
        constituent_id_found = False
        
        # List of valid column names (excluding LO Contact ID)
        valid_constituent_id_cols = [
            'Raisers Edge Constituent ID', 
            'RE Constituent ID', 
            'Raiser\'s Edge Constituent ID',
            'RE System Record ID', 
            'System Record ID'
        ]
        
        for col_name in valid_constituent_id_cols:
            if col_name in df.columns:
                output_df['Constituent ID'] = df[col_name].fillna('').apply(self._clean_id)
                constituent_id_found = True
                break
        
        # Only use generic 'Constituent ID' if none of the specific ones found AND it's not 'LO Contact ID'
        if not constituent_id_found and 'Constituent ID' in df.columns:
            output_df['Constituent ID'] = df['Constituent ID'].fillna('').apply(self._clean_id)
            constituent_id_found = True
            
        if not constituent_id_found:
            output_df['Constituent ID'] = ''
        
        # Name fields
        output_df['First Name'] = self._safe_column(df, 'First Name')
        output_df['Nickname'] = ''  # Will be Excel formula =First Name
        output_df['Middle Name'] = self._safe_column(df, 'Middle Name')
        output_df['Last Name'] = self._safe_column(df, 'Last Name')
        
        # Spouse name parsing from Partner Name
        # # DEBUG SECTION - Commented out for production
        # # First, let's log what column names are available for debugging
        # available_partner_cols = [c for c in df.columns if any(x in c.lower() for x in ['partner', 'spouse'])]
        
        # Check if Partner Name column exists
        has_partner_name = 'Partner Name' in df.columns
        
        if has_partner_name:
            for idx, row in df.iterrows():
                spouse_first, spouse_middle, spouse_last = self._parse_spouse_name(row)
                output_df.loc[idx, 'Spouse First Name'] = spouse_first
                output_df.loc[idx, 'Spouse Middle Name'] = spouse_middle
                output_df.loc[idx, 'Spouse Last Name'] = spouse_last
        else:
            # No Partner Name column - leave spouse fields empty
            output_df['Spouse First Name'] = ''
            output_df['Spouse Middle Name'] = ''
            output_df['Spouse Last Name'] = ''
        
        # Spouse Nickname - will be Excel formula =Spouse First Name
        output_df['Spouse Nickname'] = ''
        
        # Addressee, Salutation - These will be EXCEL FORMULAS injected in the Excel export
        # Placeholder values here, actual formulas added in create_excel_output
        output_df['Addressee'] = ''
        output_df['Spouse Addressee'] = ''
        output_df['Salutation'] = ''
        output_df['Spouse Salutation'] = ''
        
        # Address: Concat Address 1 + " " + Address 2
        addr1 = self._safe_column(df, 'Address 1', fallback_col='Address')
        addr2 = self._safe_column(df, 'Address 2')
        output_df['Address'] = (addr1.astype(str).replace('nan', '') + ' ' + 
                                addr2.astype(str).replace('nan', '')).str.strip()
        
        output_df['City'] = self._safe_column(df, 'City')
        output_df['State'] = self._safe_column(df, 'State')
        
        # ZIP: Try multiple possible column names from EN
        zip_col = None
        for col_name in ['ZIP Code', 'ZIP', 'Zip', 'zip', 'Postal Code', 'PostalCode', 'Zip Code', 'ZipCode']:
            if col_name in df.columns:
                zip_col = col_name
                break
        if zip_col:
            output_df['ZIP'] = df[zip_col].fillna('').astype(str).str.strip()
        else:
            output_df['ZIP'] = ''
        
        # Country: "US" becomes "United States"
        if 'Country' in df.columns:
            output_df['Country'] = df['Country'].apply(
                lambda x: 'United States' if str(x).upper() == 'US' else (x if pd.notna(x) else '')
            )
        else:
            output_df['Country'] = ''
        
        # E-mail: Try multiple possible column names from EN
        email_col = None
        for col_name in ['Email', 'E-mail', 'Supporter Email', 'Email Address', 'email', 'EmailAddress']:
            if col_name in df.columns:
                email_col = col_name
                break
        if email_col:
            output_df['E-mail'] = df[email_col].fillna('').astype(str).str.strip()
        else:
            output_df['E-mail'] = ''
        
        # Cell: Mobile Number with +1 removed - try multiple column names
        cell_col = None
        for col_name in ['Mobile Number', 'Mobile Phone', 'MobilePhone', 'Cell Phone', 'CellPhone', 'Cell', 'Mobile', 'Phone', 'phone', 'Telephone']:
            if col_name in df.columns:
                cell_col = col_name
                break
        if cell_col:
            # Clean phone numbers: remove (+1), +1, or just 1 at start, clear if only +1
            # Also handle numeric formatting (remove .0)
            def clean_phone(val):
                if pd.isna(val):
                    return ''
                
                # Convert to string and handle numeric values
                val_str = str(val).strip()
                
                # Remove .0 if present (from numeric formatting)
                if val_str.endswith('.0'):
                    val_str = val_str[:-2]
                
                # Remove common prefixes
                val_str = val_str.replace('(+1)', '').replace('+1', '').strip()
                
                # If starts with just "1 " or "1-", remove it
                if val_str.startswith('1 ') or val_str.startswith('1-'):
                    val_str = val_str[2:].strip()
                # If 11 digits starting with 1, remove the 1
                elif val_str.startswith('1') and len(val_str.replace('-', '').replace(' ', '').replace('(', '').replace(')', '')) == 11:
                    val_str = val_str[1:].strip()
                elif val_str == '1':
                    val_str = ''
                
                # Keep only digits, no decimal points
                # Format as text (10 digits): ##########
                digits_only = ''.join(c for c in val_str if c.isdigit())
                
                return digits_only
            
            output_df['Cell'] = df[cell_col].apply(clean_phone)
        else:
            output_df['Cell'] = ''
        
        # Credit Type: Transform Campaign Data 6
        # Mastercard if CONTAINS(mastercard), Paypal if CONTAINS(paypal), Visa if CONTAINS(visa),
        # American Express if CONTAINS(amex), ACH if CONTAINS(ach) or CONTAINS(debit), Discover if CONTAINS(discover)
        def transform_credit_type(val):
            if pd.isna(val) or str(val).strip() == '':
                return ''
            val_lower = str(val).lower()
            if 'mastercard' in val_lower:
                return 'Mastercard'
            elif 'paypal' in val_lower:
                return 'Paypal'
            elif 'visa' in val_lower:
                return 'Visa'
            elif 'amex' in val_lower:
                return 'American Express'
            elif 'ach' in val_lower or 'debit' in val_lower:
                return 'ACH'
            elif 'discover' in val_lower:
                return 'Discover'
            else:
                return str(val)
        
        if 'Campaign Data 6' in df.columns:
            output_df['Credit Type'] = df['Campaign Data 6'].apply(transform_credit_type)
        else:
            output_df['Credit Type'] = ''
        
        # Gift Reference: From FIM/PFIM tribute matching
        output_df['Gift Reference'] = df['EN Transaction ID'].apply(
            lambda x: tribute_lookup.get(str(x).strip(), '')
        ) if 'EN Transaction ID' in df.columns else ''
        
        # Requests no email?: From QCB lookup by Supporter ID
        if 'Supporter ID' in df.columns:
            output_df['Requests no email?'] = df['Supporter ID'].apply(
                lambda x: qcb_lookup.get(str(x).strip(), '')
            )
        else:
            output_df['Requests no email?'] = ''
        
        # Key Indicator: "I" for individual, "O" for organization
        output_df['Key Indicator'] = 'I'
        # Mark as "O" if Org Name has a value
        org_mask = output_df['Org Name'].notna() & (output_df['Org Name'] != '')
        output_df.loc[org_mask, 'Key Indicator'] = 'O'
        
        # For QCB records, clear all non-biographical columns
        # QCB records only get: Engaging Networks ID, Org Name, First Name, Nickname, Middle Name, Last Name,
        # Spouse First Name, Spouse Middle Name, Spouse Last Name, Spouse Nickname,
        # Address, City, State, ZIP, Country, E-mail, Cell, Requests no email?
        if 'Campaign Type' in df.columns:
            qcb_indices = df[df['Campaign Type'] == 'QCB'].index
            for col in output_df.columns:
                if col not in self.QCB_COLUMNS:
                    output_df.loc[qcb_indices, col] = ''
        
        # Format dates to mm/dd/yyyy
        date_columns = ['Gift Date', 'GL Post Date', 'Monthly Donor Status Date', 'Monthly Donor Anniversary Date']
        for date_col in date_columns:
            if date_col in output_df.columns:
                output_df[date_col] = output_df[date_col].apply(self._format_date)
        
        # Reorder columns to match COLUMN_ORDER
        # Add any missing columns with empty values
        for col in self.COLUMN_ORDER:
            if col not in output_df.columns:
                output_df[col] = ''
        
        # Reorder
        ordered_cols = [col for col in self.COLUMN_ORDER if col in output_df.columns]
        # Add any extra columns not in COLUMN_ORDER at the end
        extra_cols = [col for col in output_df.columns if col not in self.COLUMN_ORDER]
        output_df = output_df[ordered_cols + extra_cols]
        
        # Create exceptions dataframe
        exceptions_df = pd.DataFrame(self.exceptions) if self.exceptions else pd.DataFrame()
        
        return output_df, exceptions_df, self.p2p_pending
    
    def _format_date(self, date_val):
        """Format date value to mm/dd/yyyy"""
        if pd.isna(date_val) or str(date_val).strip() == '':
            return ''
        try:
            # Try to parse the date
            date_str = str(date_val).strip()
            # Handle common formats
            for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d']:
                try:
                    dt = datetime.strptime(date_str[:10], fmt)
                    return dt.strftime('%m/%d/%Y')
                except:
                    pass
            # If it's already in the right format or can't be parsed, return as is
            return date_str
        except:
            return str(date_val)
    
    def _safe_column(self, df: pd.DataFrame, col_name: str, fallback_col: str = None) -> pd.Series:
        """Safely get a column, returning empty strings if not found"""
        if col_name in df.columns:
            return df[col_name].fillna('')
        elif fallback_col and fallback_col in df.columns:
            return df[fallback_col].fillna('')
        else:
            return pd.Series([''] * len(df), index=df.index)
    
    def _handle_pftc_record(self, row: pd.Series, p2p_config: dict, re_api, row_number: int = None) -> None:
        """Handle PFTC (P2P page creator) records for solicitor matching"""
        campaign_number = str(row.get('Campaign Number', ''))
        
        # Skip if already in P2P config
        if campaign_number in p2p_config:
            return
        
        # Try multiple column names for RE System Record ID
        system_record_id = ''
        for col_name in ['RE Constituent System Record ID', 'System Record ID', 'RE System Record ID', 'Raisers Edge Constituent ID', 
                         'RE Constituent ID', 'Raiser\'s Edge ID', 'RE ID', 
                         'SystemRecordID', 'RESystemRecordID']:
            val = row.get(col_name, '')
            cleaned = self._clean_id(val)
            if cleaned:
                system_record_id = cleaned
                break
        
        campaign_data_10 = str(row.get('Campaign Data 10', ''))  # EN Campaign Name
        campaign_data_11 = str(row.get('Campaign Data 11', ''))  # Email
        
        re_match = None
        matched_on = None
        
        if re_api and re_api.is_authenticated():
            # Try System Record ID first
            if system_record_id:
                result = re_api.find_p2p_solicitor(system_record_id=system_record_id)
                if result:
                    re_match = result
                    matched_on = 'System Record ID'
            
            # Try email if no match
            if not re_match and campaign_data_11:
                result = re_api.find_p2p_solicitor(email=campaign_data_11)
                if result:
                    re_match = result
                    matched_on = 'Email (Campaign Data 11)'
        
        # If we found a match, update P2P config immediately
        if re_match:
            p2p_config[campaign_number] = {
                'EN Campaign Name': campaign_data_10,
                'Solicitor': re_match.get('id', '')
            }
            # Track this update so calling code can save the config
            self.p2p_config_updates[campaign_number] = p2p_config[campaign_number]
            # NOTE: P2P config will be saved by the calling code after transform completes
            # We no longer call mark_as_solicitor API - just update the config
        
        # Add to pending list for user review (even if matched, so user can verify)
        self.p2p_pending.append({
            'row_number': row_number,
            'campaign_number': campaign_number,
            'campaign_type': 'PFTC',
            'campaign_data_6': row.get('Campaign Data 6', ''),
            'campaign_data_7': row.get('Campaign Data 7', ''),
            'campaign_data_10': campaign_data_10,  # Name / EN Campaign Name
            'campaign_data_11': campaign_data_11,  # Email
            'system_record_id': system_record_id,
            're_match': re_match
        })
    
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
        # BUT NOT if Campaign Data 12 contains PayPal - those go to exceptions
        has_paypal = df['Campaign Data 12'].str.contains('paypal', case=False, na=False) if 'Campaign Data 12' in df.columns else pd.Series([False] * len(df), index=df.index)
        
        pending_mask = (
            df['Campaign Type'].isin(self.VALID_PENDING_TYPES) & 
            (df['Campaign Status'] == 'pending') &
            (~has_paypal)  # Exclude PayPal pending transactions
        )
        valid_mask = valid_mask | pending_mask
        
        # Include QCB records with Campaign ID = "General Opt-In" for biographical data
        qcb_mask = pd.Series([False] * len(df), index=df.index)
        if 'Campaign ID' in df.columns:
            qcb_mask = (df['Campaign Type'] == 'QCB') & (df['Campaign ID'] == 'General Opt-In')
        valid_mask = valid_mask | qcb_mask
        
        # Track exceptions (reject or change status)
        exception_mask = (
            df['Campaign Type'].isin(self.VALID_SUCCESS_TYPES) & 
            ((df['Campaign Status'] == 'reject') | df['Campaign Status'].str.contains('change', case=False, na=False))
        )
        
        for idx in df[exception_mask].index:
            self.exceptions.append({
                'EN Transaction ID': df.loc[idx, 'EN Transaction ID'] if 'EN Transaction ID' in df.columns else '',
                'Campaign Type': df.loc[idx, 'Campaign Type'],
                'Campaign Status': df.loc[idx, 'Campaign Status'],
                'Campaign ID': df.loc[idx, 'Campaign ID'] if 'Campaign ID' in df.columns else '',
                'Reason': 'Rejected or Changed Status'
            })
        
        # Track pending PayPal transactions as exceptions
        pending_paypal_mask = (
            df['Campaign Status'] == 'pending'
        ) & has_paypal
        
        for idx in df[pending_paypal_mask].index:
            self.exceptions.append({
                'EN Transaction ID': df.loc[idx, 'EN Transaction ID'] if 'EN Transaction ID' in df.columns else '',
                'Campaign Type': df.loc[idx, 'Campaign Type'],
                'Campaign Status': df.loc[idx, 'Campaign Status'],
                'Campaign ID': df.loc[idx, 'Campaign ID'] if 'Campaign ID' in df.columns else '',
                'Reason': 'Pending PayPal Transaction'
            })
        
        # Filter to valid records first
        filtered_df = df[valid_mask].copy()
        
        # Filter out form names starting with D.8., Y.8., or S.8. and add to exceptions
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
        
        # Only process recurring types
        if campaign_type not in self.RECURRING_TYPES:
            return ('', '', '', '', '', '', '', '', '')
        
        # Parse Campaign Date (typically yyyy-mm-dd format from EN)
        campaign_date_str = str(row.get('Campaign Date', '')).strip()
        campaign_date = None
        try:
            campaign_date = pd.to_datetime(campaign_date_str)
        except:
            pass
        
        # Parse Campaign Data 16 (dd/mm/yyyy format) - this is the recurring start date
        data_16_str = str(row.get('Campaign Data 16', '')).strip()
        data_16_date = None
        
        if data_16_str and data_16_str != 'nan':
            # Try multiple date formats
            for fmt in ['%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d', '%d-%m-%Y']:
                try:
                    data_16_date = datetime.strptime(data_16_str, fmt)
                    break
                except:
                    continue
            
            # If still None, try pandas
            if data_16_date is None:
                try:
                    data_16_date = pd.to_datetime(data_16_str, dayfirst=True)
                    if hasattr(data_16_date, 'to_pydatetime'):
                        data_16_date = data_16_date.to_pydatetime()
                except:
                    pass
        
        # Determine branch for region mapping
        branch = self._get_branch(row)
        region_map = {'Main': 'Denver', 'WSlope': 'Western Slope', 'Wyoming': 'Wyoming'}
        region = region_map.get(branch, '')
        
        # Payment method based on campaign type
        if campaign_type in ['FCR', 'PFCR']:
            payment_method = 'Credit Card/Electronic'
        elif campaign_type in ['FBR', 'PFBR']:
            payment_method = 'ACH'
        else:
            payment_method = ''
        
        # Check if this is a NEW recurring gift (Campaign Data 16 date = Campaign Date)
        is_new_recurring = False
        if campaign_date is not None and data_16_date is not None:
            try:
                cd_date = campaign_date.date() if hasattr(campaign_date, 'date') else campaign_date
                d16_date = data_16_date.date() if hasattr(data_16_date, 'date') else data_16_date
                is_new_recurring = (cd_date == d16_date)
            except:
                pass
        
        # Debug log entry
        en_txn_id = str(row.get('EN Transaction ID', ''))
        
        # Try multiple column names for RE System Record ID
        re_system_id = ''
        re_id_col_found = 'NOT FOUND'
        for col_name in ['RE Constituent System Record ID', 'System Record ID', 'RE System Record ID', 'Raisers Edge Constituent ID', 
                         'RE Constituent ID', 'Raiser\'s Edge ID', 'RE ID', 'Constituent ID',
                         'SystemRecordID', 'RESystemRecordID', 'REID', 'RE_ID', 'Supporter ID']:
            val = row.get(col_name, '')
            cleaned = self._clean_id(val)
            if cleaned:
                re_system_id = cleaned
                re_id_col_found = col_name
                break
        
        # Debug: Show what row index looks like
        row_index_info = str(row.index.tolist()[:10]) if hasattr(row, 'index') else 'no index'
        
        # # DEBUG SECTION - Commented out for production
        # # Debug logging for RE API calls
        # debug_entry = {
        #     'EN Transaction ID': en_txn_id,
        #     'Campaign Type': campaign_type,
        #     'Campaign Date': campaign_date_str,
        #     'Campaign Data 16': data_16_str,
        #     'Parsed Campaign Date': str(campaign_date) if campaign_date else 'PARSE FAILED',
        #     'Parsed Data 16': str(data_16_date) if data_16_date else 'PARSE FAILED',
        #     'Is New Recurring': is_new_recurring,
        #     'RE System Record ID': re_system_id if re_system_id else '(empty)',
        #     'RE ID Column Found': re_id_col_found,
        #     'Row Keys (sample)': row_index_info,
        #     'RE API Called': False,
        #     'RE API Response': None,
        #     'Gifts Last Month Result': ''
        # }
        
        if is_new_recurring and campaign_date is not None:
            # New recurring gift - populate all monthly donor fields
            status = 'Active'
            status_date = campaign_date.strftime('%Y-%m-%d')
            anniversary_desc = campaign_date.strftime('%B')  # Month name (MMMM)
            anniversary_date = campaign_date.strftime('%Y-%m-%d')
            statement_type = 'Emailed'
            channel = 'Digital -- Recurring'
            gifts_last_month = ''
            # debug_entry['Gifts Last Month Result'] = '(New recurring - no lookup needed)'
        else:
            # Existing recurring gift - need to look up previous month's gifts
            status = ''
            status_date = ''
            anniversary_desc = ''
            anniversary_date = ''
            statement_type = ''
            channel = ''
            gifts_last_month = 'CHECK'  # Default to CHECK
            
            # Try to call RE API if available and we have a RE System Record ID
            if re_api and re_api.is_authenticated() and re_system_id and re_system_id != 'nan' and campaign_date is not None:
                # debug_entry['RE API Called'] = True
                try:
                    # Calculate the day to look for in previous month
                    gift_day = campaign_date.day
                    
                    # Get previous month
                    if campaign_date.month == 1:
                        prev_month = 12
                        prev_year = campaign_date.year - 1
                    else:
                        prev_month = campaign_date.month - 1
                        prev_year = campaign_date.year
                    
                    # Handle end of month (31st -> 30th/28th/29th)
                    import calendar
                    days_in_prev_month = calendar.monthrange(prev_year, prev_month)[1]
                    
                    # Days to check (for end of month handling)
                    days_to_check = [min(gift_day, days_in_prev_month)]
                    if gift_day >= 28:
                        # Also check 28th and 29th for Feb edge cases
                        days_to_check = list(set([min(gift_day, days_in_prev_month), 28, 29, 30, 31]))
                        days_to_check = [d for d in days_to_check if d <= days_in_prev_month]
                    
                    # debug_entry['Days to Check'] = days_to_check
                    # debug_entry['Previous Month/Year'] = f"{prev_month}/{prev_year}"
                    # debug_entry['Target Dates'] = [f"{prev_year}-{prev_month:02d}-{d:02d}" for d in days_to_check]
                    
                    # Call RE API to get gifts
                    gifts_found, api_debug = re_api.get_constituent_gifts(
                        constituent_id=re_system_id,
                        year=prev_year,
                        month=prev_month,
                        days=days_to_check
                    )
                    
                    # # Add API debug info
                    # debug_entry['API Endpoint'] = api_debug.get('endpoint', '')
                    # debug_entry['API Params'] = str(api_debug.get('params', {}))
                    # debug_entry['API Status'] = api_debug.get('response_status', '')
                    # debug_entry['API Gifts Count'] = api_debug.get('gifts_count', 0)
                    # debug_entry['API Filtered Count'] = api_debug.get('filtered_count', 'N/A')
                    # if api_debug.get('error'):
                    #     debug_entry['API Error'] = api_debug.get('error')
                    
                    # debug_entry['RE API Response'] = str(gifts_found) if gifts_found else 'Empty'
                    
                    if gifts_found and len(gifts_found) > 0:
                        # Get current transaction gift amount for comparison
                        current_gift_amount = None
                        try:
                            current_amount_str = str(row.get('Campaign Data 4', '')).strip()
                            if current_amount_str and current_amount_str != 'nan':
                                current_gift_amount = float(current_amount_str)
                        except:
                            pass
                        
                        # Format: "date - $amount" with line breaks
                        # Append " - CHECK" if amount differs from current gift amount
                        gift_strings = []
                        for gift in gifts_found:
                            gift_date = gift.get('date', '')
                            gift_amount = gift.get('amount', '')
                            
                            # Check if amounts differ
                            amount_check = ''
                            if current_gift_amount is not None:
                                try:
                                    previous_amount = float(gift_amount) if gift_amount else 0.0
                                    if abs(current_gift_amount - previous_amount) > 0.01:  # Allow for floating point precision
                                        amount_check = ' - CHECK'
                                except:
                                    pass
                            
                            gift_strings.append(f"{gift_date} - ${gift_amount}{amount_check}")
                        
                        gifts_last_month = '\n'.join(gift_strings)
                        # debug_entry['Gifts Last Month Result'] = f"Found {len(gifts_found)} gifts"
                    else:
                        gifts_last_month = 'CHECK'
                        # debug_entry['Gifts Last Month Result'] = 'No gifts found - CHECK'
                        
                except Exception as e:
                    # debug_entry['RE API Response'] = f"ERROR: {str(e)}"
                    # debug_entry['Gifts Last Month Result'] = f'API Error - CHECK'
                    gifts_last_month = 'CHECK'
            # else:
            #     # Log why we didn't call the API
            #     if not re_api:
            #         debug_entry['Gifts Last Month Result'] = 'No RE API configured - CHECK'
            #     elif not re_api.is_authenticated():
            #         debug_entry['Gifts Last Month Result'] = 'RE API not authenticated - CHECK'
            #     elif not re_system_id or re_system_id == 'nan':
            #         debug_entry['Gifts Last Month Result'] = 'No RE System Record ID - CHECK'
            #     elif campaign_date is None:
            #         debug_entry['Gifts Last Month Result'] = 'Campaign Date parse failed - CHECK'
        
        # self.debug_log.append(debug_entry)
        
        return (status, status_date, anniversary_desc, anniversary_date, 
                statement_type, channel, payment_method, region, gifts_last_month)
    
    def _get_p2p_fields(self, row: pd.Series, p2p_config: dict, row_number: int = None) -> tuple:
        """Get P2P fundraising fields for PFCS, PFCR, PFBS, PFBR campaign types"""
        campaign_type = str(row.get('Campaign Type', ''))
        
        if campaign_type not in self.P2P_GIFT_TYPES:
            return ('', '', '', '', '')
        
        # EN Fundraising Page ID = Campaign Data 15 (clean to remove .0)
        data_15_raw = row.get('Campaign Data 15', '')
        # Handle float values from pandas
        if isinstance(data_15_raw, float):
            if pd.isna(data_15_raw):
                data_15 = ''
            else:
                data_15 = str(int(data_15_raw)) if data_15_raw == int(data_15_raw) else str(data_15_raw)
        else:
            data_15 = self._clean_id(data_15_raw)
        
        # EN Fundraising Page Name from mapping
        page_name = self.FUNDRAISING_PAGE_MAPPING.get(data_15, '')
        
        # EN Campaign ID = Campaign Number
        campaign_number = str(row.get('Campaign Number', ''))
        
        # Look up in P2P config for EN Campaign Name and Solicitor
        campaign_name = ''
        solicitor = ''
        
        if campaign_number in p2p_config:
            p2p_entry = p2p_config[campaign_number]
            campaign_name = p2p_entry.get('EN Campaign Name', '')
            solicitor = p2p_entry.get('Solicitor', '')
        else:
            # Add to pending list for manual matching if not found
            if campaign_number and campaign_number not in [p.get('campaign_number') for p in self.p2p_pending]:
                # Get system record ID using multiple column name variations
                system_record_id = ''
                for col_name in ['RE Constituent System Record ID', 'System Record ID', 'RE System Record ID', 'Raisers Edge Constituent ID', 
                                 'RE Constituent ID', 'Raiser\'s Edge ID', 'RE ID', 
                                 'SystemRecordID', 'RESystemRecordID']:
                    val = row.get(col_name, '')
                    cleaned = self._clean_id(val)
                    if cleaned:
                        system_record_id = cleaned
                        break
                
                self.p2p_pending.append({
                    'row_number': row_number,
                    'campaign_number': campaign_number,
                    'campaign_type': campaign_type,
                    'campaign_id': str(row.get('Campaign ID', '')),  # Region page (form name)
                    'page_id': data_15,  # EN Fundraising Page ID
                    'page_name': page_name,
                    'campaign_data_6': row.get('Campaign Data 6', ''),
                    'campaign_data_7': row.get('Campaign Data 7', ''),
                    'campaign_data_10': row.get('Campaign Data 10', ''),
                    'campaign_data_11': row.get('Campaign Data 11', ''),
                    'system_record_id': system_record_id,
                    're_match': None
                })
        
        return (data_15, page_name, campaign_number, campaign_name, solicitor)
    
    def _parse_spouse_name(self, row: pd.Series) -> tuple:
        """
        Parse spouse name from Partner Name field.
        
        Logic:
        - IF Partner Name contains a digit/number, write entire value to Spouse First Name (no parsing)
        - Otherwise parse normally:
          - Spouse First Name = TEXTBEFORE first " "
          - IF second part is single letter or letter+period, write to Spouse Middle Name
          - IF TEXTAFTER (middle name or " ") <>"" THEN Spouse Last Name = TEXTAFTER, ELSE Spouse Last Name = Last Name
        - IF Spouse First Name = First Name and Spouse Last Name = Last Name, clear all three fields
        """
        # Get Partner Name value directly from the Series
        partner_name = ''
        
        # Access using direct indexing since it's a pandas Series
        if 'Partner Name' in row.index:
            val = row['Partner Name']
            if pd.notna(val):
                partner_name = str(val).strip()
        
        # Get first and last name for comparison
        first_name = ''
        last_name = ''
        
        if 'First Name' in row.index:
            val = row['First Name']
            if pd.notna(val):
                first_name = str(val).strip()
        
        if 'Last Name' in row.index:
            val = row['Last Name']
            if pd.notna(val):
                last_name = str(val).strip()
        
        spouse_first = ''
        spouse_middle = ''
        spouse_last = ''
        
        # If no partner name, return empty
        if not partner_name:
            return (spouse_first, spouse_middle, spouse_last)
        
        # Check if partner name contains any digit - if so, write entire value to Spouse First Name only
        import re as regex
        if regex.search(r'\d', partner_name):
            # Contains a digit - write entire value to Spouse First Name, leave others blank
            spouse_first = partner_name
            return (spouse_first, '', '')
        
        # No digits - parse normally
        parts = partner_name.split()
        
        if len(parts) >= 1:
            # Spouse First Name = TEXTBEFORE first " "
            spouse_first = parts[0]
        
        if len(parts) >= 2:
            # Check if second part is a middle initial (single letter or letter + period)
            second_part = parts[1]
            is_middle_initial = (len(second_part) == 1 or 
                                (len(second_part) == 2 and second_part.endswith('.')))
            
            if is_middle_initial:
                # IF single letter or single letter + "." then write to Spouse Middle Name
                spouse_middle = second_part.rstrip('.')
                if len(parts) >= 3:
                    # Spouse Last Name = remaining text after middle name
                    spouse_last = ' '.join(parts[2:])
                else:
                    # No last name provided after middle initial, use primary last name
                    spouse_last = last_name
            else:
                # No middle initial, rest is last name
                spouse_last = ' '.join(parts[1:])
        else:
            # Only first name provided, use primary last name
            spouse_last = last_name
        
        # IF Spouse First Name = First Name and Spouse Last Name = Last Name, clear all three fields
        if spouse_first == first_name and spouse_last == last_name:
            return ('', '', '')
        
        return (spouse_first, spouse_middle, spouse_last)
