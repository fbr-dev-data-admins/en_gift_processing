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

    # QCB Campaign IDs that carry consent data
    QCB_CONSENT_CAMPAIGN_IDS = {'General Opt-In', 'Newsletter Opt-In', 'Fundraising Opt-In', 'SMS Opt-In'}
    
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
        'Gift Amount', 'Receipt Amount', 'Donation Type', 'Gifts Last Month', 'Letter Code',
        'EN Donation Form Name', 'Branch', 'Gift Reference', 'Campaign', 'Department', 'Appeal ID',
        'Package', 'Gift Subtype', 'Pay Method', 'Addressee', 'Spouse Addressee', 'Salutation',
        'Spouse Salutation', 'Monthly Donor Status Description', 'Monthly Donor Status Date',
        'Monthly Donor Anniversary Description', 'Monthly Donor Anniversary Date',
        'Monthly Donor Annual Statement Type', 'Monthly Donor Channel', 'Monthly Donor Payment Method',
        'Monthly Donor Region', 'EN Campaign ID', 'EN Campaign Name', 'Solicitor',
        'EN Fundraising Page ID', 'EN Fundraising Page Name', 'Credit Type', 'Stripe Transaction ID',
        'Consents Date', 'Email Channel Opt-In', 'Email Direct Marketing Consent',
        'Email Direct Marketing Consent Statement', 'Email Fundraising Consent',
        'Email Fundraising Consent Statement', 'Email Newsletter Consent',
        'Email Newsletter Consent Statement', 'SMS Channel Response', 'SMS Consent Statement'
    ]
    
    # Columns to fill for QCB (opt-in/out) records - biographical info + consent fields
    QCB_COLUMNS = [
        'Engaging Networks ID', 'Constituent ID', 'Org Name', 'First Name', 'Nickname', 'Middle Name', 'Last Name',
        'Spouse First Name', 'Spouse Nickname', 'Spouse Middle Name', 'Spouse Last Name',
        'Address', 'City', 'State', 'ZIP', 'Country', 'E-mail', 'Cell',
        'Addressee', 'Spouse Addressee', 'Salutation', 'Spouse Salutation',
        'Consents Date', 'Email Channel Opt-In', 'Email Direct Marketing Consent',
        'Email Direct Marketing Consent Statement', 'Email Fundraising Consent',
        'Email Fundraising Consent Statement', 'Email Newsletter Consent',
        'Email Newsletter Consent Statement', 'SMS Channel Response', 'SMS Consent Statement'
    ]
    
    def _clean_id(self, val) -> str:
        """Clean an ID value - remove decimal places from floats, handle NaN"""
        if val is None:
            return ''
        try:
            if pd.isna(val):
                return ''
        except (ValueError, TypeError):
            pass
        
        val_str = str(val).strip()
        if val_str.lower() == 'nan' or val_str == '':
            return ''
        
        if val_str.endswith('.0'):
            val_str = val_str[:-2]
        
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
        self.p2p_config_updates = {}
        self.cached_gifts = {}
        self.gifts_cache_debug = {}
        self.custom_notes = {}
        self.consent_lookup = {}

    def _compute_consent_fields(self, entry: dict) -> dict:
        """
        Derive all 10 consent field values from a consent_lookup entry.

        entry keys: general_status, fundraising_status, newsletter_status,
                    sms_status, email, mobile, dates
        """
        general      = entry.get('general_status')       # 'Y', 'N', or None
        fundraising  = entry.get('fundraising_status')   # 'Y', 'N', or None
        newsletter   = entry.get('newsletter_status')    # 'Y', 'N', or None
        sms          = entry.get('sms_status')           # 'Y', 'N', or None
        email        = entry.get('email', '')
        mobile       = entry.get('mobile', '')
        dates        = entry.get('dates', [])

        result = {
            'Consents Date': '',
            'Email Channel Opt-In': '',
            'Email Direct Marketing Consent': '',
            'Email Direct Marketing Consent Statement': '',
            'Email Fundraising Consent': '',
            'Email Fundraising Consent Statement': '',
            'Email Newsletter Consent': '',
            'Email Newsletter Consent Statement': '',
            'SMS Channel Response': '',
            'SMS Consent Statement': ''
        }

        # Consents Date = earliest Campaign Date across all QCB rows for this supporter
        valid_dates = [d for d in dates if d is not None]
        if valid_dates:
            result['Consents Date'] = min(valid_dates).strftime('%Y-%m-%d')

        # ── Email consent logic ────────────────────────────────────────────────
        if general == 'Y':
            # General opt-in: all email fields become Opt-In
            result['Email Channel Opt-In']                    = 'Opt-In'
            result['Email Direct Marketing Consent']          = 'Opt-In'
            result['Email Direct Marketing Consent Statement'] = f'EN general opt-in for email: {email}'
            result['Email Fundraising Consent']               = 'Opt-In'
            result['Email Fundraising Consent Statement']     = f'EN general opt-in for email: {email}'
            result['Email Newsletter Consent']                = 'Opt-In'
            result['Email Newsletter Consent Statement']      = f'EN general opt-in for email: {email}'

        elif general == 'N':
            fund_y = (fundraising == 'Y')
            news_y = (newsletter  == 'Y')

            if not fund_y and not news_y:
                # All opt-out; Email Channel Opt-In stays blank
                result['Email Direct Marketing Consent']          = 'Opt-Out'
                result['Email Direct Marketing Consent Statement'] = f'EN general opt-out for email: {email}'
                result['Email Fundraising Consent']               = 'Opt-Out'
                result['Email Fundraising Consent Statement']     = f'EN general opt-out for email: {email}'
                result['Email Newsletter Consent']                = 'Opt-Out'
                result['Email Newsletter Consent Statement']      = f'EN general opt-out for email: {email}'

            elif fund_y and not news_y:
                result['Email Channel Opt-In']                    = 'Opt-In'
                result['Email Direct Marketing Consent']          = 'Opt-Out'
                result['Email Direct Marketing Consent Statement'] = f'EN general opt-out for email: {email}'
                result['Email Fundraising Consent']               = 'Opt-In'
                result['Email Fundraising Consent Statement']     = f'EN fundraising opt-in for email: {email}'
                result['Email Newsletter Consent']                = 'Opt-Out'
                result['Email Newsletter Consent Statement']      = f'EN general opt-out for email: {email}'

            elif not fund_y and news_y:
                result['Email Channel Opt-In']                    = 'Opt-In'
                result['Email Direct Marketing Consent']          = 'Opt-Out'
                result['Email Direct Marketing Consent Statement'] = f'EN general opt-out for email: {email}'
                result['Email Fundraising Consent']               = 'Opt-Out'
                result['Email Fundraising Consent Statement']     = f'EN general opt-out for email: {email}'
                result['Email Newsletter Consent']                = 'Opt-In'
                result['Email Newsletter Consent Statement']      = f'EN newsletter opt-in for email: {email}'

            else:  # fund_y and news_y
                result['Email Channel Opt-In']                    = 'Opt-In'
                result['Email Direct Marketing Consent']          = 'Opt-Out'
                result['Email Direct Marketing Consent Statement'] = f'EN general opt-out for email: {email}'
                result['Email Fundraising Consent']               = 'Opt-In'
                result['Email Fundraising Consent Statement']     = f'EN fundraising opt-in for email: {email}'
                result['Email Newsletter Consent']                = 'Opt-In'
                result['Email Newsletter Consent Statement']      = f'EN newsletter opt-in for email: {email}'

        # general is None → all email fields remain blank

        # ── SMS logic (independent of email) ──────────────────────────────────
        if sms == 'Y':
            result['SMS Channel Response']  = 'Opt-In'
            result['SMS Consent Statement'] = f'EN SMS opt-in for phone number: {mobile}'
        elif sms == 'N':
            result['SMS Channel Response']  = 'Opt-Out'
            result['SMS Consent Statement'] = f'EN SMS opt-out for phone number: {mobile}'

        return result

    def transform(
        self,
        df: pd.DataFrame,
        mapping_config: dict,
        p2p_config: dict,
        tribute_df: Optional[pd.DataFrame] = None,
        re_api=None,
        cached_gifts: Optional[Dict[str, List[Dict]]] = None,
        custom_notes_config: Optional[dict] = None,
        md_acks: Optional[List[str]] = None
    ) -> Tuple[pd.DataFrame, pd.DataFrame, List[dict]]:
        """
        Main transformation method
        
        Args:
            df: Input dataframe from EN
            mapping_config: Form mappings configuration
            p2p_config: P2P solicitor configuration
            tribute_df: Optional tribute records for gift reference
            re_api: RE API client (used for batch gift fetching if cached_gifts not provided)
            cached_gifts: Pre-fetched gifts indexed by constituent ID (avoids per-row API calls)
            custom_notes_config: Email->note mapping loaded from custom_notes.json
            md_acks: List of email addresses (lowercase) from md_acks.txt; recurring gifts
                     from these donors receive Letter Code "Monthly Donor Notification"
        
        Returns:
            Tuple of (processed_df, exceptions_df, p2p_pending_list)
        """
        self.exceptions = []
        self.p2p_pending = []
        self.p2p_config_updates = {}
        self.cached_gifts = cached_gifts or {}
        self.gifts_cache_debug = {}
        self.consent_lookup = {}
        
        # Normalize md_acks to a set of lowercase emails for fast lookup
        self.md_acks = set(md_acks) if md_acks else set()
        
        # Use passed-in custom notes config
        raw_notes = custom_notes_config or {}
        self.custom_notes = {k.lower().strip(): v for k, v in raw_notes.items()}
        
        # Create a copy to avoid modifying original
        df = df.copy()

        # ── Build consent_lookup from all 4 QCB Campaign IDs ─────────────────
        # Must be built from the original (unfiltered) df so gift rows can
        # inherit consent values even when a supporter's QCB row is later dropped.
        if (
            'Campaign Type' in df.columns and
            'Supporter ID' in df.columns and
            'Campaign ID' in df.columns
        ):
            # Detect email and mobile columns once
            _email_col = None
            for _c in ['Email', 'E-mail', 'Supporter Email', 'Email Address', 'email', 'EmailAddress']:
                if _c in df.columns:
                    _email_col = _c
                    break

            _cell_col = None
            for _c in ['Mobile Number', 'Mobile Phone', 'MobilePhone', 'Cell Phone', 'CellPhone',
                       'Cell', 'Mobile', 'Phone', 'phone', 'Telephone']:
                if _c in df.columns:
                    _cell_col = _c
                    break

            qcb_consent_df = df[
                (df['Campaign Type'] == 'QCB') &
                (df['Campaign ID'].isin(self.QCB_CONSENT_CAMPAIGN_IDS))
            ]

            for _, row in qcb_consent_df.iterrows():
                supporter_id = str(row.get('Supporter ID', '')).strip()
                if not supporter_id or supporter_id == 'nan':
                    continue

                if supporter_id not in self.consent_lookup:
                    self.consent_lookup[supporter_id] = {
                        'general_status': None,
                        'fundraising_status': None,
                        'newsletter_status': None,
                        'sms_status': None,
                        'email': '',
                        'mobile': '',
                        'dates': []
                    }

                entry = self.consent_lookup[supporter_id]
                campaign_id = str(row.get('Campaign ID', '')).strip()
                status      = str(row.get('Campaign Status', '')).strip()

                if campaign_id == 'General Opt-In':
                    entry['general_status'] = status
                elif campaign_id == 'Fundraising Opt-In':
                    entry['fundraising_status'] = status
                elif campaign_id == 'Newsletter Opt-In':
                    entry['newsletter_status'] = status
                elif campaign_id == 'SMS Opt-In':
                    entry['sms_status'] = status

                # Capture email (first non-empty value found)
                if _email_col and not entry['email']:
                    v = str(row.get(_email_col, '')).strip()
                    if v and v != 'nan':
                        entry['email'] = v

                # Capture mobile – keep rightmost 10 digits
                if _cell_col and not entry['mobile']:
                    v = str(row.get(_cell_col, '')).strip()
                    if v and v != 'nan':
                        digits = ''.join(c for c in v if c.isdigit())
                        entry['mobile'] = digits[-10:] if len(digits) >= 10 else digits

                # Collect Campaign Date for Consents Date calculation
                date_str = str(row.get('Campaign Date', '')).strip()
                if date_str and date_str != 'nan':
                    try:
                        dt = pd.to_datetime(date_str)
                        if not pd.isna(dt):
                            if hasattr(dt, 'to_pydatetime'):
                                dt = dt.to_pydatetime()
                            entry['dates'].append(dt)
                    except Exception:
                        pass
        
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
                    data_9  = str(row.get('Campaign Data 9',  '')).lower()
                    data_11 = str(row.get('Campaign Data 11', ''))
                    tribute_lookup[txn_id] = f"{data_9} {data_11}".strip()
        
        # Initialize output dataframe
        output_df = pd.DataFrame(index=df.index)
        
        # Get fiscal year designation from mapping
        fy_designation = mapping_config.get('fiscal_year_designation', '')
        forms_config   = mapping_config.get('forms', {})
        match_forms    = mapping_config.get('MATCH', [])
        
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
                fund   = form_config.get('fund', '')
                
                if form_name.startswith('CCDEN') or (appeal and appeal.startswith('CCDEN')):
                    campaign = 'CCDEN'
                else:
                    campaign = 'FY' + fy_designation if fy_designation else ''
                
                if appeal and not appeal.startswith('CCDEN'):
                    appeal = fy_designation + appeal
                
                output_df.loc[idx, 'Campaign'] = campaign
                output_df.loc[idx, 'Appeal ID'] = appeal
                output_df.loc[idx, 'Fund ID']   = fund
                output_df.loc[idx, 'Package']   = 'MATCH' if form_name in match_forms else ''
            else:
                output_df.loc[idx, 'Campaign'] = ''
                output_df.loc[idx, 'Appeal ID'] = ''
                output_df.loc[idx, 'Fund ID']   = ''
                output_df.loc[idx, 'Package']   = ''

        # Package overrides: utm_source tsm- prefix and Source = newsletter
        if 'utm_source' in df.columns and 'Source' in df.columns:
            tsm_mask = df['utm_source'].astype(str).str.strip().str.startswith('tsm-', na=False)
            output_df.loc[tsm_mask, 'Package'] = df.loc[tsm_mask, 'Source']

        if 'Source' in df.columns:
            newsletter_mask = df['Source'].astype(str).str.strip().str.lower() == 'newsletter'
            output_df.loc[newsletter_mask, 'Package'] = 'NEWS'
        
        # Gift Subtype
        output_df['Gift Subtype'] = df.apply(self._get_gift_subtype, axis=1)
        
        # Department
        def get_department(campaign_val):
            if pd.isna(campaign_val) or str(campaign_val).strip() == '':
                return ''
            campaign_str = str(campaign_val).strip()
            if campaign_str == 'CCDEN':
                return '93 - Denver Capital Campaign'
            elif campaign_str.startswith('FY'):
                return '15 - General Operating'
            return ''
        
        output_df['Department'] = output_df['Campaign'].apply(get_department)
        
        # Pay Method
        CREDIT_CARD_TYPES = ['FCS', 'FCR', 'PFCS', 'PFCR']
        OTHER_TYPES       = ['FBS', 'FBR', 'PFBS', 'PFBR']
        
        def get_pay_method(ct_val):
            if pd.isna(ct_val) or str(ct_val).strip() == '':
                return ''
            ct = str(ct_val).strip()
            if ct in CREDIT_CARD_TYPES:
                return 'Credit Card'
            elif ct in OTHER_TYPES:
                return 'Other'
            return ''
        
        output_df['Pay Method'] = df['Campaign Type'].apply(get_pay_method) if 'Campaign Type' in df.columns else ''
        
        # Donation Type
        output_df['Donation Type'] = df['Campaign Type'].apply(
            lambda x: 'Recurring' if x in self.RECURRING_TYPES else ''
        ) if 'Campaign Type' in df.columns else ''
        
        # Monthly Donor fields
        for idx, row in df.iterrows():
            monthly_fields = self._get_monthly_donor_fields(row, re_api)
            output_df.loc[idx, 'Monthly Donor Status Description']       = monthly_fields[0]
            output_df.loc[idx, 'Monthly Donor Status Date']              = monthly_fields[1]
            output_df.loc[idx, 'Monthly Donor Anniversary Description']  = monthly_fields[2]
            output_df.loc[idx, 'Monthly Donor Anniversary Date']         = monthly_fields[3]
            output_df.loc[idx, 'Monthly Donor Annual Statement Type']    = monthly_fields[4]
            output_df.loc[idx, 'Monthly Donor Channel']                  = monthly_fields[5]
            output_df.loc[idx, 'Monthly Donor Payment Method']           = monthly_fields[6]
            output_df.loc[idx, 'Monthly Donor Region']                   = monthly_fields[7]
            output_df.loc[idx, 'Gifts Last Month']                       = monthly_fields[8]
        
        # P2P Fields
        for idx, row in df.iterrows():
            p2p_fields = self._get_p2p_fields(row, p2p_config, row_number=idx)
            output_df.loc[idx, 'EN Fundraising Page ID']   = p2p_fields[0]
            output_df.loc[idx, 'EN Fundraising Page Name'] = p2p_fields[1]
            output_df.loc[idx, 'EN Campaign ID']           = p2p_fields[2]
            output_df.loc[idx, 'EN Campaign Name']         = p2p_fields[3]
            output_df.loc[idx, 'Solicitor']                = p2p_fields[4]
        
        # Direct field mappings
        output_df['EN Transaction ID']   = self._safe_column(df, 'EN Transaction ID')
        output_df['Gift Amount']         = self._safe_column(df, 'Campaign Data 4')
        output_df['Receipt Amount']      = self._safe_column(df, 'Campaign Data 4')
        if 'Campaign Data 6' in df.columns:
            daf_mask = df['Campaign Data 6'].astype(str).str.strip().str.upper() == 'DAF'
            output_df.loc[daf_mask, 'Receipt Amount'] = '0.00'
        output_df['Gift Date']           = self._safe_column(df, 'Campaign Date')
        output_df['GL Post Date']        = self._safe_column(df, 'Campaign Date')
        output_df['Stripe Transaction ID'] = self._safe_column(df, 'Campaign Data 2')
        
        # Letter Code
        output_df['Letter Code'] = 'No Letter'
        if 'Campaign Data 6' in df.columns:
            daf_mask = df['Campaign Data 6'].astype(str).str.strip().str.upper() == 'DAF'
            output_df.loc[daf_mask, 'Letter Code'] = 'DAF'
        
        # Engaging Networks ID
        output_df['Engaging Networks ID'] = self._safe_column(df, 'Supporter ID').apply(self._clean_id)
        
        # Org Name
        output_df['Org Name'] = self._safe_column(df, 'Organization or Company',
                                    fallback_col='Company/Org Name')
        
        # Constituent ID
        constituent_id_found = False
        valid_constituent_id_cols = [
            'Raisers Edge Constituent ID',
            'RE Constituent ID',
            "Raiser's Edge Constituent ID",
            'RE System Record ID',
            'System Record ID'
        ]
        for col_name in valid_constituent_id_cols:
            if col_name in df.columns:
                output_df['Constituent ID'] = df[col_name].fillna('').apply(self._clean_id)
                constituent_id_found = True
                break
        if not constituent_id_found and 'Constituent ID' in df.columns:
            output_df['Constituent ID'] = df['Constituent ID'].fillna('').apply(self._clean_id)
            constituent_id_found = True
        if not constituent_id_found:
            output_df['Constituent ID'] = ''
        
        # Name fields
        output_df['First Name']  = self._safe_column(df, 'First Name')
        output_df['Nickname']    = ''
        output_df['Middle Name'] = self._safe_column(df, 'Middle Name')
        output_df['Last Name']   = self._safe_column(df, 'Last Name')
        
        # Spouse name parsing
        has_partner_name = 'Partner Name' in df.columns
        if has_partner_name:
            for idx, row in df.iterrows():
                spouse_first, spouse_middle, spouse_last = self._parse_spouse_name(row)
                output_df.loc[idx, 'Spouse First Name']   = spouse_first
                output_df.loc[idx, 'Spouse Middle Name']  = spouse_middle
                output_df.loc[idx, 'Spouse Last Name']    = spouse_last
        else:
            output_df['Spouse First Name']  = ''
            output_df['Spouse Middle Name'] = ''
            output_df['Spouse Last Name']   = ''
        
        output_df['Spouse Nickname']   = ''
        output_df['Addressee']         = ''
        output_df['Spouse Addressee']  = ''
        output_df['Salutation']        = ''
        output_df['Spouse Salutation'] = ''
        
        # Address
        addr1 = self._safe_column(df, 'Address 1', fallback_col='Address')
        addr2 = self._safe_column(df, 'Address 2')
        output_df['Address'] = (addr1.astype(str).replace('nan', '') + ' ' +
                                addr2.astype(str).replace('nan', '')).str.strip()
        
        output_df['City']  = self._safe_column(df, 'City')
        output_df['State'] = self._safe_column(df, 'State')
        
        # ZIP
        zip_col = None
        for col_name in ['ZIP Code', 'ZIP', 'Zip', 'zip', 'Postal Code', 'PostalCode', 'Zip Code', 'ZipCode']:
            if col_name in df.columns:
                zip_col = col_name
                break
        output_df['ZIP'] = df[zip_col].fillna('').astype(str).str.strip() if zip_col else ''
        
        # Country
        if 'Country' in df.columns:
            output_df['Country'] = df['Country'].apply(
                lambda x: 'United States' if str(x).upper() == 'US' else (x if pd.notna(x) else '')
            )
        else:
            output_df['Country'] = ''
        
        # E-mail
        email_col = None
        for col_name in ['Email', 'E-mail', 'Supporter Email', 'Email Address', 'email', 'EmailAddress']:
            if col_name in df.columns:
                email_col = col_name
                break
        output_df['E-mail'] = df[email_col].fillna('').astype(str).str.strip() if email_col else ''
        
        # Custom notes
        if self.custom_notes:
            for idx in output_df.index:
                email_val = str(output_df.at[idx, 'E-mail']).lower().strip()
                if email_val and email_val in self.custom_notes:
                    note_entry  = self.custom_notes[email_val]
                    custom_note = note_entry.get('Custom Note', '') if isinstance(note_entry, dict) else str(note_entry)
                    if custom_note:
                        current_id = str(output_df.at[idx, 'Constituent ID'])
                        output_df.at[idx, 'Constituent ID'] = f"{current_id}% {custom_note}" if current_id else f"% {custom_note}"
        
        # Monthly Donor Notification Letter Code override
        if self.md_acks and 'Campaign Type' in df.columns:
            for idx in output_df.index:
                campaign_type = str(df.at[idx, 'Campaign Type']).strip()
                if campaign_type in self.RECURRING_TYPES:
                    email_val = str(output_df.at[idx, 'E-mail']).lower().strip()
                    if email_val and email_val in self.md_acks:
                        output_df.at[idx, 'Letter Code'] = 'Monthly Donor Notification'
        
        # Cell
        cell_col = None
        for col_name in ['Mobile Number', 'Mobile Phone', 'MobilePhone', 'Cell Phone', 'CellPhone',
                         'Cell', 'Mobile', 'Phone', 'phone', 'Telephone']:
            if col_name in df.columns:
                cell_col = col_name
                break
        if cell_col:
            def clean_phone(val):
                if pd.isna(val):
                    return ''
                val_str = str(val).strip()
                if val_str.endswith('.0'):
                    val_str = val_str[:-2]
                val_str = val_str.replace('(+1)', '').replace('+1', '').strip()
                if val_str.startswith('1 ') or val_str.startswith('1-'):
                    val_str = val_str[2:].strip()
                elif val_str.startswith('1') and len(val_str.replace('-', '').replace(' ', '').replace('(', '').replace(')', '')) == 11:
                    val_str = val_str[1:].strip()
                elif val_str == '1':
                    val_str = ''
                return ''.join(c for c in val_str if c.isdigit())
            output_df['Cell'] = df[cell_col].apply(clean_phone)
        else:
            output_df['Cell'] = ''
        
        # Credit Type
        def transform_credit_type(val):
            if pd.isna(val) or str(val).strip() == '':
                return ''
            if str(val).strip().upper() == 'DAF':
                return ''
            val_lower = str(val).lower()
            if 'mastercard' in val_lower:
                return 'Mastercard'
            elif 'paypal' in val_lower or 'venmo' in val_lower:
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
        
        output_df['Credit Type'] = df['Campaign Data 6'].apply(transform_credit_type) if 'Campaign Data 6' in df.columns else ''
        
        # Gift Reference
        output_df['Gift Reference'] = df['EN Transaction ID'].apply(
            lambda x: tribute_lookup.get(str(x).strip(), '')
        ) if 'EN Transaction ID' in df.columns else ''
        
        # Key Indicator
        output_df['Key Indicator'] = 'I'

        # ── Populate 10 consent columns for all rows ──────────────────────────
        consent_col_names = [
            'Consents Date', 'Email Channel Opt-In', 'Email Direct Marketing Consent',
            'Email Direct Marketing Consent Statement', 'Email Fundraising Consent',
            'Email Fundraising Consent Statement', 'Email Newsletter Consent',
            'Email Newsletter Consent Statement', 'SMS Channel Response', 'SMS Consent Statement'
        ]
        for col in consent_col_names:
            output_df[col] = ''

        if 'Supporter ID' in df.columns:
            for idx in output_df.index:
                sid = str(df.loc[idx, 'Supporter ID']).strip()
                if sid and sid != 'nan' and sid in self.consent_lookup:
                    fields = self._compute_consent_fields(self.consent_lookup[sid])
                    for col, val in fields.items():
                        output_df.at[idx, col] = val

        # ── QCB row handling ──────────────────────────────────────────────────
        # Suppress QCB rows whose supporter already has surviving gift rows.
        # Deduplicate remaining QCB rows to one per Supporter ID (earliest Campaign Date).
        # VolunteerHub QCBs with no RE Constituent ID are also suppressed.
        qcb_indices = set()
        if 'Campaign Type' in df.columns:
            qcb_indices = set(df[df['Campaign Type'] == 'QCB'].index)

            non_qcb_mask = df['Campaign Type'] != 'QCB'
            supporter_ids_with_gifts = set(
                df.loc[non_qcb_mask, 'Supporter ID'].astype(str).str.strip().unique()
            ) if 'Supporter ID' in df.columns else set()

            # Suppress QCBs whose supporter already has a gift row
            qcb_indices_to_drop = set()
            if 'Supporter ID' in df.columns:
                for idx in qcb_indices:
                    sid = str(df.loc[idx, 'Supporter ID']).strip()
                    if sid and sid in supporter_ids_with_gifts:
                        qcb_indices_to_drop.add(idx)

            surviving_qcb_indices = qcb_indices - qcb_indices_to_drop

            # Suppress VolunteerHub QCBs with no RE Constituent ID
            volunteer_sources = {'VolunteerHub', 'VolunteerHub WS'}
            qcb_volunteer_suppress = set()
            for idx in list(surviving_qcb_indices):
                data_source  = str(df.loc[idx, 'Data Source']).strip() if 'Data Source' in df.columns else ''
                re_const_id  = str(output_df.loc[idx, 'Constituent ID']).strip()
                if data_source in volunteer_sources and not re_const_id:
                    qcb_volunteer_suppress.add(idx)
                    self.exceptions.append({
                        'EN Transaction ID': df.loc[idx, 'EN Transaction ID'] if 'EN Transaction ID' in df.columns else '',
                        'Campaign Type':     df.loc[idx, 'Campaign Type'],
                        'Campaign Status':   df.loc[idx, 'Campaign Status'] if 'Campaign Status' in df.columns else '',
                        'Campaign ID':       df.loc[idx, 'Campaign ID']     if 'Campaign ID'     in df.columns else '',
                        'Reason':            'VolunteerHub record with no RE Constituent ID — excluded from import'
                    })

            surviving_qcb_indices -= qcb_volunteer_suppress

            # Suppress Programs QCBs with no associated gift row
            qcb_programs_suppress = set()
            for idx in list(surviving_qcb_indices):
                internal_grouping = str(df.loc[idx, 'Internal Grouping']).strip() if 'Internal Grouping' in df.columns else ''
                if internal_grouping == 'Programs':
                    qcb_programs_suppress.add(idx)
                    self.exceptions.append({
                        'EN Transaction ID': df.loc[idx, 'EN Transaction ID'] if 'EN Transaction ID' in df.columns else '',
                        'Campaign Type':     df.loc[idx, 'Campaign Type'],
                        'Campaign Status':   df.loc[idx, 'Campaign Status'] if 'Campaign Status' in df.columns else '',
                        'Campaign ID':       df.loc[idx, 'Campaign ID']     if 'Campaign ID'     in df.columns else '',
                        'Reason':            'Programs QCB with no associated gift — excluded from import'
                    })
            
            surviving_qcb_indices -= qcb_programs_suppress

            # Deduplicate: keep one QCB row per Supporter ID (earliest Campaign Date)
            qcb_to_drop_dupe = set()
            if 'Supporter ID' in df.columns:
                seen_sids = {}  # supporter_id -> (best_idx, best_date)
                for idx in list(surviving_qcb_indices):
                    sid = str(df.loc[idx, 'Supporter ID']).strip()
                    if not sid or sid == 'nan':
                        continue
                    date_str = str(df.loc[idx, 'Campaign Date']).strip() if 'Campaign Date' in df.columns else ''
                    try:
                        dt = pd.to_datetime(date_str)
                        if pd.isna(dt):
                            dt = pd.Timestamp.max
                    except Exception:
                        dt = pd.Timestamp.max

                    if sid not in seen_sids:
                        seen_sids[sid] = (idx, dt)
                    else:
                        existing_idx, existing_dt = seen_sids[sid]
                        if dt < existing_dt:
                            qcb_to_drop_dupe.add(existing_idx)
                            seen_sids[sid] = (idx, dt)
                        else:
                            qcb_to_drop_dupe.add(idx)

                surviving_qcb_indices -= qcb_to_drop_dupe

            # Clear non-biographical columns on surviving (standalone) QCB rows
            for col in output_df.columns:
                if col not in self.QCB_COLUMNS:
                    output_df.loc[list(surviving_qcb_indices), col] = ''

            # Drop all suppressed QCB rows
            all_qcb_to_drop = qcb_indices_to_drop | qcb_volunteer_suppress | qcb_programs_suppress | qcb_to_drop_dupe
            if all_qcb_to_drop:
                output_df = output_df.drop(index=list(all_qcb_to_drop))
                df        = df.drop(index=list(all_qcb_to_drop))

            qcb_indices = surviving_qcb_indices
        
        # Format dates to mm/dd/yyyy
        date_columns = [
            'Gift Date', 'GL Post Date',
            'Monthly Donor Status Date', 'Monthly Donor Anniversary Date',
            'Consents Date'
        ]
        for date_col in date_columns:
            if date_col in output_df.columns:
                output_df[date_col] = output_df[date_col].apply(self._format_date)
        
        # Reorder columns to match COLUMN_ORDER
        for col in self.COLUMN_ORDER:
            if col not in output_df.columns:
                output_df[col] = ''
        
        ordered_cols = [col for col in self.COLUMN_ORDER if col in output_df.columns]
        extra_cols   = [col for col in output_df.columns if col not in self.COLUMN_ORDER]
        output_df    = output_df[ordered_cols + extra_cols]
        
        # Sort: Org Name filled (top), regular gifts (middle), QCB (bottom)
        def get_sort_key(idx):
            is_qcb  = idx in qcb_indices
            has_org = output_df.loc[idx, 'Org Name'] != '' and pd.notna(output_df.loc[idx, 'Org Name'])
            if is_qcb:
                return 2
            elif has_org:
                return 0
            else:
                return 1
        
        output_df['_sort_key'] = output_df.index.map(get_sort_key)
        output_df = output_df.sort_values('_sort_key').drop(columns=['_sort_key'])
        output_df = output_df.reset_index(drop=True)
        
        exceptions_df = pd.DataFrame(self.exceptions) if self.exceptions else pd.DataFrame()
        
        return output_df, exceptions_df, self.p2p_pending
    
    def _format_date(self, date_val):
        """Format date value to mm/dd/yyyy"""
        if pd.isna(date_val) or str(date_val).strip() == '':
            return ''
        try:
            date_str = str(date_val).strip()
            for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d']:
                try:
                    dt = datetime.strptime(date_str[:10], fmt)
                    return dt.strftime('%m/%d/%Y')
                except:
                    pass
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
        """Handle PFTC (P2P page creator) records for solicitor matching."""
        campaign_number = self._clean_id(row.get('Campaign Number', ''))
        
        if campaign_number in p2p_config:
            return
        
        system_record_id = ''
        for col_name in ['RE Constituent System Record ID', 'System Record ID', 'RE System Record ID',
                         'Raisers Edge Constituent ID', 'RE Constituent ID', "Raiser's Edge ID",
                         'RE ID', 'SystemRecordID', 'RESystemRecordID']:
            val     = row.get(col_name, '')
            cleaned = self._clean_id(val)
            if cleaned:
                system_record_id = cleaned
                break
        
        campaign_data_10 = str(row.get('Campaign Data 10', ''))
        campaign_data_11 = str(row.get('Campaign Data 11', ''))
        
        if campaign_number:
            p2p_config[campaign_number] = {
                'EN Campaign Name': campaign_data_10,
                'Solicitor': ''
            }
            self.p2p_config_updates[campaign_number] = p2p_config[campaign_number]
        
        self.p2p_pending.append({
            'row_number':       row_number,
            'campaign_number':  campaign_number,
            'campaign_type':    'PFTC',
            'campaign_data_6':  row.get('Campaign Data 6',  ''),
            'campaign_data_7':  row.get('Campaign Data 7',  ''),
            'campaign_data_10': campaign_data_10,
            'campaign_data_11': campaign_data_11,
            'system_record_id': system_record_id,
            're_match':         None
        })
    
    def _filter_by_status(self, df: pd.DataFrame) -> pd.DataFrame:
        """Filter records by campaign type and status"""
        if 'Campaign Type' not in df.columns or 'Campaign Status' not in df.columns:
            return df
        
        valid_mask = pd.Series([False] * len(df), index=df.index)
        
        success_mask = (
            df['Campaign Type'].isin(self.VALID_SUCCESS_TYPES) &
            (df['Campaign Status'] == 'success')
        )
        valid_mask = valid_mask | success_mask
        
        has_paypal = df['Campaign Data 12'].str.contains('paypal', case=False, na=False) if 'Campaign Data 12' in df.columns else pd.Series([False] * len(df), index=df.index)
        
        pending_mask = (
            df['Campaign Type'].isin(self.VALID_PENDING_TYPES) &
            (df['Campaign Status'] == 'pending') &
            (~has_paypal)
        )
        valid_mask = valid_mask | pending_mask
        
        # Include QCB records for all 4 consent Campaign IDs
        qcb_mask = pd.Series([False] * len(df), index=df.index)
        if 'Campaign ID' in df.columns:
            qcb_mask = (
                (df['Campaign Type'] == 'QCB') &
                (df['Campaign ID'].isin(self.QCB_CONSENT_CAMPAIGN_IDS))
            )
        valid_mask = valid_mask | qcb_mask
        
        exception_mask = (
            df['Campaign Type'].isin(self.VALID_SUCCESS_TYPES) &
            ((df['Campaign Status'] == 'reject') | df['Campaign Status'].str.contains('change', case=False, na=False))
        )
        for idx in df[exception_mask].index:
            self.exceptions.append({
                'EN Transaction ID': df.loc[idx, 'EN Transaction ID'] if 'EN Transaction ID' in df.columns else '',
                'Campaign Type':     df.loc[idx, 'Campaign Type'],
                'Campaign Status':   df.loc[idx, 'Campaign Status'],
                'Campaign ID':       df.loc[idx, 'Campaign ID'] if 'Campaign ID' in df.columns else '',
                'Reason':            'Rejected or Changed Status'
            })
        
        pending_paypal_mask = (df['Campaign Status'] == 'pending') & has_paypal
        for idx in df[pending_paypal_mask].index:
            self.exceptions.append({
                'EN Transaction ID': df.loc[idx, 'EN Transaction ID'] if 'EN Transaction ID' in df.columns else '',
                'Campaign Type':     df.loc[idx, 'Campaign Type'],
                'Campaign Status':   df.loc[idx, 'Campaign Status'],
                'Campaign ID':       df.loc[idx, 'Campaign ID'] if 'Campaign ID' in df.columns else '',
                'Reason':            'Pending PayPal Transaction'
            })
        
        filtered_df = df[valid_mask].copy()
        
        if 'Campaign ID' in filtered_df.columns:
            excluded_form_mask = filtered_df['Campaign ID'].str.startswith(('D.8.', 'Y.8.', 'S.8.'), na=False)
            for idx in filtered_df[excluded_form_mask].index:
                self.exceptions.append({
                    'EN Transaction ID': filtered_df.loc[idx, 'EN Transaction ID'] if 'EN Transaction ID' in filtered_df.columns else '',
                    'Campaign Type':     filtered_df.loc[idx, 'Campaign Type'],
                    'Campaign Status':   filtered_df.loc[idx, 'Campaign Status'],
                    'Campaign ID':       filtered_df.loc[idx, 'Campaign ID'],
                    'Reason':            'Excluded Form Name (D.8./Y.8./S.8.)'
                })
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
        data_6  = str(row.get('Campaign Data 6',  '')).strip()
        if data_6.upper() == 'DAF':
            return 'Chariot'
        campaign_id = str(row.get('Campaign ID', ''))
        is_wyoming  = campaign_id.startswith('Y.')
        
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
        
        campaign_date_str = str(row.get('Campaign Date', '')).strip()
        campaign_date     = None
        try:
            campaign_date = pd.to_datetime(campaign_date_str)
        except:
            pass
        
        data_16_str  = str(row.get('Campaign Data 16', '')).strip()
        data_16_date = None
        
        if data_16_str and data_16_str != 'nan':
            for fmt in ['%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d', '%d-%m-%Y']:
                try:
                    data_16_date = datetime.strptime(data_16_str, fmt)
                    break
                except:
                    continue
            if data_16_date is None:
                try:
                    data_16_date = pd.to_datetime(data_16_str, dayfirst=True)
                    if hasattr(data_16_date, 'to_pydatetime'):
                        data_16_date = data_16_date.to_pydatetime()
                except:
                    pass
        
        branch     = self._get_branch(row)
        region_map = {'Main': 'Denver', 'WSlope': 'Western Slope', 'Wyoming': 'Wyoming'}
        region     = region_map.get(branch, '')
        
        if campaign_type in ['FCR', 'PFCR']:
            payment_method = 'Credit Card/Electronic'
        elif campaign_type in ['FBR', 'PFBR']:
            payment_method = 'ACH'
        else:
            payment_method = ''
        
        is_new_recurring = False
        if campaign_date is not None and data_16_date is not None:
            try:
                cd_date  = campaign_date.date() if hasattr(campaign_date, 'date') else campaign_date
                d16_date = data_16_date.date() if hasattr(data_16_date, 'date') else data_16_date
                is_new_recurring = (cd_date == d16_date)
            except:
                pass
        
        re_system_id = ''
        possible_re_id_columns = [
            'RE Constituent System Record ID', 'System Record ID', 'RE System Record ID',
            'Raisers Edge Constituent ID', 'RE Constituent ID', "Raiser's Edge ID",
            'RE ID', 'Constituent ID', 'SystemRecordID', 'RESystemRecordID',
            'REID', 'RE_ID', 'Supporter ID'
        ]
        for col_name in possible_re_id_columns:
            val     = row.get(col_name, '')
            cleaned = self._clean_id(val)
            if cleaned:
                re_system_id = cleaned
                break
        
        if is_new_recurring and campaign_date is not None:
            status           = 'Active'
            status_date      = campaign_date.strftime('%Y-%m-%d')
            anniversary_desc = campaign_date.strftime('%B')
            anniversary_date = campaign_date.strftime('%Y-%m-%d')
            statement_type   = 'Emailed'
            channel          = 'Digital - Recurring'
            gifts_last_month = ''
        else:
            status = status_date = anniversary_desc = anniversary_date = statement_type = channel = ''
            gifts_last_month = 'CHECK'
            
            can_lookup = (
                self.cached_gifts and
                re_system_id and
                re_system_id != 'nan' and
                campaign_date is not None
            )
            
            if can_lookup:
                try:
                    gift_day = campaign_date.day
                    if campaign_date.month == 1:
                        prev_month = 12
                        prev_year  = campaign_date.year - 1
                    else:
                        prev_month = campaign_date.month - 1
                        prev_year  = campaign_date.year
                    
                    import calendar
                    days_in_prev_month = calendar.monthrange(prev_year, prev_month)[1]
                    days_to_check = [min(gift_day, days_in_prev_month)]
                    if gift_day >= 28:
                        days_to_check = list(set([min(gift_day, days_in_prev_month), 28, 29, 30, 31]))
                        days_to_check = [d for d in days_to_check if d <= days_in_prev_month]
                    
                    target_dates      = [f"{prev_year}-{prev_month:02d}-{d:02d}" for d in days_to_check]
                    constituent_gifts = self.cached_gifts.get(re_system_id, [])
                    gifts_found       = [g for g in constituent_gifts if g.get('date', '') in target_dates]
                    
                    if gifts_found:
                        current_gift_amount = None
                        try:
                            amt_str = str(row.get('Campaign Data 4', '')).strip()
                            if amt_str and amt_str != 'nan':
                                current_gift_amount = float(amt_str)
                        except:
                            pass
                        
                        gift_strings = []
                        for gift in gifts_found:
                            gift_date   = gift.get('date', '')
                            gift_amount = gift.get('amount', '')
                            amount_check = ''
                            if current_gift_amount is not None:
                                try:
                                    prev_amt = float(gift_amount) if gift_amount else 0.0
                                    if abs(current_gift_amount - prev_amt) > 0.01:
                                        amount_check = ' - CHECK'
                                except:
                                    pass
                            gift_strings.append(f"{gift_date} - ${gift_amount}{amount_check}")
                        gifts_last_month = '\n'.join(gift_strings)
                    else:
                        gifts_last_month = 'CHECK'
                except Exception:
                    gifts_last_month = 'CHECK'
        
        return (status, status_date, anniversary_desc, anniversary_date,
                statement_type, channel, payment_method, region, gifts_last_month)
    
    def _get_p2p_fields(self, row: pd.Series, p2p_config: dict, row_number: int = None) -> tuple:
        """Get P2P fundraising fields for PFCS, PFCR, PFBS, PFBR campaign types"""
        campaign_type = str(row.get('Campaign Type', ''))
        
        if campaign_type not in self.P2P_GIFT_TYPES:
            return ('', '', '', '', '')
        
        data_15_raw = row.get('Campaign Data 15', '')
        if isinstance(data_15_raw, float):
            if pd.isna(data_15_raw):
                data_15 = ''
            else:
                data_15 = str(int(data_15_raw)) if data_15_raw == int(data_15_raw) else str(data_15_raw)
        else:
            data_15 = self._clean_id(data_15_raw)
        
        page_name        = self.FUNDRAISING_PAGE_MAPPING.get(data_15, '')
        campaign_number  = self._clean_id(row.get('Campaign Number', ''))

        if campaign_number in self.FUNDRAISING_PAGE_MAPPING:
            page_name = self.FUNDRAISING_PAGE_MAPPING.get(campaign_number, '')
            return (campaign_number, page_name, '', '', '')
        
        campaign_name = ''
        solicitor     = ''
        
        if campaign_number in p2p_config:
            p2p_entry     = p2p_config[campaign_number]
            campaign_name = p2p_entry.get('EN Campaign Name', '')
            solicitor     = p2p_entry.get('Solicitor', '')
        else:
            if campaign_number and campaign_number not in [p.get('campaign_number') for p in self.p2p_pending]:
                system_record_id = ''
                for col_name in ['RE Constituent System Record ID', 'System Record ID', 'RE System Record ID',
                                 'Raisers Edge Constituent ID', 'RE Constituent ID', "Raiser's Edge ID",
                                 'RE ID', 'SystemRecordID', 'RESystemRecordID']:
                    val     = row.get(col_name, '')
                    cleaned = self._clean_id(val)
                    if cleaned:
                        system_record_id = cleaned
                        break
                
                self.p2p_pending.append({
                    'row_number':       row_number,
                    'campaign_number':  campaign_number,
                    'campaign_type':    campaign_type,
                    'campaign_id':      str(row.get('Campaign ID', '')),
                    'page_id':          data_15,
                    'page_name':        page_name,
                    'campaign_data_6':  row.get('Campaign Data 6',  ''),
                    'campaign_data_7':  row.get('Campaign Data 7',  ''),
                    'campaign_data_10': row.get('Campaign Data 10', ''),
                    'campaign_data_11': row.get('Campaign Data 11', ''),
                    'system_record_id': system_record_id,
                    're_match':         None
                })
        
        return (data_15, page_name, campaign_number, campaign_name, solicitor)
    
    def _parse_spouse_name(self, row: pd.Series) -> tuple:
        """
        Parse spouse name from Partner Name field.
        """
        partner_name = ''
        if 'Partner Name' in row.index:
            val = row['Partner Name']
            if pd.notna(val):
                partner_name = str(val).strip()
        
        first_name = ''
        last_name  = ''
        if 'First Name' in row.index:
            val = row['First Name']
            if pd.notna(val):
                first_name = str(val).strip()
        if 'Last Name' in row.index:
            val = row['Last Name']
            if pd.notna(val):
                last_name = str(val).strip()
        
        spouse_first = spouse_middle = spouse_last = ''
        
        if not partner_name:
            return (spouse_first, spouse_middle, spouse_last)
        
        import re as regex
        if regex.search(r'\d', partner_name):
            return (partner_name, '', '')
        
        parts = partner_name.split()
        
        if len(parts) >= 1:
            spouse_first = parts[0]
        
        if len(parts) >= 2:
            second_part      = parts[1]
            is_middle_initial = (len(second_part) == 1 or
                                 (len(second_part) == 2 and second_part.endswith('.')))
            if is_middle_initial:
                spouse_middle = second_part.rstrip('.')
                spouse_last   = ' '.join(parts[2:]) if len(parts) >= 3 else last_name
            else:
                spouse_last = ' '.join(parts[1:])
        else:
            spouse_last = last_name
        
        if spouse_first == first_name and spouse_last == last_name:
            return ('', '', '')
        
        return (spouse_first, spouse_middle, spouse_last)
