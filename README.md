# EN Gift Transformation Application

A Streamlit application for transforming Engaging Networks (EN) donation export data into Raiser's Edge (RE) import format.

## Features

- **EN Data Export**: Fetch transaction data directly from Engaging Networks API
- **Gift Transformation**: Comprehensive data transformation following business rules
- **P2P Solicitor Matching**: Match peer-to-peer fundraiser pages to RE constituents
- **RE Sky API Integration**: Lookup constituents, verify gifts, mark solicitors
- **Excel Output**: Generate formatted Excel files with conditional formatting
- **Configuration Management**: JSON-based mapping for appeals, funds, and P2P data

## Project Structure

```
gift_transform_app/
├── streamlit_gift_transform.py    # Main Streamlit application
├── transform.py                   # Gift transformation logic
├── en_api.py                      # Engaging Networks API integration
├── re_skyapi.py                   # Raiser's Edge Sky API integration
├── requirements.txt               # Python dependencies
├── config/
│   ├── mapping.json              # Form name to appeal/fund mappings
│   ├── P2P.json                  # P2P campaign to solicitor mappings
│   └── re_tokens.json            # RE API OAuth tokens (auto-generated)
├── .streamlit/
│   └── secrets.toml              # Application secrets (create this)
└── README.md                      # This file
```

## Installation

### 1. Clone/Download the Application

```bash
# Create project directory
mkdir gift_transform_app
cd gift_transform_app

# Copy all files to this directory
```

### 2. Create Virtual Environment

```bash
python -m venv venv

# Activate (Linux/Mac)
source venv/bin/activate

# Activate (Windows)
venv\Scripts\activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Create Secrets File

Create `.streamlit/secrets.toml`:

```toml
[app]
password = "your-app-password"

[en_api]
token = "your-engaging-networks-api-token"

[re_api]
client_id = "your-blackbaud-client-id"
client_secret = "your-blackbaud-client-secret"
redirect_uri = "http://localhost:8501/callback"
subscription_key = "your-blackbaud-subscription-key"
```

### 5. Run the Application

```bash
streamlit run streamlit_gift_transform.py
```

---

## Raiser's Edge Sky API Setup

### Overview

The RE Sky API uses OAuth 2.0 for authentication and requires a Blackbaud developer account with an approved application.

### Step 1: Create a Blackbaud Developer Account

1. Go to [Blackbaud Developer Portal](https://developer.blackbaud.com/)
2. Click "Sign Up" and create an account
3. Verify your email address

### Step 2: Create an Application

1. Log into the Developer Portal
2. Navigate to **My Applications** → **Add**
3. Fill in the application details:
   - **Application name**: Your organization name + "EN Gift Transform"
   - **Organization name**: Your nonprofit's name
   - **Application website**: Your organization's website
   - **Redirect URIs**: Add these:
     - `http://localhost:8501/callback` (for local development)
     - `https://your-app-domain.com/callback` (for production)
4. Click **Save**

### Step 3: Get Your API Credentials

After creating your application:

1. **Client ID**: Found in your application's overview page
2. **Client Secret**: Click "Show" in the application settings
3. **Subscription Key**: 
   - Go to [API Subscriptions](https://developer.blackbaud.com/subscriptions)
   - Subscribe to the **SKY API Standard Edition** (free tier available)
   - Copy your Primary or Secondary key

### Step 4: Connect to Your RE NXT Environment

Your application needs to be connected to your organization's RE NXT environment:

1. In the Developer Portal, go to your application
2. Click **Marketplace** tab
3. Your organization's RE admin needs to:
   - Go to RE NXT → Control Panel → Applications
   - Find your application and **Connect**
   - Approve the required scopes

### Step 5: Required API Scopes

Ensure your application requests these scopes:

- `Constituent` - Read and write constituent data
- `Gift` - Read gift data
- `Constituent Code` - Add constituent codes (for solicitor marking)

### Step 6: OAuth Authentication Flow

The application handles OAuth automatically:

1. First run: Click "Authenticate with RE" in the sidebar
2. You'll be redirected to Blackbaud to log in
3. Approve the application's access
4. Copy the authorization code from the redirect URL
5. Paste into the application
6. Tokens are saved and refreshed automatically

### API Rate Limits

- **Standard Tier**: 10,000 calls/day
- **Rate limiting**: 5 calls/second recommended
- The application implements automatic retry with backoff

### Common API Endpoints Used

| Endpoint | Purpose |
|----------|---------|
| `GET /constituent/v1/constituents` | Search constituents |
| `GET /constituent/v1/constituents/{id}` | Get constituent details |
| `GET /gift/v1/constituents/{id}/gifts` | Get constituent's gifts |
| `POST /constituent/v1/constituentcodes` | Add solicitor code |

---

## Configuration Files

### mapping.json Structure

```json
{
  "fiscal_year_designation": "FY26",
  "forms": {
    "D.General Donation": {
      "appeal": "DGENWEB",
      "fund": "UNRESTRICTED"
    }
  },
  "MATCH": [
    "D.Corporate Match"
  ]
}
```

**Key points:**
- `fiscal_year_designation`: Prepended to all appeals EXCEPT those starting with "CCDEN"
- `forms`: Maps EN form names (Campaign ID) to RE appeal and fund codes
- `MATCH`: Form names that should have "MATCH" as their Package value

### P2P.json Structure

```json
{
  "123456": {
    "EN Campaign Name": "John's Birthday Fundraiser",
    "Solicitor": "RE12345"
  }
}
```

**Key points:**
- Key is the EN Campaign Number
- `EN Campaign Name`: Display name for the fundraising page
- `Solicitor`: RE Constituent ID who created the page

---

## Transformation Rules

### Campaign Type Filtering

| Campaign Type | Status | Action |
|--------------|--------|--------|
| FCS, FBS, FCR, FBR, PFCS, PFBS, PFCR, PFBR | success | Process |
| FBS, FBR, PFBS, PFBR | pending | Process |
| Any valid type | reject/change | Exception |
| FIM, PFIM | any | Used for tribute matching |
| QCB | Y/N | Email preference |

### Branch Determination

| Campaign ID Prefix | Branch |
|-------------------|--------|
| D. | Main |
| S. | WSlope |
| Y. | Wyoming |

### Gift Subtype Logic

| Campaign Data 12 | Campaign ID | Gift Subtype |
|-----------------|-------------|--------------|
| Starts with "Stripe" | Y.* | Stripe-ENWY |
| Starts with "Stripe" | Other | Stripe-ENCO |
| Starts with "IATS" | Any | IATS |
| Starts with "Blackbaud" | Y.* | BBMS-WY |
| Starts with "Blackbaud" | Other | BBMS |
| Contains "PayPal" | Any | PayPal |

### Monthly Donor Fields

For recurring gifts (FCR, FBR, PFCR, PFBR):

**New recurring (Campaign Data 16 date = Campaign Date):**
- Status: "Active"
- Status Date: Campaign Date
- Anniversary Description: Month name
- Anniversary Date: Campaign Date
- Statement Type: "Emailed"
- Channel: "Digital -- Recurring"
- Payment Method: "Credit Card/Electronic" (FCR/PFCR) or "ACH" (FBR/PFBR)
- Region: Based on Branch

**Existing recurring (dates don't match):**
- Gifts Last Month: Lookup from RE or "CHECK"

---

## Troubleshooting

### EN API Issues

- **Token expired**: Get new token from EN admin console
- **No data returned**: Verify date range format (MMDDYYYY)
- **Encoding errors**: Application tries multiple encodings automatically

### RE API Issues

- **401 Unauthorized**: Refresh tokens or re-authenticate
- **403 Forbidden**: Check application scopes and environment connection
- **429 Rate Limited**: Application implements automatic backoff

### Common Errors

1. **"Missing form names in mapping.json"**
   - Add new form mappings in the Transform tab
   - Or update mapping.json directly

2. **"RE ID is missing"**
   - Records with Key Indicator = "O" (organizations) need manual RE ID entry

3. **"P2P records pending matching"**
   - Complete matching in the P2P Matching tab before final export

---

## Support

For issues related to:
- **Engaging Networks API**: Contact EN support
- **Blackbaud RE Sky API**: Check [Blackbaud Developer Docs](https://developer.blackbaud.com/skyapi/apis)
- **Application bugs**: Review error messages in browser console

---

## License

Internal use only. © Your Organization
