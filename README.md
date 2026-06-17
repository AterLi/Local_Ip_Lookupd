# IP Info Checker

A Python application for looking up IP address geolocation information. Supports both online API (ipinfo.io) and offline local database modes.

## Features

- **Dual Mode Operation**: 
  - Online: Uses ipinfo.io API (requires token)
  - Offline: Uses local SQLite database with CIDR range matching
  
- **Input Methods**:
  - Upload CSV file with IPs
  - Enter single IP
  - Enter multiple IPs manually

- **Output**: Excel files with IP information

## Installation

```bash
pip install -r requirements.txt
```

## Configuration

### API Mode (Online)

Create a `.env` file in the application directory:

```
IPINFO_TOKEN=your_token_here
```

### Local Database Mode (Offline)

1. Obtain an IP database CSV file (e.g., from ipinfo.io, MaxMind, or other providers)
2. Launch the application
3. Go to Settings → Setup/Update Database
4. Select your CSV file and import

**Expected CSV format (tab-separated):**
```
network	country	country_code	continent	continent_code	asn	as_name	as_domain
1.0.0.0/24	Australia	AU	Oceania	OC	AS13335	Cloudflare, Inc.	cloudflare.com
```

## Usage

```bash
python ip_checker_v2.py
```

### First Time Setup

On first launch, the application will prompt you to:
1. Set up a local database, OR
2. Configure an API token

### Switching Modes

Click the ⚙ Settings button to switch between API and Local Database modes.

### Updating the Database

1. Go to Settings → Setup/Update Database
2. Select a new CSV file
3. Click "Import Database"

The import process converts CIDR ranges to integer ranges for fast O(log n) lookups.

## Technical Details

### Local Database Performance

The local database uses SQLite with optimized indexing:

| Records | DB Size | Lookup Time |
|---------|---------|-------------|
| 100K    | ~10 MB  | <1ms        |
| 1M      | ~100 MB | ~1-2ms      |
| 10M     | ~1 GB   | ~2-5ms      |

### How CIDR Matching Works

IP networks like `1.0.0.0/24` are converted to integer ranges:
- `1.0.0.0` → `16777216` (start)
- `1.0.0.255` → `16777471` (end)

Lookups use: `WHERE start_ip <= ? AND end_ip >= ?`

The most specific (smallest) matching range is returned.

## File Structure

```
ip_checker/
├── ip_checker_v2.py       # Main application
├── local_ip_db.py         # Local database module
├── requirements.txt       # Dependencies
├── sample_ip_data.tsv     # Sample data for testing
├── .env                   # API token (create this)
├── ip_checker_config.db   # App settings (auto-created)
└── db/ip_database.db      # IP data (after import)
```

## Data Sources

You can use IP database files from:
- [IPinfo.io](https://ipinfo.io/products/free-ip-database) - Free tier available
- [MaxMind GeoLite2](https://dev.maxmind.com/geoip/geolite2-free-geolocation-data) - Free with registration
- [DB-IP](https://db-ip.com/db/lite.php) - Free lite version

Convert to the expected TSV format if necessary.
