"""
Local IP Database Module
Converts CIDR network ranges to a SQLite database for fast IP lookups.
Supports both IPv4 and IPv6 addresses.
"""

import sqlite3
import ipaddress
import csv
import os
import logging
from datetime import datetime
from typing import Optional, Dict, Any, Callable
from pathlib import Path

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")


class LocalIPDatabase:
    """Fast local IP geolocation database using SQLite with optimized range queries."""

    def __init__(self, db_path: str = "ip_database.db"):
        self.db_path = db_path
        self.conn = None
        self._connect()

    def _connect(self):
        """Establish database connection."""
        self.conn = sqlite3.connect(self.db_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.execute("PRAGMA cache_size=10000")

    def create_database(self):
        """Create the database schema with optimized indexes."""
        cursor = self.conn.cursor()

        cursor.execute("DROP TABLE IF EXISTS ip_ranges_v4")
        cursor.execute("DROP TABLE IF EXISTS ip_ranges_v6")
        cursor.execute("DROP TABLE IF EXISTS metadata")

        # Create IPv4 table (uses INTEGER for fast numeric comparisons)
        cursor.execute("""
            CREATE TABLE ip_ranges_v4 (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                network TEXT NOT NULL,
                start_ip INTEGER NOT NULL,
                end_ip INTEGER NOT NULL,
                country TEXT,
                country_code TEXT,
                continent TEXT,
                continent_code TEXT,
                asn TEXT,
                as_name TEXT,
                as_domain TEXT
            )
        """)

        # Create IPv6 table (uses TEXT for 128-bit addresses stored as hex)
        # Format: 32-character zero-padded hex string for proper string comparison
        cursor.execute("""
            CREATE TABLE ip_ranges_v6 (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                network TEXT NOT NULL,
                start_ip TEXT NOT NULL,
                end_ip TEXT NOT NULL,
                country TEXT,
                country_code TEXT,
                continent TEXT,
                continent_code TEXT,
                asn TEXT,
                as_name TEXT,
                as_domain TEXT
            )
        """)

        # Create metadata table
        cursor.execute("""
            CREATE TABLE metadata (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)

        # Create indexes for IPv4 range queries
        cursor.execute("CREATE INDEX idx_v4_start_ip ON ip_ranges_v4(start_ip)")
        cursor.execute("CREATE INDEX idx_v4_end_ip ON ip_ranges_v4(end_ip)")
        cursor.execute("CREATE INDEX idx_v4_ip_range ON ip_ranges_v4(start_ip, end_ip)")

        # Create indexes for IPv6 range queries (string comparison works with zero-padded hex)
        cursor.execute("CREATE INDEX idx_v6_start_ip ON ip_ranges_v6(start_ip)")
        cursor.execute("CREATE INDEX idx_v6_end_ip ON ip_ranges_v6(end_ip)")
        cursor.execute("CREATE INDEX idx_v6_ip_range ON ip_ranges_v6(start_ip, end_ip)")

        self.conn.commit()
        logging.info("Database schema created successfully")

    @staticmethod
    def ip_to_int(ip: str) -> int:
        """Convert IPv4 address string to integer."""
        return int(ipaddress.ip_address(ip))

    @staticmethod
    def ipv6_to_hex(ip: str) -> str:
        """Convert IPv6 address to 32-character zero-padded hex string."""
        addr = ipaddress.ip_address(ip)
        return format(int(addr), '032x')

    @staticmethod
    def is_ipv6_network(network: str) -> bool:
        """Check if a network is IPv6."""
        try:
            net = ipaddress.ip_network(network, strict=False)
            return net.version == 6
        except ValueError:
            return False

    @staticmethod
    def network_to_range_v4(network: str) -> tuple:
        """Convert IPv4 CIDR notation to (start_ip, end_ip) integers."""
        net = ipaddress.ip_network(network, strict=False)
        return int(net.network_address), int(net.broadcast_address)

    @staticmethod
    def network_to_range_v6(network: str) -> tuple:
        """Convert IPv6 CIDR notation to (start_ip, end_ip) hex strings."""
        net = ipaddress.ip_network(network, strict=False)
        start_hex = format(int(net.network_address), '032x')
        end_hex = format(int(net.broadcast_address), '032x')
        return start_hex, end_hex

    def import_from_csv(self, csv_path: str, delimiter: str = '\t',
                        batch_size: int = 10000,
                        progress_callback: Optional[Callable[[int, int], None]] = None) -> int:
        """
        Import IP ranges from a CSV/TSV file.

        Args:
            csv_path: Path to the CSV file
            delimiter: Field delimiter (tab for TSV, comma for CSV)
            batch_size: Number of records to insert per transaction
            progress_callback: Optional callback(current, total) for progress updates

        Returns:
            Total number of records imported
        """
        if not os.path.exists(csv_path):
            raise FileNotFoundError(f"CSV file not found: {csv_path}")

        # Count total lines for progress
        total_lines = sum(1 for _ in open(csv_path, 'r', encoding='utf-8')) - 1

        self.create_database()
        cursor = self.conn.cursor()

        total_imported = 0
        batch_v4 = []
        batch_v6 = []
        errors = 0
        v4_count = 0
        v6_count = 0

        logging.info(f"Starting import from {csv_path} ({total_lines:,} records)")

        with open(csv_path, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f, delimiter=delimiter)

            for row in reader:
                try:
                    network = row.get('network', '').strip()
                    if not network:
                        continue

                    # Prepare common data
                    row_data = (
                        row.get('country', ''),
                        row.get('country_code', ''),
                        row.get('continent', ''),
                        row.get('continent_code', ''),
                        row.get('asn', ''),
                        row.get('as_name', ''),
                        row.get('as_domain', '')
                    )

                    if self.is_ipv6_network(network):
                        # IPv6 - store as hex strings
                        start_ip, end_ip = self.network_to_range_v6(network)
                        batch_v6.append((network, start_ip, end_ip) + row_data)
                        v6_count += 1
                    else:
                        # IPv4 - store as integers
                        start_ip, end_ip = self.network_to_range_v4(network)
                        batch_v4.append((network, start_ip, end_ip) + row_data)
                        v4_count += 1

                    # Flush batches when full
                    if len(batch_v4) >= batch_size:
                        cursor.executemany("""
                            INSERT INTO ip_ranges_v4 
                            (network, start_ip, end_ip, country, country_code, 
                             continent, continent_code, asn, as_name, as_domain)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, batch_v4)
                        total_imported += len(batch_v4)
                        batch_v4 = []
                        self.conn.commit()

                        if progress_callback:
                            progress_callback(total_imported + len(batch_v6), total_lines)
                        logging.info(f"Imported {total_imported + len(batch_v6):,}/{total_lines:,} records...")

                    if len(batch_v6) >= batch_size:
                        cursor.executemany("""
                            INSERT INTO ip_ranges_v6 
                            (network, start_ip, end_ip, country, country_code, 
                             continent, continent_code, asn, as_name, as_domain)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, batch_v6)
                        total_imported += len(batch_v6)
                        batch_v6 = []
                        self.conn.commit()

                        if progress_callback:
                            progress_callback(total_imported + len(batch_v4), total_lines)
                        logging.info(f"Imported {total_imported + len(batch_v4):,}/{total_lines:,} records...")

                except Exception as e:
                    errors += 1
                    if errors <= 10:
                        logging.warning(f"Error processing row: {row}. Error: {e}")

            # Insert remaining records
            if batch_v4:
                cursor.executemany("""
                    INSERT INTO ip_ranges_v4 
                    (network, start_ip, end_ip, country, country_code, 
                     continent, continent_code, asn, as_name, as_domain)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, batch_v4)
                total_imported += len(batch_v4)

            if batch_v6:
                cursor.executemany("""
                    INSERT INTO ip_ranges_v6 
                    (network, start_ip, end_ip, country, country_code, 
                     continent, continent_code, asn, as_name, as_domain)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, batch_v6)
                total_imported += len(batch_v6)

            self.conn.commit()

        # Store metadata
        cursor.execute("INSERT OR REPLACE INTO metadata (key, value) VALUES (?, ?)",
                      ('import_date', datetime.now().isoformat()))
        cursor.execute("INSERT OR REPLACE INTO metadata (key, value) VALUES (?, ?)",
                      ('source_file', os.path.basename(csv_path)))
        cursor.execute("INSERT OR REPLACE INTO metadata (key, value) VALUES (?, ?)",
                      ('total_records', str(total_imported)))
        cursor.execute("INSERT OR REPLACE INTO metadata (key, value) VALUES (?, ?)",
                      ('ipv4_records', str(v4_count)))
        cursor.execute("INSERT OR REPLACE INTO metadata (key, value) VALUES (?, ?)",
                      ('ipv6_records', str(v6_count)))
        cursor.execute("INSERT OR REPLACE INTO metadata (key, value) VALUES (?, ?)",
                      ('errors', str(errors)))

        # Analyze for query optimization
        cursor.execute("ANALYZE ip_ranges_v4")
        cursor.execute("ANALYZE ip_ranges_v6")
        self.conn.commit()

        if progress_callback:
            progress_callback(total_imported, total_lines)

        logging.info(f"Import complete. Total: {total_imported:,} (IPv4: {v4_count:,}, IPv6: {v6_count:,}), Errors: {errors}")
        return total_imported

    def lookup(self, ip: str) -> Optional[Dict[str, Any]]:
        """
        Look up geolocation data for an IP address.

        Args:
            ip: IP address string (IPv4 or IPv6)

        Returns:
            Dictionary with IP info or None if not found
        """
        try:
            addr = ipaddress.ip_address(ip)
        except ValueError:
            logging.warning(f"Invalid IP address: {ip}")
            return None

        cursor = self.conn.cursor()

        if addr.version == 6:
            # IPv6 lookup using hex string comparison
            ip_hex = format(int(addr), '032x')
            cursor.execute("""
                SELECT network, country, country_code, continent, continent_code,
                       asn, as_name, as_domain, start_ip, end_ip
                FROM ip_ranges_v6
                WHERE start_ip <= ? AND end_ip >= ?
                ORDER BY end_ip ASC, start_ip DESC
                LIMIT 1
            """, (ip_hex, ip_hex))
        else:
            # IPv4 lookup using integer comparison
            ip_int = int(addr)
            cursor.execute("""
                SELECT network, country, country_code, continent, continent_code,
                       asn, as_name, as_domain
                FROM ip_ranges_v4
                WHERE start_ip <= ? AND end_ip >= ?
                ORDER BY (end_ip - start_ip) ASC
                LIMIT 1
            """, (ip_int, ip_int))

        row = cursor.fetchone()

        if row:
            return {
                'ip': ip,
                'network': row['network'],
                'country': row['country'] or 'N/A',
                'country_code': row['country_code'] or 'N/A',
                'continent': row['continent'] or 'N/A',
                'continent_code': row['continent_code'] or 'N/A',
                'asn': row['asn'] or 'N/A',
                'as_name': row['as_name'] or 'N/A',
                'as_domain': row['as_domain'] or 'N/A',
                'org': f"{row['asn']} {row['as_name']}".strip() if row['asn'] else 'N/A'
            }

        return None

    def get_stats(self) -> Dict[str, Any]:
        """Get database statistics."""
        cursor = self.conn.cursor()

        cursor.execute("SELECT COUNT(*) FROM ip_ranges_v4")
        v4_records = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM ip_ranges_v6")
        v6_records = cursor.fetchone()[0]

        total_records = v4_records + v6_records

        cursor.execute("SELECT COUNT(DISTINCT country_code) FROM ip_ranges_v4 WHERE country_code != ''")
        v4_countries = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(DISTINCT country_code) FROM ip_ranges_v6 WHERE country_code != ''")
        v6_countries = cursor.fetchone()[0]

        # Get metadata
        cursor.execute("SELECT key, value FROM metadata")
        metadata = {row['key']: row['value'] for row in cursor.fetchall()}

        file_size = os.path.getsize(self.db_path) if os.path.exists(self.db_path) else 0

        return {
            'total_records': total_records,
            'ipv4_records': v4_records,
            'ipv6_records': v6_records,
            'unique_countries': max(v4_countries, v6_countries),
            'database_size_mb': round(file_size / (1024 * 1024), 2),
            'import_date': metadata.get('import_date', 'Unknown'),
            'source_file': metadata.get('source_file', 'Unknown'),
            'errors': int(metadata.get('errors', 0))
        }

    def is_valid(self) -> bool:
        """Check if the database is valid and has data."""
        try:
            cursor = self.conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM ip_ranges_v4")
            v4_count = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM ip_ranges_v6")
            v6_count = cursor.fetchone()[0]
            return (v4_count + v6_count) > 0
        except:
            return False

    def close(self):
        """Close database connection."""
        if self.conn:
            self.conn.close()
            self.conn = None


class ConfigManager:
    """Manages application configuration stored in SQLite."""

    def __init__(self, config_path: str = "ip_checker_config.db"):
        self.config_path = config_path
        self.conn = None
        self._connect()
        self._create_tables()

    def _connect(self):
        """Establish database connection."""
        self.conn = sqlite3.connect(self.config_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row

    def _create_tables(self):
        """Create configuration tables if they don't exist."""
        cursor = self.conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS config (
                key TEXT PRIMARY KEY,
                value TEXT,
                updated_at TEXT
            )
        """)
        self.conn.commit()

    def get(self, key: str, default: str = None) -> Optional[str]:
        """Get a configuration value."""
        cursor = self.conn.cursor()
        cursor.execute("SELECT value FROM config WHERE key = ?", (key,))
        row = cursor.fetchone()
        return row['value'] if row else default

    def set(self, key: str, value: str):
        """Set a configuration value."""
        cursor = self.conn.cursor()
        cursor.execute("""
            INSERT OR REPLACE INTO config (key, value, updated_at)
            VALUES (?, ?, ?)
        """, (key, value, datetime.now().isoformat()))
        self.conn.commit()

    def delete(self, key: str):
        """Delete a configuration value."""
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM config WHERE key = ?", (key,))
        self.conn.commit()

    def get_ip_database_path(self) -> Optional[str]:
        """Get the configured IP database path."""
        path = self.get('ip_database_path')
        if path and os.path.exists(path):
            return path
        return None

    def set_ip_database_path(self, path: str):
        """Set the IP database path."""
        self.set('ip_database_path', path)

    def get_use_local_db(self) -> bool:
        """Get whether to use local database or API."""
        return self.get('use_local_db', 'false').lower() == 'true'

    def set_use_local_db(self, use_local: bool):
        """Set whether to use local database or API."""
        self.set('use_local_db', 'true' if use_local else 'false')

    def close(self):
        """Close database connection."""
        if self.conn:
            self.conn.close()
            self.conn = None


# Global instances
_db_instance: Optional[LocalIPDatabase] = None
_config_instance: Optional[ConfigManager] = None


def get_config() -> ConfigManager:
    """Get or create the config manager instance."""
    global _config_instance
    if _config_instance is None:
        _config_instance = ConfigManager()
    return _config_instance


def get_local_db(db_path: str = None) -> Optional[LocalIPDatabase]:
    """Get or create the local IP database instance."""
    global _db_instance

    if db_path is None:
        config = get_config()
        db_path = config.get_ip_database_path()

    if db_path is None:
        return None

    if _db_instance is None or _db_instance.db_path != db_path:
        if _db_instance:
            _db_instance.close()
        _db_instance = LocalIPDatabase(db_path)

    return _db_instance


def fetch_ip_data_local(ip: str) -> Optional[Dict[str, Any]]:
    """
    Look up IP using local database.
    Returns data formatted to match ipinfo.io API response.
    """
    db = get_local_db()
    if db is None:
        return None

    result = db.lookup(ip)

    if result:
        return {
            'ip': result['ip'],
            'org': result['org'],
            'hostname': 'N/A',
            'region': result['continent'],
            'country': result['country_code'],
            'country_name': result['country'],
            'continent': result['continent'],
            'asn': result['asn'],
            'as_name': result['as_name']
        }

    return None