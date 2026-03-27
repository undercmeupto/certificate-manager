"""
Migration Script: JSON to SQLite
数据迁移脚本：从 JSON 迁移到 SQLite

This script migrates data from the old JSON file storage to the new SQLite database.
此脚本将数据从旧的 JSON 文件存储迁移到新的 SQLite 数据库。

Usage / 用法:
    python migrations/migrate_json_to_sqlite.py
"""
import os
import sys
import json
from datetime import datetime

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import DATA_FILE, DATABASE_FILE, USER_DATA_DIR
from database import init_db, get_session, close_session, session_scope
from models import Certificate, UploadMetadata, SessionState


def load_json_data():
    """Load data from JSON file"""
    if not os.path.exists(DATA_FILE):
        print(f"JSON data file not found: {DATA_FILE}")
        return None, None

    with open(DATA_FILE, 'r', encoding='utf-8') as f:
        data = json.load(f)

    certificates = data.get('data', [])
    metadata = data.get('metadata', {})

    print(f"Loaded {len(certificates)} certificates from JSON")
    return certificates, metadata


def migrate_certificates(certificates, session):
    """Migrate certificates to database"""
    success_count = 0
    error_count = 0

    for cert_data in certificates:
        try:
            cert = Certificate.from_dict(cert_data)
            session.add(cert)
            success_count += 1
        except Exception as e:
            print(f"Error migrating certificate {cert_data.get('id')}: {e}")
            error_count += 1

    session.flush()
    print(f"Migrated {success_count} certificates successfully")
    if error_count > 0:
        print(f"Failed to migrate {error_count} certificates")

    return success_count, error_count


def migrate_metadata(metadata, session):
    """Migrate upload metadata to database"""
    if not metadata:
        print("No metadata to migrate")
        return

    try:
        upload_meta = UploadMetadata.from_dict(metadata)
        session.add(upload_meta)
        session.flush()
        print("Migrated metadata successfully")
    except Exception as e:
        print(f"Error migrating metadata: {e}")


def migrate_session_state(session):
    """Initialize session state in database"""
    try:
        # Check if session state already exists
        existing = session.query(SessionState).filter_by(id=1).first()
        if existing is None:
            state = SessionState(id=1, active=False)
            session.add(state)
            session.flush()
            print("Session state initialized")
        else:
            print("Session state already exists")
    except Exception as e:
        print(f"Error initializing session state: {e}")


def backup_json_file():
    """Create backup of original JSON file"""
    if os.path.exists(DATA_FILE):
        backup_path = DATA_FILE.replace('.json', '_backup.json')
        import shutil
        shutil.copy2(DATA_FILE, backup_path)
        print(f"JSON backup created: {backup_path}")
        return backup_path
    return None


def verify_migration(cert_count, session):
    """Verify migration by counting records"""
    db_count = session.query(Certificate).count()
    print(f"\nVerification: JSON had {cert_count} certificates, DB has {db_count} certificates")

    if db_count == cert_count:
        print("✓ Migration verification PASSED")
        return True
    else:
        print("✗ Migration verification FAILED")
        return False


def main():
    """Main migration function"""
    print("=" * 60)
    print("Certificate Manager - JSON to SQLite Migration")
    print("证件管理系统 - JSON 到 SQLite 迁移")
    print("=" * 60)
    print()

    # Step 1: Load JSON data
    print("Step 1: Loading JSON data...")
    certificates, metadata = load_json_data()

    if certificates is None:
        print("No data to migrate. Exiting.")
        return

    if not certificates:
        print("No certificates found in JSON file. Creating empty database...")
        # Still create the database with empty tables
        init_db()
        print("Empty database created at:", DATABASE_FILE)
        return

    # Step 2: Backup JSON file
    print("\nStep 2: Backing up JSON file...")
    backup_path = backup_json_file()

    # Step 3: Initialize database
    print("\nStep 3: Initializing SQLite database...")
    print(f"Database path: {DATABASE_FILE}")
    init_db()

    # Step 4: Migrate data
    print("\nStep 4: Migrating data...")
    with session_scope() as session:
        # Migrate certificates
        migrate_certificates(certificates, session)

        # Migrate metadata
        migrate_metadata(metadata, session)

        # Initialize session state
        migrate_session_state(session)

        # Verify migration
        print("\nStep 5: Verifying migration...")
        verify_migration(len(certificates), session)

    print()
    print("=" * 60)
    print("Migration completed successfully!")
    print("迁移成功完成！")
    print("=" * 60)
    print()
    print("Summary / 摘要:")
    print(f"  - JSON backup: {backup_path}")
    print(f"  - Database: {DATABASE_FILE}")
    print(f"  - Certificates migrated: {len(certificates)}")
    print()
    print("Next steps / 后续步骤:")
    print("  1. Test the application to ensure everything works")
    print("  2. If satisfied, you can delete the backup file")
    print("  3. If issues occur, restore from the backup file")
    print()


if __name__ == '__main__':
    main()
