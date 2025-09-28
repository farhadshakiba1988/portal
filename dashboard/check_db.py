"""
بررسی ساختار دیتابیس SharePoint
برای شناسایی نام صحیح ستون‌ها
"""
import pyodbc
from django.conf import settings
import json


class DatabaseInspector:
    def __init__(self):
        self.connection_string = (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={settings.SHAREPOINT_DB['SERVER']};"
            f"DATABASE={settings.SHAREPOINT_DB['DATABASE']};"
            f"UID={settings.SHAREPOINT_DB['USERNAME']};"
            f"PWD={settings.SHAREPOINT_DB['PASSWORD']};"
            f"TrustServerCertificate=yes;"
        )

    def get_tables(self):
        """لیست تمام جداول"""
        query = """
        SELECT TABLE_NAME 
        FROM INFORMATION_SCHEMA.TABLES 
        WHERE TABLE_TYPE = 'BASE TABLE'
        ORDER BY TABLE_NAME
        """

        conn = pyodbc.connect(self.connection_string)
        cursor = conn.cursor()
        cursor.execute(query)

        tables = [row[0] for row in cursor.fetchall()]
        conn.close()

        return tables

    def get_columns(self, table_name):
        """ستون‌های یک جدول"""
        query = """
        SELECT 
            COLUMN_NAME,
            DATA_TYPE,
            CHARACTER_MAXIMUM_LENGTH,
            IS_NULLABLE
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = ?
        ORDER BY ORDINAL_POSITION
        """

        conn = pyodbc.connect(self.connection_string)
        cursor = conn.cursor()
        cursor.execute(query, table_name)

        columns = []
        for row in cursor.fetchall():
            columns.append({
                'name': row[0],
                'type': row[1],
                'length': row[2],
                'nullable': row[3]
            })

        conn.close()
        return columns

    def find_important_tables(self):
        """پیدا کردن جداول مهم SharePoint"""
        important_patterns = [
            'List', 'Doc', 'User', 'Web', 'Site',
            'Content', 'Item', 'Announcement', 'Task'
        ]

        all_tables = self.get_tables()
        important = {}

        for table in all_tables:
            for pattern in important_patterns:
                if pattern.lower() in table.lower():
                    columns = self.get_columns(table)
                    important[table] = columns
                    break

        return important

    def check_specific_table(self, table_name):
        """بررسی دقیق یک جدول خاص"""
        query = f"SELECT TOP 5 * FROM {table_name}"

        try:
            conn = pyodbc.connect(self.connection_string)
            cursor = conn.cursor()
            cursor.execute(query)

            # نام ستون‌ها
            columns = [column[0] for column in cursor.description]

            # چند رکورد نمونه
            rows = []
            for row in cursor.fetchall():
                rows.append(dict(zip(columns, row)))

            conn.close()

            return {
                'columns': columns,
                'sample_data': rows
            }
        except Exception as e:
            return {'error': str(e)}

    def find_announcements_table(self):
        """پیدا کردن جدول اطلاعیه‌ها"""
        possible_queries = [
            # روش 1: جدول مستقیم
            "SELECT TOP 1 * FROM Announcements",

            # روش 2: از طریق Lists
            """
            SELECT TOP 1 * FROM Lists 
            WHERE Title LIKE '%Announcement%' 
               OR Title LIKE '%اطلاعیه%'
            """,

            # روش 3: AllLists pattern
            "SELECT TOP 1 * FROM AllLists WHERE tp_Title LIKE '%Announcement%'",

            # روش 4: UserData pattern
            "SELECT TOP 1 * FROM UserData",

            # روش 5: AllUserData pattern
            "SELECT TOP 1 * FROM AllUserData",
        ]

        conn = pyodbc.connect(self.connection_string)
        cursor = conn.cursor()

        for query in possible_queries:
            try:
                cursor.execute(query)
                columns = [column[0] for column in cursor.description]
                print(f"✅ Query worked: {query[:50]}...")
                print(f"   Columns: {columns[:5]}...")
                return columns
            except:
                continue

        conn.close()
        return None


# اجرای بررسی
if __name__ == "__main__":
    inspector = DatabaseInspector()

    print("=" * 60)
    print("🔍 بررسی ساختار دیتابیس SharePoint")
    print("=" * 60)

    # 1. جداول موجود
    print("\n📋 جداول موجود:")
    tables = inspector.get_tables()
    for i, table in enumerate(tables[:20], 1):  # فقط 20 تای اول
        print(f"   {i}. {table}")

    # 2. بررسی جداول مهم
    print("\n🎯 جداول احتمالی مهم:")
    important = inspector.find_important_tables()
    for table_name, columns in list(important.items())[:5]:
        print(f"\n   📌 {table_name}:")
        for col in columns[:10]:  # فقط 10 ستون اول
            print(f"      - {col['name']} ({col['type']})")

    # 3. جستجوی جدول اطلاعیه‌ها
    print("\n🔎 جستجوی جدول اطلاعیه‌ها:")
    inspector.find_announcements_table()