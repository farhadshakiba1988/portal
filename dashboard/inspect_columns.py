"""
بررسی دقیق ستون‌های جداول SharePoint
"""
import pyodbc
from django.conf import settings
import json


def inspect_sharepoint_tables():
    """بررسی ساختار جداول SharePoint"""

    connection_string = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={settings.SHAREPOINT_DB['SERVER']};"
        f"DATABASE={settings.SHAREPOINT_DB['DATABASE']};"
        f"UID={settings.SHAREPOINT_DB['USERNAME']};"
        f"PWD={settings.SHAREPOINT_DB['PASSWORD']};"
        f"TrustServerCertificate=yes;"
    )

    conn = pyodbc.connect(connection_string)
    cursor = conn.cursor()

    important_tables = ['AllLists', 'AllUserData', 'UserInfo']

    results = {}

    for table in important_tables:
        print(f"\n{'=' * 60}")
        print(f"📋 جدول: {table}")
        print('=' * 60)

        # دریافت ستون‌ها
        query = f"""
        SELECT 
            COLUMN_NAME,
            DATA_TYPE,
            CHARACTER_MAXIMUM_LENGTH,
            IS_NULLABLE
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = '{table}'
        ORDER BY ORDINAL_POSITION
        """

        cursor.execute(query)
        columns = []

        print("\n🔹 ستون‌ها:")
        for row in cursor.fetchall():
            col_name = row[0]
            col_type = row[1]
            col_length = row[2] if row[2] else ''
            nullable = row[3]

            columns.append(col_name)
            print(f"   - {col_name:<30} {col_type:<15} {col_length:<10} {nullable}")

        results[table] = columns

        # نمایش چند رکورد نمونه
        print(f"\n🔹 نمونه داده‌ها از {table}:")
        try:
            sample_query = f"SELECT TOP 2 * FROM {table}"
            cursor.execute(sample_query)

            # نام ستون‌ها
            col_names = [column[0] for column in cursor.description]

            # داده‌ها
            for row in cursor.fetchall():
                print("\n   رکورد:")
                for i, value in enumerate(row):
                    if value is not None:
                        value_str = str(value)[:50]  # محدود کردن طول
                        if len(str(value)) > 50:
                            value_str += '...'
                        print(f"      {col_names[i]}: {value_str}")
        except Exception as e:
            print(f"   خطا در دریافت نمونه: {e}")

    # بررسی خاص برای لیست‌ها
    print(f"\n{'=' * 60}")
    print("🔍 جستجوی لیست‌های SharePoint")
    print('=' * 60)

    # پیدا کردن لیست Announcements
    try:
        # روش 1: جستجو در AllLists
        list_query = """
        SELECT TOP 10 *
        FROM AllLists
        """
        cursor.execute(list_query)

        print("\n📌 لیست‌های موجود در AllLists:")
        col_names = [column[0] for column in cursor.description]

        # پیدا کردن ستون Title یا مشابه
        title_col = None
        for col in col_names:
            if 'title' in col.lower() or 'name' in col.lower():
                title_col = col
                break

        if title_col:
            print(f"   ستون عنوان: {title_col}")

            # دریافت لیست‌ها
            cursor.execute(f"SELECT TOP 20 {title_col} FROM AllLists WHERE {title_col} IS NOT NULL")
            print("\n   لیست‌های موجود:")
            for row in cursor.fetchall():
                print(f"      - {row[0]}")

    except Exception as e:
        print(f"خطا: {e}")

    conn.close()
    return results


def find_announcement_structure():
    """پیدا کردن ساختار اطلاعیه‌ها"""

    connection_string = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={settings.SHAREPOINT_DB['SERVER']};"
        f"DATABASE={settings.SHAREPOINT_DB['DATABASE']};"
        f"UID={settings.SHAREPOINT_DB['USERNAME']};"
        f"PWD={settings.SHAREPOINT_DB['PASSWORD']};"
        f"TrustServerCertificate=yes;"
    )

    conn = pyodbc.connect(connection_string)
    cursor = conn.cursor()

    print(f"\n{'=' * 60}")
    print("🔎 جستجوی ساختار داده‌های لیست‌ها")
    print('=' * 60)

    # بررسی AllUserData
    try:
        query = """
        SELECT TOP 5 *
        FROM AllUserData
        """
        cursor.execute(query)

        columns = [column[0] for column in cursor.description]

        print("\n📊 ستون‌های AllUserData که احتمالاً مفید هستند:")

        # دسته‌بندی ستون‌ها
        id_cols = [c for c in columns if 'id' in c.lower()]
        text_cols = [c for c in columns if any(x in c.lower() for x in ['nvarchar', 'ntext', 'varchar', 'text'])]
        date_cols = [c for c in columns if any(x in c.lower() for x in ['date', 'time', 'created', 'modified'])]
        int_cols = [c for c in columns if 'int' in c.lower()]

        print(f"\n   🔑 ستون‌های ID: {id_cols[:10]}")
        print(f"\n   📝 ستون‌های متنی: {text_cols[:15]}")
        print(f"\n   📅 ستون‌های تاریخ: {date_cols[:10]}")
        print(f"\n   🔢 ستون‌های عددی: {int_cols[:10]}")

        # نمونه داده
        print("\n📋 نمونه داده از AllUserData:")
        cursor.execute(query)

        for i, row in enumerate(cursor.fetchall(), 1):
            print(f"\n   رکورد {i}:")
            for j, col in enumerate(columns):
                if row[j] is not None and str(row[j]).strip():
                    # فقط مقادیر غیر خالی
                    if any(x in col.lower() for x in ['nvarchar', 'ntext']):
                        value = str(row[j])[:100]
                        if value:
                            print(f"      {col}: {value}")

    except Exception as e:
        print(f"خطا: {e}")

    conn.close()


if __name__ == "__main__":
    print("🚀 شروع بررسی جداول SharePoint...")
    results = inspect_sharepoint_tables()
    print("\n" + "=" * 60)
    print("✅ بررسی کامل شد")
    print("\nجداول و ستون‌های یافت شده:")
    for table, columns in results.items():
        print(f"\n{table}: {len(columns)} ستون")
        print(f"   نمونه ستون‌ها: {columns[:5]}...")

    # بررسی اطلاعیه‌ها
    find_announcement_structure()