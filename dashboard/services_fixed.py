# dashboard/services_fixed.py
"""
سرویس SharePoint با Query های اصلاح شده
"""
from typing import List, Dict, Any
from .db_connection import SharePointDatabase
import logging

logger = logging.getLogger(__name__)


class SharePointServiceFixed:
    """سرویس با query های سازگار"""

    def __init__(self):
        self.db = SharePointDatabase()

    def get_all_lists_simple(self):
        """دریافت لیست‌ها - ساده"""
        # ابتدا ستون‌ها رو پیدا می‌کنیم
        query_columns = """
        SELECT COLUMN_NAME
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = 'AllLists'
        """

        try:
            columns = self.db.execute_query(query_columns)
            column_names = [col['COLUMN_NAME'] for col in columns]

            # استفاده از ستون‌های موجود
            query = f"""
            SELECT TOP 20 *
            FROM AllLists
            """

            results = self.db.execute_query(query)

            return {
                'columns': column_names,
                'data': results,
                'count': len(results)
            }

        except Exception as e:
            logger.error(f"Error: {e}")
            return {
                'error': str(e),
                'columns': [],
                'data': []
            }

    def get_user_data_simple(self):
        """دریافت داده‌های کاربری - ساده"""
        # ابتدا بررسی ستون‌ها
        query_check = """
        SELECT TOP 1 *
        FROM AllUserData
        """

        try:
            result = self.db.execute_query(query_check)
            if result:
                # لیست ستون‌های موجود
                columns = list(result[0].keys())

                # پیدا کردن ستون‌های متنی
                text_columns = [col for col in columns if 'nvarchar' in col.lower() or 'ntext' in col.lower()]

                if text_columns:
                    # استفاده از اولین ستون متنی
                    text_col = text_columns[0] if text_columns else '*'

                    query = f"""
                    SELECT TOP 10 *
                    FROM AllUserData
                    WHERE {text_col} IS NOT NULL
                    """

                    data = self.db.execute_query(query)

                    return {
                        'columns': columns,
                        'text_columns': text_columns,
                        'data': data
                    }

            return {'error': 'No data found'}

        except Exception as e:
            logger.error(f"Error: {e}")
            return {'error': str(e)}

    def get_users_simple(self):
        """دریافت کاربران - ساده"""
        query = """
        SELECT TOP 20 *
        FROM UserInfo
        """

        try:
            results = self.db.execute_query(query)

            if results:
                # پیدا کردن ستون‌های مهم
                sample = results[0]
                columns = list(sample.keys())

                # شناسایی ستون‌های احتمالی
                name_cols = [col for col in columns if
                             any(x in col.lower() for x in ['name', 'title', 'login', 'email'])]

                return {
                    'users': results,
                    'columns': columns,
                    'important_columns': name_cols,
                    'count': len(results)
                }

            return {'users': [], 'columns': []}

        except Exception as e:
            logger.error(f"Error: {e}")
            return {'error': str(e)}

    def test_all_tables(self):
        """تست همه جداول"""
        tables = ['AllLists', 'AllUserData', 'UserInfo']
        results = {}

        for table in tables:
            query = f"SELECT COUNT(*) as cnt FROM {table}"
            try:
                result = self.db.execute_query(query)
                count = result[0]['cnt'] if result else 0

                # دریافت نمونه
                sample_query = f"SELECT TOP 1 * FROM {table}"
                sample = self.db.execute_query(sample_query)

                results[table] = {
                    'count': count,
                    'columns': list(sample[0].keys()) if sample else [],
                    'accessible': True
                }
            except Exception as e:
                results[table] = {
                    'count': 0,
                    'error': str(e),
                    'accessible': False
                }

        return results