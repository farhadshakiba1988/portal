# dashboard/services_simple.py
"""
سرویس ساده SharePoint - نسخه سازگار
"""
from typing import List, Dict, Any
from .db_connection import SharePointDatabase
import logging

logger = logging.getLogger(__name__)


class SharePointServiceSimple:
    """سرویس ساده برای شروع"""

    def __init__(self):
        self.db = SharePointDatabase()

    def test_query(self):
        """تست ساده"""
        query = "SELECT TOP 10 * FROM AllLists"
        try:
            results = self.db.execute_query(query)
            logger.info(f"Test query returned {len(results)} results")
            return results
        except Exception as e:
            logger.error(f"Test query failed: {e}")
            return []

    def get_all_tables(self):
        """لیست جداول"""
        query = """
        SELECT TABLE_NAME, TABLE_TYPE
        FROM INFORMATION_SCHEMA.TABLES
        WHERE TABLE_TYPE = 'BASE TABLE'
        ORDER BY TABLE_NAME
        """
        try:
            return self.db.execute_query(query)
        except Exception as e:
            logger.error(f"Error: {e}")
            return []

    def get_table_columns(self, table_name):
        """ستون‌های یک جدول"""
        query = """
        SELECT COLUMN_NAME, DATA_TYPE
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = ?
        ORDER BY ORDINAL_POSITION
        """
        try:
            return self.db.execute_query(query, (table_name,))
        except Exception as e:
            logger.error(f"Error: {e}")
            return []

    def get_simple_data(self):
        """داده‌های ساده برای تست"""
        # این یک query خیلی ساده است که احتمالاً کار می‌کند
        query = """
        SELECT TOP 10 
            Id,
            Title,
            Created,
            Modified
        FROM AllLists
        WHERE Title IS NOT NULL
        """

        try:
            results = self.db.execute_query(query)
            return {
                'lists': results,
                'count': len(results)
            }
        except Exception as e:
            logger.error(f"Simple query failed: {e}")

            # اگر AllLists هم نبود، یک query ساده‌تر
            try:
                query2 = "SELECT name FROM sys.tables WHERE type = 'U'"
                tables = self.db.execute_query(query2)
                return {
                    'tables': tables,
                    'message': 'فقط لیست جداول موجود است'
                }
            except:
                return {
                    'error': 'عدم دسترسی به دیتابیس',
                    'message': str(e)
                }

    def get_statistics_simple(self):
        """آمار ساده"""
        stats = {}

        # تعداد جداول
        try:
            query = "SELECT COUNT(*) as cnt FROM sys.tables WHERE type = 'U'"
            result = self.db.execute_query(query)
            stats['total_tables'] = result[0]['cnt'] if result else 0
        except:
            stats['total_tables'] = 0

        # حجم دیتابیس
        try:
            query = """
            SELECT 
                DB_NAME() as database_name,
                CAST(SUM(size) * 8 / 1024.0 AS DECIMAL(10,2)) as size_mb
            FROM sys.master_files
            WHERE DB_NAME(database_id) = DB_NAME()
            """
            result = self.db.execute_query(query)
            stats['database_size_mb'] = result[0]['size_mb'] if result else 0
        except:
            stats['database_size_mb'] = 0

        return stats