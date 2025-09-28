"""
SharePoint Database Connection Manager
کلاس اصلی برای مدیریت اتصال به دیتابیس SharePoint
"""
import pyodbc
import logging
from typing import List, Dict, Any, Optional, Union
from contextlib import contextmanager
from django.conf import settings
from django.core.cache import cache
import json
import hashlib

logger = logging.getLogger(__name__)


class SharePointDatabase:
    """
    مدیریت اتصال به دیتابیس SharePoint SQL Server
    """

    def __init__(self):
        """
        مقداردهی اولیه تنظیمات اتصال
        """
        self.server = settings.SHAREPOINT_DB['SERVER']
        self.database = settings.SHAREPOINT_DB['DATABASE']
        self.username = settings.SHAREPOINT_DB['USERNAME']
        self.password = settings.SHAREPOINT_DB['PASSWORD']
        self.driver = settings.SHAREPOINT_DB['DRIVER']

        self.connection_string = (
            f"DRIVER={{{self.driver}}};"
            f"SERVER={self.server};"
            f"DATABASE={self.database};"
            f"UID={self.username};"
            f"PWD={self.password};"
            f"TrustServerCertificate=yes;"
            f"Connection Timeout=30;"
        )

        logger.info(f"SharePoint DB initialized for server: {self.server}")

    @contextmanager
    def get_connection(self):
        """
        Context manager برای مدیریت اتصال
        """
        conn = None
        try:
            conn = pyodbc.connect(self.connection_string)
            yield conn
        except pyodbc.Error as e:
            logger.error(f"Database connection error: {e}")
            raise
        finally:
            if conn:
                conn.close()

    def execute_query(
            self,
            query: str,
            params: Optional[Union[tuple, list]] = None,
            cache_key: Optional[str] = None,
            cache_timeout: int = 300
    ) -> List[Dict[str, Any]]:
        """
        اجرای کوئری و بازگرداندن نتایج به صورت لیست از دیکشنری

        Args:
            query: SQL query string
            params: پارامترهای کوئری
            cache_key: کلید کش (اختیاری)
            cache_timeout: مدت زمان کش به ثانیه

        Returns:
            لیست از دیکشنری‌های حاوی نتایج
        """
        # بررسی کش
        if cache_key:
            cached_data = cache.get(cache_key)
            if cached_data:
                logger.debug(f"Cache hit for key: {cache_key}")
                return cached_data

        results = []

        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()

                if params:
                    cursor.execute(query, params)
                else:
                    cursor.execute(query)

                # تبدیل نتایج به دیکشنری
                columns = [column[0] for column in cursor.description] if cursor.description else []

                for row in cursor.fetchall():
                    results.append(dict(zip(columns, row)))

                logger.info(f"Query executed successfully, returned {len(results)} rows")

                # ذخیره در کش
                if cache_key and results:
                    cache.set(cache_key, results, cache_timeout)
                    logger.debug(f"Cached results for key: {cache_key}")

        except Exception as e:
            logger.error(f"Error executing query: {e}")
            logger.error(f"Query: {query[:200]}...")  # نمایش بخشی از کوئری
            raise

        return results

    def execute_scalar(self, query: str, params: Optional[Union[tuple, list]] = None) -> Any:
        """
        اجرای کوئری و بازگرداندن یک مقدار
        """
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()

                if params:
                    cursor.execute(query, params)
                else:
                    cursor.execute(query)

                row = cursor.fetchone()
                return row[0] if row else None

        except Exception as e:
            logger.error(f"Error executing scalar query: {e}")
            raise

    def test_connection(self) -> bool:
        """
        تست اتصال به دیتابیس
        """
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT 1")
                result = cursor.fetchone()
                logger.info("Database connection test successful")
                return result[0] == 1
        except Exception as e:
            logger.error(f"Database connection test failed: {e}")
            return False

    def get_database_info(self) -> Dict[str, Any]:
        """
        دریافت اطلاعات دیتابیس
        """
        query = """
        SELECT 
            DB_NAME() as DatabaseName,
            @@VERSION as ServerVersion,
            @@SERVERNAME as ServerName,
            SUSER_NAME() as CurrentUser,
            (SELECT COUNT(*) FROM sys.tables) as TableCount
        """

        results = self.execute_query(query)
        return results[0] if results else {}