"""
SharePoint Services
سرویس‌های مختلف برای دریافت داده‌ها از SharePoint
"""
from typing import List, Dict, Any, Optional
from datetime import datetime, timedelta
import re
import html
from .db_connection import SharePointDatabase
import logging

logger = logging.getLogger(__name__)


class SharePointService:
    """
    سرویس اصلی برای تعامل با داده‌های SharePoint
    """

    def __init__(self):
        self.db = SharePointDatabase()

    def get_announcements(self, limit: int = 10) -> List[Dict[str, Any]]:
        """
        دریافت آخرین اطلاعیه‌ها
        """
        query = """
        SELECT TOP (?)
            UD.tp_ID as Id,
            UD.nvarchar1 as Title,
            UD.ntext2 as Body,
            UD.tp_Created as Created,
            UD.tp_Modified as Modified,
            U.tp_Title as AuthorName,
            U.tp_Email as AuthorEmail,
            L.tp_Title as ListName
        FROM AllUserData UD WITH (NOLOCK)
        INNER JOIN AllLists L WITH (NOLOCK) ON UD.tp_ListId = L.tp_ID
        LEFT JOIN UserInfo U WITH (NOLOCK) ON UD.tp_Author = U.tp_ID
        WHERE L.tp_Title = 'Announcements'
            AND UD.tp_DeleteTransactionId = 0x
            AND UD.tp_IsCurrentVersion = 1
            AND UD.tp_RowOrdinal = 0
        ORDER BY UD.tp_Created DESC
        """

        results = self.db.execute_query(
            query,
            params=(limit,),
            cache_key=f'announcements_{limit}',
            cache_timeout=60  # کش کوتاه برای اطلاعیه‌ها
        )

        # پردازش و تمیز کردن داده‌ها
        for item in results:
            # تمیز کردن HTML از Body
            if item.get('Body'):
                clean_body = self._clean_html(item['Body'])
                item['Body'] = clean_body[:500]  # محدود کردن طول
                item['BodyFull'] = clean_body

            # فرمت تاریخ
            if item.get('Created'):
                item['CreatedFormatted'] = self._format_date(item['Created'])
                item['CreatedRelative'] = self._get_relative_time(item['Created'])

        logger.info(f"Fetched {len(results)} announcements")
        return results

    def get_documents(self, limit: int = 20, extensions: Optional[List[str]] = None) -> List[Dict[str, Any]]:
        """
        دریافت آخرین اسناد
        """
        if extensions is None:
            extensions = ['pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx']

        extensions_str = "','".join(extensions)

        query = f"""
        SELECT TOP (?)
            D.Id,
            D.LeafName as FileName,
            D.DirName as FolderPath,
            D.Extension,
            D.TimeCreated as Created,
            D.TimeLastModified as Modified,
            D.Size as SizeBytes,
            CAST(D.Size / 1024.0 as DECIMAL(10,2)) as SizeKB,
            CAST(D.Size / 1048576.0 as DECIMAL(10,2)) as SizeMB,
            D.CheckoutUserId,
            U.tp_Title as AuthorName,
            CU.tp_Title as CheckedOutTo
        FROM AllDocs D WITH (NOLOCK)
        LEFT JOIN UserInfo U WITH (NOLOCK) ON D.AuthorId = U.tp_ID
        LEFT JOIN UserInfo CU WITH (NOLOCK) ON D.CheckoutUserId = CU.tp_ID
        WHERE D.Type = 0  -- فقط فایل‌ها
            AND D.DeleteTransactionId = 0x
            AND D.Extension IN ('{extensions_str}')
            AND D.DirName NOT LIKE '%_vti_%'  -- فیلتر فولدرهای سیستمی
        ORDER BY D.TimeLastModified DESC
        """

        results = self.db.execute_query(
            query,
            params=(limit,),
            cache_key=f'documents_{limit}_{extensions_str}',
            cache_timeout=180
        )

        # افزودن اطلاعات اضافی
        for doc in results:
            doc['Icon'] = self._get_file_icon(doc.get('Extension', ''))
            doc['SizeFormatted'] = self._format_file_size(doc.get('SizeBytes', 0))
            doc['ModifiedFormatted'] = self._format_date(doc.get('Modified'))
            doc['IsCheckedOut'] = doc.get('CheckoutUserId') is not None
            doc['DownloadUrl'] = self._build_document_url(doc.get('FolderPath'), doc.get('FileName'))

        logger.info(f"Fetched {len(results)} documents")
        return results

    def get_tasks(self, assigned_to: Optional[str] = None, status: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        دریافت وظایف
        """
        query = """
        SELECT 
            UD.tp_ID as Id,
            UD.nvarchar1 as Title,
            UD.nvarchar3 as Status,
            UD.nvarchar4 as Priority,
            UD.datetime1 as DueDate,
            UD.datetime2 as StartDate,
            UD.float1 as PercentComplete,
            UD.ntext2 as Description,
            UD.tp_Created as Created,
            U.tp_Title as AssignedTo,
            U.tp_Email as AssignedToEmail,
            C.tp_Title as CreatedBy
        FROM AllUserData UD WITH (NOLOCK)
        INNER JOIN AllLists L WITH (NOLOCK) ON UD.tp_ListId = L.tp_ID
        LEFT JOIN UserInfo U WITH (NOLOCK) ON UD.int1 = U.tp_ID
        LEFT JOIN UserInfo C WITH (NOLOCK) ON UD.tp_Author = C.tp_ID
        WHERE L.tp_Title = 'Tasks'
            AND UD.tp_DeleteTransactionId = 0x
            AND UD.tp_IsCurrentVersion = 1
            AND (? IS NULL OR U.tp_Title = ?)
            AND (? IS NULL OR UD.nvarchar3 = ?)
        ORDER BY 
            CASE 
                WHEN UD.nvarchar4 = 'High' THEN 1
                WHEN UD.nvarchar4 = 'Normal' THEN 2
                WHEN UD.nvarchar4 = 'Low' THEN 3
                ELSE 4
            END,
            UD.datetime1 ASC
        """

        results = self.db.execute_query(
            query,
            params=(assigned_to, assigned_to, status, status),
            cache_key=f'tasks_{assigned_to}_{status}' if assigned_to or status else 'tasks_all'
        )

        # پردازش وظایف
        for task in results:
            # محاسبه وضعیت
            task['IsOverdue'] = False
            if task.get('DueDate') and task.get('Status') != 'Completed':
                task['IsOverdue'] = datetime.now() > task['DueDate']

            # فرمت تاریخ
            task['DueDateFormatted'] = self._format_date(task.get('DueDate'))
            task['StartDateFormatted'] = self._format_date(task.get('StartDate'))

            # رنگ‌بندی اولویت
            priority_colors = {
                'High': 'danger',
                'Normal': 'warning',
                'Low': 'info'
            }
            task['PriorityColor'] = priority_colors.get(task.get('Priority'), 'secondary')

            # درصد تکمیل
            task['PercentComplete'] = int(task.get('PercentComplete', 0) * 100) if task.get('PercentComplete') else 0

        logger.info(f"Fetched {len(results)} tasks")
        return results

    def get_calendar_events(self, days_ahead: int = 30, days_behind: int = 0) -> List[Dict[str, Any]]:
        """
        دریافت رویدادهای تقویم
        """
        query = """
        SELECT 
            UD.tp_ID as Id,
            UD.nvarchar1 as Title,
            UD.datetime1 as StartTime,
            UD.datetime2 as EndTime,
            UD.nvarchar3 as Location,
            UD.ntext2 as Description,
            UD.bit1 as AllDayEvent,
            UD.nvarchar5 as Category,
            UD.tp_Created as Created,
            U.tp_Title as Organizer
        FROM AllUserData UD WITH (NOLOCK)
        INNER JOIN AllLists L WITH (NOLOCK) ON UD.tp_ListId = L.tp_ID
        LEFT JOIN UserInfo U WITH (NOLOCK) ON UD.tp_Author = U.tp_ID
        WHERE L.tp_Title IN ('Calendar', 'Events', 'تقویم')
            AND UD.tp_DeleteTransactionId = 0x
            AND UD.datetime1 >= DATEADD(day, -?, GETDATE())
            AND UD.datetime1 <= DATEADD(day, ?, GETDATE())
        ORDER BY UD.datetime1 ASC
        """

        results = self.db.execute_query(
            query,
            params=(days_behind, days_ahead),
            cache_key=f'events_{days_ahead}_{days_behind}'
        )

        # پردازش رویدادها
        for event in results:
            event['StartTimeFormatted'] = self._format_datetime(event.get('StartTime'))
            event['EndTimeFormatted'] = self._format_datetime(event.get('EndTime'))
            event['Duration'] = self._calculate_duration(event.get('StartTime'), event.get('EndTime'))

            # تعیین وضعیت رویداد
            now = datetime.now()
            if event.get('StartTime'):
                if event['StartTime'] > now:
                    event['Status'] = 'upcoming'
                elif event.get('EndTime') and event['EndTime'] < now:
                    event['Status'] = 'past'
                else:
                    event['Status'] = 'ongoing'

        logger.info(f"Fetched {len(results)} calendar events")
        return results

    def get_lists_info(self) -> List[Dict[str, Any]]:
        """
        دریافت اطلاعات لیست‌های SharePoint
        """
        query = """
        SELECT 
            L.tp_ID as Id,
            L.tp_Title as Title,
            L.tp_Description as Description,
            L.tp_ItemCount as ItemCount,
            L.tp_Created as Created,
            L.tp_Modified as Modified,
            L.tp_ServerTemplate as Template,
            L.tp_BaseType as BaseType,
            L.tp_Hidden as IsHidden,
            W.FullUrl as WebUrl
        FROM AllLists L WITH (NOLOCK)
        INNER JOIN Webs W WITH (NOLOCK) ON L.tp_WebId = W.Id
        WHERE L.tp_Hidden = 0
            AND L.tp_ItemCount > 0
        ORDER BY L.tp_ItemCount DESC
        """

        results = self.db.execute_query(
            query,
            cache_key='lists_info',
            cache_timeout=600
        )

        # تعیین نوع لیست
        template_names = {
            100: 'لیست سفارشی',
            101: 'کتابخانه اسناد',
            102: 'نظرسنجی',
            103: 'لینک‌ها',
            104: 'اطلاعیه‌ها',
            105: 'مخاطبین',
            106: 'رویدادها',
            107: 'وظایف',
            108: 'بحث‌ها',
            109: 'کتابخانه تصاویر'
        }

        for lst in results:
            lst['TemplateType'] = template_names.get(lst.get('Template'), 'نامشخص')
            lst['Icon'] = self._get_list_icon(lst.get('Template'))
            lst['CreatedFormatted'] = self._format_date(lst.get('Created'))

        return results

    def get_users(self, active_only: bool = True) -> List[Dict[str, Any]]:
        """
        دریافت لیست کاربران
        """
        query = """
        SELECT 
            tp_ID as Id,
            tp_Title as FullName,
            tp_Email as Email,
            tp_Login as LoginName,
            tp_Created as CreatedDate,
            tp_IsActive as IsActive,
            tp_Deleted as IsDeleted,
            tp_SiteAdmin as IsSiteAdmin
        FROM UserInfo WITH (NOLOCK)
        WHERE tp_Deleted = 0
            AND tp_Email IS NOT NULL
            AND tp_Email != ''
            {}
        ORDER BY tp_Title
        """.format("AND tp_IsActive = 1" if active_only else "")

        results = self.db.execute_query(
            query,
            cache_key=f'users_{"active" if active_only else "all"}',
            cache_timeout=1800  # کش طولانی برای کاربران
        )

        for user in results:
            user['Initials'] = self._get_initials(user.get('FullName', ''))
            user['Domain'] = user.get('LoginName', '').split('\\')[0] if '\\' in user.get('LoginName', '') else ''

        logger.info(f"Fetched {len(results)} users")
        return results

    def get_statistics(self) -> Dict[str, Any]:
        """
        دریافت آمار کلی سایت
        """
        query = """
        SELECT 
            (SELECT COUNT(DISTINCT tp_ID) FROM AllLists WHERE tp_Hidden = 0) as TotalLists,
            (SELECT COUNT(*) FROM AllDocs WHERE Type = 0 AND DeleteTransactionId = 0x) as TotalDocuments,
            (SELECT COUNT(*) FROM UserInfo WHERE tp_IsActive = 1) as ActiveUsers,
            (SELECT CAST(SUM(Size) / 1048576.0 as DECIMAL(10,2)) FROM AllDocs WHERE Type = 0) as TotalSizeMB,
            (SELECT COUNT(*) FROM AllUserData WHERE tp_Created >= CAST(GETDATE() AS DATE)) as ItemsToday,
            (SELECT COUNT(*) FROM AllUserData WHERE tp_Modified >= DATEADD(day, -7, GETDATE())) as ItemsThisWeek
        """

        result = self.db.execute_query(query, cache_key='statistics', cache_timeout=120)
        stats = result[0] if result else {}

        # آمار اضافی
        additional_stats_query = """
        SELECT 
            L.tp_Title as ListName,
            COUNT(UD.tp_ID) as ItemCount
        FROM AllLists L WITH (NOLOCK)
        LEFT JOIN AllUserData UD WITH (NOLOCK) ON L.tp_ID = UD.tp_ListId
        WHERE L.tp_Hidden = 0
            AND UD.tp_DeleteTransactionId = 0x
        GROUP BY L.tp_Title
        ORDER BY ItemCount DESC
        """

        lists_stats = self.db.execute_query(additional_stats_query, cache_key='lists_statistics')
        stats['ListsBreakdown'] = lists_stats[:10]  # Top 10 lists

        # فرمت اعداد
        stats['TotalSizeGB'] = round(stats.get('TotalSizeMB', 0) / 1024, 2)
        stats['TotalSizeFormatted'] = self._format_file_size(stats.get('TotalSizeMB', 0) * 1048576)

        logger.info("Fetched site statistics")
        return stats

    def search(self, keyword: str, search_type: str = 'all', limit: int = 50) -> List[Dict[str, Any]]:
        """
        جستجوی عمومی در محتوا
        """
        if len(keyword) < 3:
            return []

        search_term = f'%{keyword}%'
        results = []

        if search_type in ['all', 'items']:
            # جستجو در آیتم‌های لیست‌ها
            items_query = """
            SELECT TOP (?)
                'ListItem' as Type,
                UD.tp_ID as Id,
                UD.nvarchar1 as Title,
                UD.ntext2 as Content,
                L.tp_Title as Location,
                UD.tp_Created as Created,
                UD.tp_Modified as Modified,
                U.tp_Title as Author
            FROM AllUserData UD WITH (NOLOCK)
            INNER JOIN AllLists L WITH (NOLOCK) ON UD.tp_ListId = L.tp_ID
            LEFT JOIN UserInfo U WITH (NOLOCK) ON UD.tp_Author = U.tp_ID
            WHERE (UD.nvarchar1 LIKE ? OR UD.ntext2 LIKE ?)
                AND UD.tp_DeleteTransactionId = 0x
                AND L.tp_Hidden = 0
            ORDER BY UD.tp_Modified DESC
            """

            items = self.db.execute_query(
                items_query,
                params=(limit, search_term, search_term)
            )
            results.extend(items)

        if search_type in ['all', 'documents']:
            # جستجو در اسناد
            docs_query = """
            SELECT TOP (?)
                'Document' as Type,
                D.Id,
                D.LeafName as Title,
                D.DirName as Content,
                'Documents' as Location,
                D.TimeCreated as Created,
                D.TimeLastModified as Modified,
                U.tp_Title as Author
            FROM AllDocs D WITH (NOLOCK)
            LEFT JOIN UserInfo U WITH (NOLOCK) ON D.AuthorId = U.tp_ID
            WHERE D.LeafName LIKE ?
                AND D.Type = 0
                AND D.DeleteTransactionId = 0x
            ORDER BY D.TimeLastModified DESC
            """

            docs = self.db.execute_query(
                docs_query,
                params=(limit, search_term)
            )
            results.extend(docs)

        # پردازش نتایج
        for result in results:
            result['CreatedFormatted'] = self._format_date(result.get('Created'))
            result['ModifiedFormatted'] = self._format_date(result.get('Modified'))

            # Highlight کردن کلمه جستجو
            if result.get('Title'):
                result['TitleHighlighted'] = self._highlight_text(result['Title'], keyword)
            if result.get('Content'):
                result['ContentSnippet'] = self._get_snippet(result['Content'], keyword)

        logger.info(f"Search for '{keyword}' returned {len(results)} results")
        return results

    # متدهای کمکی (Helper Methods)

    def _clean_html(self, html_content: str) -> str:
        """حذف تگ‌های HTML"""
        if not html_content:
            return ''

        # Unescape HTML entities
        text = html.unescape(html_content)

        # Remove HTML tags
        clean = re.compile('<.*?>')
        text = re.sub(clean, '', text)

        # Remove extra whitespace
        text = ' '.join(text.split())

        return text.strip()

    def _format_date(self, date_obj: Any) -> str:
        """فرمت‌دهی تاریخ به شمسی"""
        if not date_obj:
            return ''

        try:
            if isinstance(date_obj, str):
                date_obj = datetime.fromisoformat(date_obj)

            # اینجا می‌تونید از کتابخانه jdatetime برای تبدیل به شمسی استفاده کنید
            return date_obj.strftime('%Y/%m/%d')
        except:
            return str(date_obj)

    def _format_datetime(self, datetime_obj: Any) -> str:
        """فرمت‌دهی تاریخ و زمان"""
        if not datetime_obj:
            return ''

        try:
            if isinstance(datetime_obj, str):
                datetime_obj = datetime.fromisoformat(datetime_obj)

            return datetime_obj.strftime('%Y/%m/%d %H:%M')
        except:
            return str(datetime_obj)

    def _get_relative_time(self, date_obj: Any) -> str:
        """محاسبه زمان نسبی (مثلا: 2 ساعت پیش)"""
        if not date_obj:
            return ''

        try:
            if isinstance(date_obj, str):
                date_obj = datetime.fromisoformat(date_obj)

            now = datetime.now()
            diff = now - date_obj

            if diff.days > 30:
                return f"{diff.days // 30} ماه پیش"
            elif diff.days > 0:
                return f"{diff.days} روز پیش"
            elif diff.seconds > 3600:
                return f"{diff.seconds // 3600} ساعت پیش"
            elif diff.seconds > 60:
                return f"{diff.seconds // 60} دقیقه پیش"
            else:
                return "همین الان"
        except:
            return ''

    def _format_file_size(self, size_bytes: int) -> str:
        """فرمت‌دهی حجم فایل"""
        if size_bytes == 0:
            return '0 B'

        units = ['B', 'KB', 'MB', 'GB', 'TB']
        i = 0
        size = float(size_bytes)

        while size >= 1024 and i < len(units) - 1:
            size /= 1024
            i += 1

        return f"{size:.2f} {units[i]}"

    def _get_file_icon(self, extension: str) -> str:
        """آیکون بر اساس پسوند فایل"""
        icons = {
            'pdf': '📄',
            'doc': '📝',
            'docx': '📝',
            'xls': '📊',
            'xlsx': '📊',
            'ppt': '📽️',
            'pptx': '📽️',
            'zip': '🗜️',
            'rar': '🗜️',
            'jpg': '🖼️',
            'jpeg': '🖼️',
            'png': '🖼️',
            'gif': '🖼️',
            'mp4': '🎬',
            'mp3': '🎵',
            'txt': '📃'
        }
        return icons.get(extension.lower(), '📎')

    def _get_list_icon(self, template_id: int) -> str:
        """آیکون بر اساس نوع لیست"""
        icons = {
            100: '📋',  # Custom List
            101: '📁',  # Document Library
            102: '📊',  # Survey
            103: '🔗',  # Links
            104: '📢',  # Announcements
            105: '👥',  # Contacts
            106: '📅',  # Events
            107: '✓',  # Tasks
            108: '💬',  # Discussion Board
            109: '🖼️'  # Picture Library
        }
        return icons.get(template_id, '📄')

    def _build_document_url(self, folder_path: str, file_name: str) -> str:
        """ساخت URL دانلود سند"""
        base_url = settings.SHAREPOINT_DB['SITE_URL']
        path = f"{folder_path}/{file_name}".replace('//', '/').strip('/')
        return f"{base_url}/{path}"

    def _calculate_duration(self, start_time: Any, end_time: Any) -> str:
        """محاسبه مدت زمان"""
        if not start_time or not end_time:
            return ''

        try:
            if isinstance(start_time, str):
                start_time = datetime.fromisoformat(start_time)
            if isinstance(end_time, str):
                end_time = datetime.fromisoformat(end_time)

            diff = end_time - start_time
            hours = diff.seconds // 3600
            minutes = (diff.seconds % 3600) // 60

            if diff.days > 0:
                return f"{diff.days} روز"
            elif hours > 0:
                return f"{hours} ساعت"
            else:
                return f"{minutes} دقیقه"
        except:
            return ''

    def _get_initials(self, full_name: str) -> str:
        """دریافت حروف اول نام"""
        if not full_name:
            return '?'

        parts = full_name.split()
        if len(parts) >= 2:
            return f"{parts[0][0]}{parts[-1][0]}".upper()
        elif parts:
            return parts[0][:2].upper()
        return '?'

    def _highlight_text(self, text: str, keyword: str) -> str:
        """Highlight کردن کلمه در متن"""
        if not text or not keyword:
            return text

        pattern = re.compile(re.escape(keyword), re.IGNORECASE)
        return pattern.sub(f'<mark>{keyword}</mark>', text)

    def _get_snippet(self, content: str, keyword: str, context_length: int = 100) -> str:
        """دریافت بخشی از متن حاوی کلمه جستجو"""
        if not content or not keyword:
            return content[:200] if content else ''

        # پیدا کردن موقعیت کلمه
        index = content.lower().find(keyword.lower())
        if index == -1:
            return content[:200]

        # استخراج متن اطراف کلمه
        start = max(0, index - context_length)
        end = min(len(content), index + len(keyword) + context_length)

        snippet = content[start:end]
        if start > 0:
            snippet = '...' + snippet
        if end < len(content):
            snippet = snippet + '...'

        return self._highlight_text(snippet, keyword)