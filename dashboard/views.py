"""
Dashboard Views
ویوهای اصلی داشبورد
"""
from django.shortcuts import render, redirect
from django.http import JsonResponse, HttpResponse
from django.views.decorators.cache import cache_page
from django.views.decorators.http import require_http_methods
from django.core.paginator import Paginator
from django.contrib import messages
import json
import logging
from datetime import datetime
from .services import SharePointService
from .db_connection import SharePointDatabase

logger = logging.getLogger(__name__)


class DashboardViews:
    """کلاس اصلی برای مدیریت ویوهای داشبورد"""

    def __init__(self):
        self.sp_service = SharePointService()
        self.sp_db = SharePointDatabase()

    def index(self, request):
        """صفحه اصلی داشبورد"""
        try:
            # بررسی اتصال
            if not self.sp_db.test_connection():
                messages.error(request, 'عدم اتصال به دیتابیس SharePoint')
                return render(request, 'dashboard/error.html')

            # دریافت داده‌ها
            context = {
                'page_title': 'داشبورد پرتال',
                'announcements': self.sp_service.get_announcements(5),
                'documents': self.sp_service.get_documents(10),
                'tasks': self.sp_service.get_tasks()[:5],
                'events': self.sp_service.get_calendar_events(7),
                'statistics': self.sp_service.get_statistics(),
                'current_time': datetime.now(),
            }

            return render(request, 'dashboard/index.html', context)

        except Exception as e:
            logger.error(f"Error loading dashboard: {e}")
            messages.error(request, 'خطا در بارگذاری داشبورد')
            return render(request, 'dashboard/error.html')

    def announcements_list(self, request):
        """صفحه لیست اطلاعیه‌ها"""
        try:
            announcements = self.sp_service.get_announcements(50)

            # صفحه‌بندی
            paginator = Paginator(announcements, 10)
            page_number = request.GET.get('page')
            page_obj = paginator.get_page(page_number)

            context = {
                'page_title': 'اطلاعیه‌ها',
                'page_obj': page_obj,
            }

            return render(request, 'dashboard/announcements.html', context)

        except Exception as e:
            logger.error(f"Error loading announcements: {e}")
            messages.error(request, 'خطا در بارگذاری اطلاعیه‌ها')
            return redirect('dashboard:index')

    def documents_list(self, request):
        """صفحه لیست اسناد"""
        try:
            # فیلترها
            extension = request.GET.get('ext')
            extensions = [extension] if extension else None

            documents = self.sp_service.get_documents(100, extensions)

            # صفحه‌بندی
            paginator = Paginator(documents, 20)
            page_number = request.GET.get('page')
            page_obj = paginator.get_page(page_number)

            context = {
                'page_title': 'اسناد',
                'page_obj': page_obj,
                'current_filter': extension,
            }

            return render(request, 'dashboard/documents.html', context)

        except Exception as e:
            logger.error(f"Error loading documents: {e}")
            messages.error(request, 'خطا در بارگذاری اسناد')
            return redirect('dashboard:index')

    def search(self, request):
        """جستجو"""
        query = request.GET.get('q', '')
        search_type = request.GET.get('type', 'all')

        if len(query) < 3:
            messages.warning(request, 'لطفا حداقل 3 کاراکتر وارد کنید')
            return redirect('dashboard:index')

        try:
            results = self.sp_service.search(query, search_type)

            context = {
                'page_title': f'نتایج جستجو: {query}',
                'query': query,
                'results': results,
                'result_count': len(results),
            }

            return render(request, 'dashboard/search.html', context)

        except Exception as e:
            logger.error(f"Search error: {e}")
            messages.error(request, 'خطا در جستجو')
            return redirect('dashboard:index')


# ایجاد instance از کلاس
dashboard_views = DashboardViews()


# View functions
def index(request):
    return dashboard_views.index(request)


def announcements_list(request):
    return dashboard_views.announcements_list(request)


def documents_list(request):
    return dashboard_views.documents_list(request)


def search(request):
    return dashboard_views.search(request)


# API Views
@require_http_methods(["GET"])
def api_announcements(request):
    """API برای دریافت اطلاعیه‌ها"""
    try:
        limit = int(request.GET.get('limit', 10))
        sp_service = SharePointService()
        data = sp_service.get_announcements(limit)

        return JsonResponse({
            'success': True,
            'data': data,
            'count': len(data)
        })
    except Exception as e:
        logger.error(f"API Error: {e}")
        return JsonResponse({
            'success': False,
            'error': str(e)
        }, status=500)


@require_http_methods(["GET"])
def api_documents(request):
    """API برای دریافت اسناد"""
    try:
        limit = int(request.GET.get('limit', 20))
        sp_service = SharePointService()
        data = sp_service.get_documents(limit)

        return JsonResponse({
            'success': True,
            'data': data,
            'count': len(data)
        })
    except Exception as e:
        logger.error(f"API Error: {e}")
        return JsonResponse({
            'success': False,
            'error': str(e)
        }, status=500)


@require_http_methods(["GET"])
def api_tasks(request):
    """API برای دریافت وظایف"""
    try:
        assigned_to = request.GET.get('assigned_to')
        status = request.GET.get('status')

        sp_service = SharePointService()
        data = sp_service.get_tasks(assigned_to, status)

        return JsonResponse({
            'success': True,
            'data': data,
            'count': len(data)
        })
    except Exception as e:
        logger.error(f"API Error: {e}")
        return JsonResponse({
            'success': False,
            'error': str(e)
        }, status=500)


@require_http_methods(["GET"])
@cache_page(60)  # کش برای 1 دقیقه
def api_statistics(request):
    """API برای دریافت آمار"""
    try:
        sp_service = SharePointService()
        data = sp_service.get_statistics()

        return JsonResponse({
            'success': True,
            'data': data
        })
    except Exception as e:
        logger.error(f"API Error: {e}")
        return JsonResponse({
            'success': False,
            'error': str(e)
        }, status=500)


@require_http_methods(["GET"])
def api_search(request):
    """API جستجو"""
    try:
        query = request.GET.get('q', '')
        search_type = request.GET.get('type', 'all')

        if len(query) < 3:
            return JsonResponse({
                'success': False,
                'error': 'حداقل 3 کاراکتر وارد کنید'
            }, status=400)

        sp_service = SharePointService()
        results = sp_service.search(query, search_type)

        return JsonResponse({
            'success': True,
            'query': query,
            'results': results,
            'count': len(results)
        })
    except Exception as e:
        logger.error(f"Search API Error: {e}")
        return JsonResponse({
            'success': False,
            'error': str(e)
        }, status=500)


@require_http_methods(["GET"])
def api_calendar(request):
    """API تقویم"""
    try:
        days = int(request.GET.get('days', 30))
        sp_service = SharePointService()
        data = sp_service.get_calendar_events(days)

        return JsonResponse({
            'success': True,
            'events': data,
            'count': len(data)
        })
    except Exception as e:
        logger.error(f"Calendar API Error: {e}")
        return JsonResponse({
            'success': False,
            'error': str(e)
        }, status=500)


@require_http_methods(["GET"])
def test_connection(request):
    """تست اتصال به دیتابیس"""
    try:
        sp_db = SharePointDatabase()

        if sp_db.test_connection():
            info = sp_db.get_database_info()
            return JsonResponse({
                'success': True,
                'message': 'اتصال موفق',
                'info': info
            })
        else:
            return JsonResponse({
                'success': False,
                'message': 'عدم اتصال به دیتابیس'
            }, status=500)

    except Exception as e:
        return JsonResponse({
            'success': False,
            'error': str(e)
        }, status=500)