# dashboard/views_temp.py
"""ویو موقت برای تست"""
from django.shortcuts import render
from django.http import JsonResponse
from .services_simple import SharePointServiceSimple
import logging

logger = logging.getLogger(__name__)


def test_dashboard(request):
    """داشبورد تست"""
    sp_service = SharePointServiceSimple()

    context = {
        'page_title': 'تست اتصال',
        'data': sp_service.get_simple_data(),
        'stats': sp_service.get_statistics_simple(),
        'tables': sp_service.get_all_tables()[:20],  # فقط 20 جدول اول
    }

    return render(request, 'dashboard/test.html', context)


def api_test_connection(request):
    """API تست"""
    sp_service = SharePointServiceSimple()
    data = sp_service.test_query()

    return JsonResponse({
        'success': len(data) > 0,
        'message': f'دریافت {len(data)} رکورد',
        'sample': data[0] if data else None
    })