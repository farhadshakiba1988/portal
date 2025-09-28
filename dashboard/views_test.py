# dashboard/views_test.py
"""ویو تست بروز شده"""
from django.shortcuts import render
from django.http import JsonResponse
from .services_fixed import SharePointServiceFixed
import json


def test_dashboard_v2(request):
    """داشبورد تست نسخه 2"""
    sp_service = SharePointServiceFixed()

    # تست همه جداول
    all_tables = sp_service.test_all_tables()

    # دریافت داده‌ها
    lists_data = sp_service.get_all_lists_simple()
    users_data = sp_service.get_users_simple()
    userdata = sp_service.get_user_data_simple()

    context = {
        'page_title': 'تست اتصال SharePoint - نسخه 2',
        'tables_status': all_tables,
        'lists': lists_data,
        'users': users_data,
        'userdata': userdata,
    }

    return render(request, 'dashboard/test_v2.html', context)