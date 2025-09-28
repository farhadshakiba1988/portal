"""
Dashboard URL Configuration
"""
from django.urls import path
from . import views, views_test
from . import views_temp  # اضافه کنید

app_name = 'dashboard'

urlpatterns = [
    path('test/v2/', views_test.test_dashboard_v2, name='test_v2'),
    # تست
    path('test/', views_temp.test_dashboard, name='test'),
    path('api/test/', views_temp.api_test_connection, name='api_test'),
    # صفحات اصلی
    path('', views.index, name='index'),
    path('announcements/', views.announcements_list, name='announcements'),
    path('documents/', views.documents_list, name='documents'),
    path('search/', views.search, name='search'),

    # API Endpoints
    path('api/announcements/', views.api_announcements, name='api_announcements'),
    path('api/documents/', views.api_documents, name='api_documents'),
    path('api/tasks/', views.api_tasks, name='api_tasks'),
    path('api/statistics/', views.api_statistics, name='api_statistics'),
    path('api/search/', views.api_search, name='api_search'),
    path('api/calendar/', views.api_calendar, name='api_calendar'),
    path('api/test/', views.test_connection, name='test_connection'),
]
