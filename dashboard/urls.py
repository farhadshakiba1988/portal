"""
Dashboard URL Configuration
"""
from django.urls import path
from . import views

app_name = 'dashboard'

urlpatterns = [
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