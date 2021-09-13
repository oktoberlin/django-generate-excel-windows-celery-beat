from . import views
from django.urls import path

app_name = 'report'
urlpatterns = [
    path('', views.index, name="report"),
    path('api/', views.RepairListView.as_view(), name="api"),
    path('success', views.success, name="success"),
]