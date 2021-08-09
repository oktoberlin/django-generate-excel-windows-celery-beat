from . import views
from django.urls import path
app_name = 'home'
urlpatterns = [
    path('', views.index, name="home"),
    path('success', views.success, name="success"),
]