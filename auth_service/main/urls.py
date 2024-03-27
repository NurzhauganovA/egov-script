from django.urls import path
from . import views


urlpatterns = [
    path('', views.SendEgovAuthRequest.as_view(), name='main'),
]