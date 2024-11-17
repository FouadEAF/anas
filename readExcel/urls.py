from django.urls import path
from .views import run_excel

urlpatterns = [
    # path('/read', get_excel),
    path('sent', run_excel),
]
