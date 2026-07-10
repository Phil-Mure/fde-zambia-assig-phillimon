from django.urls import path

from . import views

urlpatterns = [
    path("", views.dashboard, name="dashboard"),
    path("api/bulletin/", views.bulletin_json, name="bulletin_json"),
    path("export/", views.bulletin_export, name="bulletin_export"),
]
