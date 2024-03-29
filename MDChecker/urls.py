"""MDChecker URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/4.0/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import path
from . import views

urlpatterns = [
    # path('admin/', admin.site.urls),
    path('MDCheckerMain/', views.view_MDChecker_Main, name='MDCheckerMain'),
    path('MDCheckerPension/', views.view_MDCheckerPension, name='MDCheckerPension'),
    path('MDCheckerSubcon/', views.view_MDCheckerSubcon, name='MDCheckerSubcon'),
    path('MDCheckerSubconAjax/', views.view_MDCheckerSubconAjax, name='MDCheckerSubconAjax'),
    path('MDCheckerAddress/', views.view_MDCheckerAddress, name='MDCheckerAddress'),
    path('MDCheckerAddressAdd/', views.view_MDCheckerAddressAdd, name='MDCheckerAddressAdd'),
    path('MDCheckerAddressDel/', views.view_MDCheckerAddressDel, name='MDCheckerAddressDel'),
    path('MDCheckerRunAll/', views.view_MDCheckerRunAll, name='MDCheckerRunAll'),
    path('MDCheckerRunAllAdmin/', views.view_MDCheckerRunAll_ADMIN, name='MDCheckerRunAllAdmin'),
    path('MDCheckerDownloadAll/', views.view_MDCheckerDownloadAll, name='MDCheckerDownloadAll'),

]
