# -*- coding: utf-8 -*-
"""
Created on Tue Oct 31 23:01:43 2023

@author: cjstj
"""

from django.urls import path

from . import views

urlpatterns = [
    path('', views.upload_excel_view, name='upload_excel_view'),
    # 다른 URL 패턴들
]