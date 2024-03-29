from django.http import Http404
from django.shortcuts import render, get_object_or_404, redirect
from django.http import HttpResponse, HttpResponseRedirect
from django.http import JsonResponse
from django.template import loader
from django.urls import reverse
from django.views import generic
from django.utils import timezone
from django.contrib import messages

import os
import subprocess
from datetime import datetime, timedelta
import json
import pandas as pd

from MailControler.model_data import DataControl
from .settings import DEFAULT_DB_CONNECTION

# DEFAULT_DB_CONNECTION = 'SERVER'

def view_MDChecker_Main(request):
    dc = DataControl(DEFAULT_DB_CONNECTION)
    df = dc.request_mails_summary()
    data = df.to_json(orient='records')
    mails = json.loads(data)

    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    context = {
        'mails': mails,
        'now': now,
    }
    return render(request, 'MDChecker_Main.html', context)

def view_MDCheckerPension(request):
    dc = DataControl(DEFAULT_DB_CONNECTION)
    if request.method == 'GET':
        attachment = request.GET['attachment']
        df = dc.request_pension_data(attachment)
        if df is not None:
            data = df.to_json(orient='records')
            pension = json.loads(data)
            site_names = df['현장명p'].to_list()
            site_name = site_names[0]
            site_code = dc.request_site_code_by_site_name(site_name)
            df['현장코드'] = site_code
        else:
            pension = None
            site_code = None
    else:
        attachment = request.POST.get('attachment')
        site_code = request.POST.get('site_code')
        df = dc.request_pension_data(attachment)
        data = df.to_json(orient='records')
        pension = json.loads(data)
        df['현장코드'] = site_code

        # 하나의 파일에 있는 신규 현장명 등록 : mdc_mst_site : site_code / site_name_p upload
        dc.insert_new_site_name_to_mdc_mst_site(df)

    if site_code is not None:
        df = dc.request_pension_result_data(site_code)
        if df is None:
            result = None
        else:
            data = df.to_json(orient='records')
            result = json.loads(data)
    else:
        result = None

    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    context = {
        'attachment': attachment,
        'site_code': site_code,
        'pension': pension,
        'result': result,
        'now': now,
    }
    return render(request, 'MDChecker_Pension.html', context)

def view_MDCheckerSubcon(request):
    dc = DataControl(DEFAULT_DB_CONNECTION)
    error_message = False

    if request.method == 'POST':
        param = request.POST.dict()
        subcon_name_key = param['subcon_name_key']
        subcon_name_simular = param['subcon_name_simular']
        result = dc.insert_subcon_maching_data_to_db(subcon_name_key, subcon_name_simular)
        if not result:
            error_message = "등록에 실패하였습니다. 관리자 문의 바랍니다."

    df = dc.call_df_from_db_with_column_name('mdc_mst_subcon')
    df = df.sort_values(['업체명key', '업체명'])
    data = df.to_json(orient='records')
    subcon_list = json.loads(data)

    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    context = {
        'subcon_list': subcon_list,
        'now': now,
        'error_message': error_message,
    }

    return render(request, 'MDChecker_Subcon.html', context)

def view_MDCheckerSubconAjax(request):
    dc = DataControl(DEFAULT_DB_CONNECTION)
    param = json.loads(request.body)
    subcon_name_key = param['subcon_name_key']
    subcon_name_simular = param['subcon_name_simular']
    result = dc.delete_subcon_maching_data_to_db(subcon_name_key, subcon_name_simular)
    error_message = False
    if not result:
        error_message = "삭제에 실패하였습니다. 관리자 문의 바랍니다."

    df = dc.call_df_from_db_with_column_name('mdc_mst_subcon')
    df = df.sort_values(['업체명key', '업체명'])
    data = df.to_json(orient='records')
    subcon_list = json.loads(data)

    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    context = {
        'subcon_list': subcon_list,
        'now': now,
        'error_message': error_message,
    }

    return JsonResponse(context)

def view_MDCheckerAddress(request):
    dc = DataControl(DEFAULT_DB_CONNECTION)
    error_message = False

    df = dc.request_address_mdc_all()
    data = df.to_json(orient='records')
    address_list = json.loads(data)

    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    context = {
        'address_list': address_list,
        'now': now,
        'error_message': error_message,
    }

    return render(request, 'MDChecker_Address.html', context)

def view_MDCheckerAddressAdd(request):
    dc = DataControl(DEFAULT_DB_CONNECTION)
    error_message = False
    if request.method == 'POST':
        param = request.POST.dict()
        address_name = param['address_name']
        address_mail = param['address_mail']
        address_site_code = param['address_site_code']
        address_department = param['address_department']
        address_managing_bonbu = param['address_managing_bonbu']
        result = dc.insert_address_data_to_db(address_name, address_mail, address_site_code, address_department, address_managing_bonbu)
        if not result:
            error_message = "등록에 실패하였습니다. 관리자 문의 바랍니다."
    else:
        error_message = "잘못된 접근입니다."

    df = dc.request_address_mdc_all()
    data = df.to_json(orient='records')
    address_list = json.loads(data)

    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    context = {
        'address_list': address_list,
        'now': now,
        'error_message': error_message,
    }

    return render(request, 'MDChecker_Address.html', context)

def view_MDCheckerAddressDel(request):
    dc = DataControl(DEFAULT_DB_CONNECTION)
    error_message = False
    if request.method == 'POST':
        param = json.loads(request.body)
        del_address = param['del_address']
        result = dc.delete_address_data_to_db(del_address)
        if not result:
            error_message = "삭제에 실패하였습니다. 관리자 문의 바랍니다."
    else:
        error_message = "잘못된 접근입니다."

    context = {
        'error_message': error_message,
    }

    return JsonResponse(context)

def view_MDCheckerRunAll(request):
    import win32com.client

    scheduler = win32com.client.Dispatch('Schedule.Service')
    scheduler.Connect()

    objTaskFolder = scheduler.GetFolder("\\")
    colTasks = objTaskFolder.GetTasks(1)

    for task in colTasks:
        print(task.Name)
        # print(task.Enabled)
        # print(task.LastRunTime)
        # print(task.LastTaskResult)
        # print(task.NextRunTime)
        # print(task.NumberOfMissedRuns)
        # print(task.State)
        # print(task.Path)
        # print(task.XML)

        if task.Name == "퇴직공제_SERVER":
            task.Enabled = True
            runningTask = task.Run("")
            task.Enabled = False

    messages.info(request, "실행중이니 약 1분 후 새로고침하시기 바랍니다.")

    return redirect('MDCheckerMain')

def view_MDCheckerRunAll_ADMIN(request):
    import win32com.client

    scheduler = win32com.client.Dispatch('Schedule.Service')
    scheduler.Connect()

    objTaskFolder = scheduler.GetFolder("\\")
    colTasks = objTaskFolder.GetTasks(1)

    for task in colTasks:
        print(task.Name)
        # print(task.Enabled)
        # print(task.LastRunTime)
        # print(task.LastTaskResult)
        # print(task.NextRunTime)
        # print(task.NumberOfMissedRuns)
        # print(task.State)
        # print(task.Path)
        # print(task.XML)

        if task.Name == "퇴직공제_ADMIN":
            task.Enabled = True
            runningTask = task.Run("")
            task.Enabled = False

    messages.info(request, "실행중, 최다희 매니저에게만 메일송부 됩니다.")
    return redirect('MDCheckerMain')

def view_MDCheckerDownloadAll(request):
    dc = DataControl(DEFAULT_DB_CONNECTION)
    df = dc.request_pension_result_data_full()
    if df is None:
        messages.info(request, "요청하신 데이터가 없습니다. 담당자와 상의하세요.")
        return redirect('MDCheckerMain')

    # save in memory by pandas / io
    import io
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer) as writer:
        df.to_excel(writer, sheet_name='data')
    buffer.seek(0)

    # response
    now = datetime.now().strftime("%Y%m%d%H%M%S")
    filename = f"pension_full_data_{now}.xlsx"
    response = HttpResponse(buffer.read(),
                            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response['Content-Disposition'] = f"attachment; filename={filename}"

    # memory close
    buffer.close()

    return response


