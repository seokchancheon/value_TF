from django.http import HttpResponse, HttpResponseRedirect
from django.shortcuts import render
from .forms import ExcelUploadForm
import xlwings as xw
import os
from tempfile import NamedTemporaryFile
from pricing.TF.TF_model import *

def upload_excel_view(request):
    if request.method == 'POST':
        form = ExcelUploadForm(request.POST, request.FILES)
        if form.is_valid():
            uploaded_file = request.FILES['excel_file']
            if uploaded_file.name.endswith('.xlsx'):
                # 엑셀 파일 처리
                # 파일을 임시 파일로 저장
                with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
                    for chunk in uploaded_file.chunks():
                        temp_file.write(chunk)

                wb = None
                file_name = None
                sheet_name = 'TF모형(BDT)_python'

                if temp_file.name:  # temp_file.name이 None이 아닌 경우에만 파일을 열도록 처리
                    wb = xw.Book(temp_file.name)
                    file_name = temp_file.name

                # 이 부분이 잘 동작해야 함
                Result_file = TF_model(wb, sheet_name, file_name)
                #######################

#                if Result_file:
#                    # 결과 파일을 클라이언트에게 전달
#                    with open(Result_file, 'rb') as file:
#                        response = HttpResponse(file.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
#                        response['Content-Disposition'] = f'attachment; filename=Result_file.xlsx'
#                    return response
            else:
                # 엑셀 파일이 아닌 경우 처리
                pass
        else:
            # 폼이 유효하지 않을 때 처리
            pass
    else:
        form = ExcelUploadForm()

    return render(request, 'upload_excel.html', {'form': form})