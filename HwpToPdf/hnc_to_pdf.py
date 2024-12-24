import os  # .path.join(), .listdir(), .chdir(), .getcwd() 등 사용
import time

import win32com.client as win32  # 한/글 열 수 있는 모듈
import win32con
import win32gui  # 창 숨기기 위한 모듈
import win32api
import argparse
from PyPDF2 import PdfReader

def HWpSearch(dirname, hwpfiles):
    try:
        filenames = os.listdir(dirname)
        for filename in filenames:
            full_filename = os.path.join(dirname, filename)
            if os.path.isdir(full_filename):
                HWpSearch(full_filename, hwpfiles)
            else:
                ext = os.path.splitext(full_filename)[-1]
                if (ext.casefold() == '.hwp' or ext.casefold() == '.hwpx'):
                    hwpfiles.append(full_filename)
    except PermissionError:
        pass


def replace_hwp_to_pdf_export_path(pdf_save_path,org_hwpfile_path):
    org_hwpfile_path_lower = org_hwpfile_path.casefold()
    hwp_ext = os.path.splitext(org_hwpfile_path_lower)[-1]
    org_hwpfile_path_lower = org_hwpfile_path_lower.replace(hwp_ext, '.pdf')
    conv_pdf_filename = os.path.basename(org_hwpfile_path_lower)
    pdfs_save_filepath = os.path.join(pdf_save_path, conv_pdf_filename)
    return pdfs_save_filepath;


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('-hrp', '--hwp_root_path', type=str, help=' : Please specify the location of the HWP document',
                        default="c:\\Doc\\HWPX")
    parser.add_argument('-ps', '--pdf_save_path', type=str, help=' : Please specify the location of the PDF document',
                        default="c:\\Doc\\PDFS")
    parser.add_argument('-scm','--security_module', type=str, help=' : Please specify the security module',default='FilePathCheckerModule')
    ToConvertPdfStartMessage = win32con.WM_USER + 1000
    ToConvertPdfMessage = win32con.WM_USER + 1001

    arguments = parser.parse_args()

    hwp_root_path = arguments.hwp_root_path

    pdf_save_path = arguments.pdf_save_path

    security_module = arguments.security_module

    os.chdir(hwp_root_path)  # hwp 파일이 있는 폴더로 이동
    '''
    for i in os.listdir():  # 현재 폴더 안에 있는 모든 파일명에서
        os.rename(i, i.replace(' - 복사본 ', ''))  # ' - 복사본 ' 부분을 지워줘.
    '''
    hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')  # 한/글 열기
    hwnd = win32gui.FindWindow(None, '빈 문서 1 - 한글')  # 해당 윈도우의 핸들값 찾기

    win32gui.ShowWindow(hwnd, 0)  # 한/글 창을 숨겨줘. 0은 숨기기, 5는 보이기, 3은 풀스크린 등
    hwp.RegisterModule('FilePathCheckDLL', security_module)  # 보안모듈 적용
    hwp.XHwpWindows.Item(0).Visible = False

    BASE_DIR = hwp_root_path  # 한/글은 파일 열거나 저장할 때 전체경로를 입력해야 하므로, os.path.join(BASE_DIR, i) 식으로 사용할 것

    PDFS_DIR = pdf_save_path
    if not os.path.exists(PDFS_DIR):
        os.makedirs(PDFS_DIR)

    file_list = []

    HWpSearch(hwp_root_path, file_list)

    file_list_hwp = [file for file in file_list if file.endswith((".hwp",".HWP",".Hwp",".hwpx",".HWPX",".Hwpx"))]

    print("ToConvertPdfStartCount", 0, len(file_list_hwp), sep="|")

    '''
    변환 되지 않는 문서 (HWPX, 배포용 문서, 암호 걸린 문서)
    배포용 문서는 편집 자체가 불가능 또는 PDF 컨버팅 되지 않기 때문에 변환 할 수 없음.
    '''
    ToConvertPdfCount = 1
    for hwpfile in file_list_hwp:  # 현재폴더 안에 있는 모든 파일을
        replace_hwp2pdf_path = replace_hwp_to_pdf_export_path(pdf_save_path,hwpfile)
        isExists = os.path.isfile(replace_hwp2pdf_path)
        if isExists == False:
            hwp.XHwpWindows.Item(0).Visible = False
            hwp.Open(os.path.join(BASE_DIR, hwpfile),arg='forceopen:True;suspendpassword:True;versionworning:False')  # 한/글로 열어서
            hwp.HAction.GetDefault('FileSaveAsPdf', hwp.HParameterSet.HFileOpenSave.HSet)  # PDF로 저장할 건데, 설정값은 아래와 같이.
            hwp.HParameterSet.HFileOpenSave.filename = replace_hwp2pdf_path  # 확장자는 .pdf로,
            hwp.HParameterSet.HFileOpenSave.Format = 'PDF'  # 포맷은 PDF로,
            hwp.HAction.Execute('FileSaveAsPdf', hwp.HParameterSet.HFileOpenSave.HSet)  # 위 설정값으로 실행해줘.

            isConvertOk = True
            isExists = os.path.isfile(hwp.HParameterSet.HFileOpenSave.filename)
            if isExists == True:
                reader = PdfReader(hwp.HParameterSet.HFileOpenSave.filename, 'rb')
                page = reader.pages[0]
                pagelen = len(page.extract_text())
                if pagelen == 0 and len(reader.pages) == 1:  # Page가 1개 이면서 1개 페이지에 내용이 없으면 비밀번호 걸려있는 한글 문서로 확인.
                    isConvertOk = False
                    os.remove(hwp.HParameterSet.HFileOpenSave.filename)

            if isConvertOk == False or isExists == False:
                # 컨버팅이 되지 않는 문서들 대상으로 다음 문서 변환 시도
                continue

            print("ToConvertPdfCurrentCount", ToConvertPdfCount, replace_hwp2pdf_path, sep="|")
            ToConvertPdfCount = ToConvertPdfCount + 1

    win32gui.ShowWindow(hwnd, 0)  # 다시 숨겼던 한/글 창을 보여주고,
    hwp.XHwpDocuments.Close(isDirty=False)  # 열려있는 문서가 있다면 닫아줘(저장할지 물어보지 말고)
    hwp.Quit()  # 한/글 종료
    del hwp
    del win32
