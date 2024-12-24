## [한글(HWP)을 PDF로 변환하는 모듈]

#### 요구사항
* Language         : python
* Lenguage Version : 3.11.5

#### Build Setup
* pip install requirements.txt
  
#### 실행파일 한개 출력 (Pyinstaller 설치 필요)
* Pyinstaller --clean --onefile --windowed hnc_to_pdf.py
  Pyinstaller --clean --onefile --noconsole hnc_to_pdf.py
  Pyinstaller --clean --onefile -w hnc_to_pdf.py

#### 실행 방법
* hwp_to_pdf.exe -hp <한글 저장 폴더 경로> -pn <제품 이름(FMF,FML,FF)>
  
#### HWP OCX 우회 방법
* 컴퓨터\HKEY_CURRENT_USER\SOFTWARE\HNC\HwpAutomation\Modules 에 FilePathCheckerModule 변수를 만들고
  제품폴더/runtimes/FD_HWPSecurity/FilePathCheckerModuleExample.dll 파일위치를 레지스트리 변수에 입력한다

* 여기서 주의 할점은 HwpAutomation\Modules 키가 없으면 직접 키를 생성 해야함.
  HNC (컴퓨터\HKEY_CURRENT_USER\SOFTWARE\HNC)
   ㄴ HwpAutomation (생성)
      &nbsp;&nbsp;&nbsp;&nbsp; ㄴ Modules (생성) 
    