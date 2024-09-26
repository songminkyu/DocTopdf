using CommunityToolkit.Mvvm.Input;
using DocToPdf.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocToPdf.ViewModel
{
    public class MainViewModel : MainViewModelBase
    {
        public MainViewModel()
        {
            LoggingService.LoggingInit();
            UserControlLoadedCommand = new RelayCommand<object>(UserControlLoadedCommandExe);

            /*
             알림 팝업창 추가할때 아래내용 사용 예정

             타이틀 :               
             문서 변환 알림
             
             본문 : 

             문서 변환 목록이 발견 되었으며, 한글(HWP) 및 PowerPoint(MS Office) 문서가 열려 있습니다. 
             원활한 문서 변환을 위해 작성 중인 문서를 "저장" 또는 "닫기"를 진행 한 후 "확인" 버튼을 클릭 하여, 
             문서 변환을 진행 하십시오.
             */
        }

        private void UserControlLoadedCommandExe(object? obj)
        {
            var r = 1;          
        }
    }
}
