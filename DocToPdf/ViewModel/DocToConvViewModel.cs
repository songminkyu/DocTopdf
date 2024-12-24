using CommunityToolkit.Mvvm.Input;
using DocToPdf.Services;
using DocToPdf.UIControlServices;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Collections.ObjectModel;

namespace DocToPdf.ViewModel
{
    public class DocToConvViewModel : DocToConvViewModelBase
    {
        public DocToConvViewModel()
        {
            LoggingService.LoggingInit();
            UserControlLoadedCommand = new RelayCommand<object>(UserControlLoadedCommandExe);
            targetPathCommand = new RelayCommand<object>(TargetPathCommandExe);
            savedPathCommand = new RelayCommand<object>(SavedPathCommandExe);
            runCommand = new AsyncRelayCommand<object>(RunCommandExe);
            cancelCommand = new RelayCommand<object>(CancelCommandExe);
            savedPath = string.Empty;
            targetPath = string.Empty;
            convLogs = new ObservableCollection<Model.DocConverter>();
        }
        private void UserControlLoadedCommandExe(object? obj)
        {
            if(convLogs != null)
                convLogs.Add(new Model.DocConverter() { index = 0, description = "123123" });
        }
        
        private void TargetPathCommandExe(object? obj)
        {
            string targetInitPath = KnownFoldersService.GetPath(KnownFolder.Documents);
            targetPath = OpenFileDlg(targetInitPath);
            if(!string.IsNullOrEmpty(targetPath))
            {
                ConvMSOfficeToDocService.SetRootPath(targetPath);
                ConvHncToDocService.SetRootPath(targetPath);
            }
        }
        private void SavedPathCommandExe(object? obj)
        {
            string savedInitPath = KnownFoldersService.GetPath(KnownFolder.Documents);
            savedPath = OpenFileDlg(savedInitPath);
        }
        private async Task RunCommandExe(object? arg)
        {
            await Task.Run(DocToPdfConvert);
        }
        private void CancelCommandExe(object? obj)
        {

        }
        private async Task DocToPdfConvert()
        {
      
            Progress<ConvertReport> ConvertProgressReport = new Progress<ConvertReport>(async value => {
                await DispatcherService.BeginInvokeBackground(new Action(async delegate
                {
                    await Task.Delay(0);

                    var r = $"Converting {value.ConvertType} To PDF...( {value.CurrentCount} / {value.TotalCount} )";

                    if (convLogs != null)
                        convLogs.Add(new Model.DocConverter() { index = value.CurrentCount, description = "123123" });

                    if (value.CurrentCount == value.TotalCount)
                    {
                        await Task.Delay(100);
                    }

                }));
            });
            bool IsExsistsExportHWPDir = ConvHncToDocService.IsExsistsHWPDir();
            bool IsExsistsExportPowerPointDir = ConvMSOfficeToDocService.IsExsistsPowerPointDir();
            int NotExsistsPowerPointToPDFCount = ConvMSOfficeToDocService.CheckedNotExsistsPowerPointToPDFCount();
            int NotExsistsHWPToPDFCount = ConvHncToDocService.CheckedNotExsistsHWPToPDFCount();

            if ((IsExsistsExportPowerPointDir == true && NotExsistsPowerPointToPDFCount > 0) || (IsExsistsExportHWPDir == true && NotExsistsHWPToPDFCount > 0))
            {
                MessageBoxResult result = MessageBoxResult.OK;
                while (FileControlService.OpenDocumentProcesses() == true)
                {
                    result = MessageBox.Show(
                                        "문서 변환 목록이 발견 되었으며, 한글(HWP) 및 PowerPoint(MS Office) 문서가 열려 있습니다.\n" +
                                        "원활한 문서 변환을 위해 작성 중인 문서를 '저장' 또는 '닫기'를 진행 한 후 '확인' 버튼을 클릭 하여,\n" +
                                        "문서 변환을 진행 하십시오.",
                                        "문서 변환 알림",
                                        MessageBoxButton.OKCancel,
                                        MessageBoxImage.Information);

                    if (result == MessageBoxResult.Cancel)
                    {
                        break;
                    }
                }

                if (result == MessageBoxResult.OK)
                {
                    bool IsPowerPointInstalled = ConvMSOfficeToDocService.IsPowerPointInstalled_V16();

                    if (IsPowerPointInstalled == true && IsExsistsExportPowerPointDir == true)
                    {
                        await ConvMSOfficeToDocService.ConvertPowerPointToPDFAll(ConvertProgressReport!).ConfigureAwait(false);
                    }


                    bool IsHnCInstalled = ConvHncToDocService.IsHnCInstalled();
                    bool IsHwpToPdfConverterExsist = ConvHncToDocService.IsHwpToPdfConverterExsist();
                    bool IsRegCheckerModule = ConvHncToDocService.ReadRegistryFilePathCheckerModule();

                    if (IsHnCInstalled == true &&  // 한글이 설치 여부 확인
                        IsRegCheckerModule == true &&  // 체커 모듈이 등록 되어 있는지 확인
                        IsExsistsExportHWPDir == true &&  // Export 된곳에 HWP 폴더가 존재 하는지 확인
                        IsHwpToPdfConverterExsist == true)    // 바이러스 백신으로부터 제거가 될수 있는 요소로 모듈이 존재 하는지 체크
                    {
                        var TokenSource = new CancellationTokenSource();
                        CancellationToken Token = TokenSource.Token;

                        _ = Task.Run(async () =>
                        {
                            List<string> titles = new List<string>()
                            {
                                "한글",
                                "폴더 찾아보기",
                                "스크립트 실행",
                                "패키지 내용 열기"
                            };
                            string? ProductName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name!;

                            bool IsWindowForeground = false;

                            while (Token.IsCancellationRequested != true)
                            {
                                IntPtr Producthwnd = User32APIService.FindWindow(null!, ProductName);
                                IntPtr Hwphwnd = User32APIService.FindWindow(null!, "빈 문서 1 - 한글");

                                if (Hwphwnd != IntPtr.Zero && IsWindowForeground == false)
                                {
                                    IsWindowForeground = User32APIWrapper.IsWindowForeground();

                                    if (Producthwnd != IntPtr.Zero && IsWindowForeground == false)
                                    {
                                        // 윈도우가 최소화 되어 있다면 활성화 시킨다
                                        User32APIService.ShowWindowAsync(Producthwnd, 0x0003);
                                        await Task.Delay(700);
                                        // 윈도우에 포커스를 줘서 최상위로 만든다
                                        User32APIService.SetForegroundWindow(Producthwnd);
                                        await Task.Delay(700);
                                        // 윈도우를 활성화 한다
                                        User32APIService.SetActiveWindow(Producthwnd);
                                        await Task.Delay(700);

                                        IsWindowForeground = true;
                                    }
                                }

                                foreach (var title in titles)
                                {
                                    IntPtr childhwnd = User32APIService.FindWindow("HNC_DIALOG", title);

                                    if (childhwnd != IntPtr.Zero)
                                    {
                                        //한글 팝업창 뜨면 close 명령어를 날린다.
                                        User32APIService.SendMessage(childhwnd, 0x0010, IntPtr.Zero, null!);
                                    }
                                }
                                await Task.Delay(5);
                            }

                        }, Token);


                        await ConvHncToDocService.ConvertHWPToPDFAll(ConvertProgressReport);

                        ConvHncToDocService.Remove_gen_py();

                        TokenSource.Cancel();

                        FileControlService.DocumentProcessKill();

                    }
                }
                else
                {
                }
            }
        }
        private string? OpenFileDlg(string InitialDirectory)
        {
            CommonOpenFileDialog openFileDialog = new CommonOpenFileDialog();
            openFileDialog.InitialDirectory = InitialDirectory; 
            openFileDialog.IsFolderPicker = true;

            if (openFileDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                return openFileDialog.FileName;
            }
            else
            {
                return string.Empty;
            }
        }
    }
}
