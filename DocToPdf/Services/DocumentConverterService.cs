using Microsoft.Office.Core;
using Microsoft.Win32;
using System.IO;
using System.Windows.Xps.Packaging;
using System.Windows.Interop;
using System.Diagnostics;
using System.Windows;

namespace DocToPdf.Services
{
    public enum ContentType
    {
        DOC,
        DOCX,
        XLS,
        XLSX,
        VSD,
        VDX,
        PPT,
        PPTX,
        XDW,
        PDF,
        XPS,
        JPEG,
        JPG,
        BMP,
        PNG,
        TIF,
        TIFF,
        GIF,
        SVG,
        TXT,
        RTF,
        XML,
        CSV,
        HWP,
        HWPX,
        Temp,
        Local,
        UNKNOWN = -1
    }    
    public enum ConversionType
    {
        Doc2Pdf,
        Pdf2Img,
        Img2Img
    }
    public interface IDocumentConverter
    {
        bool Convert(String sourcePath, String targetPath, ContentType sourceType);
    }
    public interface IDocumentConverterFactory
    {
        IDocumentConverter? GetConverter(ConversionType convType);
    }
    public class ConvertReport
    {
        public ConvertReport()
        {

        }

        public string ConvertType { get; set; } = string.Empty;
        public string ConvertTarget { get; set; } = string.Empty;
        public int CurrentCount   { get; set; } = 1;
        public int TotalCount     { get; set; } = 0;        
    }
    public class Util
    {
        public static void ReleaseObject(object? obj)
        {            
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj!);
                obj = null;
            }

            catch (Exception? ex)
            {
                obj = null;
                LoggingService.Logger("ReleaseComObject Error : " + ex.Message, LogLevel.Error);                
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }           
        }
        public static string GetTargetFileName(string ContentPath, string ContentName, ContentType ContentType)
        {
            string GetTargetFileNameStatus = string.Empty;

            try
            {
                string TempFileName     = Path.GetFileNameWithoutExtension(ContentName) + ".";
                TempFileName            += ContentType.ToString().ToLower();
                GetTargetFileNameStatus = Path.Combine(ContentPath, TempFileName);
            }
            catch (Exception ex)
            {                
                LoggingService.Logger("TarGetPath Error : " + ex.Message, LogLevel.Error);
            }
            return GetTargetFileNameStatus;
        }
        public static short GetFileExtension(string SourcePath, out ContentType FileExtension)
        {
            short GetFileExtnStatus = 0;
            FileExtension = ContentType.UNKNOWN;
            try
            {
                string tempFileExtn = Path.GetExtension(SourcePath);

                tempFileExtn = tempFileExtn.Replace(".", string.Empty);
                FileExtension = (ContentType)Enum.Parse(typeof(ContentType), tempFileExtn, true);
            }
            catch (Exception ex)
            {
                LoggingService.Logger("Get File Extension Error : " + ex.Message, LogLevel.Error);                
            }
            return GetFileExtnStatus;
        }
    }
    public class DocumentConverterFactory : IDocumentConverterFactory
    {
        public IDocumentConverter? GetConverter(ConversionType convType)
        {
            IDocumentConverter? converter = null;
            switch (convType)
            {
                case ConversionType.Doc2Pdf:
                    converter = new ConvDocToPdfWithMsOfficeService();
                    break;
                case ConversionType.Pdf2Img:
                case ConversionType.Img2Img:
                default:
                    break;
            }
            return converter;
        }
    }
    public class ConvDocToPdfWithMsOfficeService : IDocumentConverter
    {
        /// <summary>  
        /// Convert MSOffice file to PDF by calling required method  
        /// </summary>  
        /// <param name="sourcePath">MSOffice file path</param>  
        /// <param name="targetPath">Target PDF path</param>  
        /// <param name="sourceType">MSOffice file type</param>  
        /// <returns>error code : 0(sucess)/ -1 or errorcode (unknown error or failure)</returns>  
        public bool Convert(String sourcePath, String targetPath, ContentType sourceType)
        {
            bool IsConvertOK = false;
            if (sourceType == ContentType.PPT || sourceType == ContentType.PPTX)
            {
                IsConvertOK = PowerPointToPDF((Object)sourcePath, (Object)targetPath);
            }
            else IsConvertOK = false;
            return IsConvertOK;
        }
        /// <summary>  
        ///  Convert  powerPoint file to PDF by calling required method  
        /// </summary>  
        /// <param name="originalPptPath">file path</param>  
        /// <param name="pdfPath">Target PDF path</param>  
        /// <returns>error code : 0(sucess)/ -1 or errorcode (unknown error or failure)</returns>  
        public bool PowerPointToPDF(object originalPptPath, object pdfPath)
        {
            bool IsConvertOK = false;
            Microsoft.Office.Interop.PowerPoint.Application? PptApplication = null;
            Microsoft.Office.Interop.PowerPoint.Presentation? PptPresentation = null;
            object unknownType = Type.Missing;
            try
            {                
                //start power point   
                PptApplication = new Microsoft.Office.Interop.PowerPoint.Application();

                //open powerpoint password checked
                PptPresentation = PptApplication!.Presentations.Open($"{(string)originalPptPath}::xopen::", MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                //export PDF from PPT   
                if (PptPresentation != null)
                {
                    PptPresentation.ExportAsFixedFormat((string)pdfPath,
                                                         Microsoft.Office.Interop.PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                                                         Microsoft.Office.Interop.PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint,
                                                         MsoTriState.msoFalse,
                                                         Microsoft.Office.Interop.PowerPoint.PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
                                                         Microsoft.Office.Interop.PowerPoint.PpPrintOutputType.ppPrintOutputSlides,
                                                         MsoTriState.msoFalse, null,
                                                         Microsoft.Office.Interop.PowerPoint.PpPrintRangeType.ppPrintAll, string.Empty,
                                                         true, true, true, true, false, unknownType);
                    
                    IsConvertOK = true;
                }
                else
                {
                    LoggingService.Logger("Error occured for conversion of office PowerPoint to PDF", LogLevel.Error);
                    
                    IsConvertOK = false;
                }
            }
            catch (Exception ex)
            {
                LoggingService.Logger($"Error occured for conversion of office PowerPoint to PDF, Exception: {ex.Message} ", LogLevel.Error);                

                IsConvertOK = false;
            }
            finally
            {             
                // Close and release the Document object.  
                if (PptPresentation != null)
                {
                    PptPresentation.Close();
                    Util.ReleaseObject(PptPresentation);
                    PptPresentation = null;
                }

                // Quit Word and release the ApplicationClass object.  
                PptApplication!.Quit();
                Util.ReleaseObject(PptApplication);
                PptApplication = null;
            }
            return IsConvertOK;
        }        
    }

    public class ConvMSOfficeToDocService
    {
        private static string RootPath = string.Empty;
        private static string Language = string.Empty;
        private static void SetLanguage(string language) 
        {
            Language = language;
        }
        public static void SetRootPath(string path)
        {
            RootPath = path;
        }
        public static XpsDocument ConvertPowerPointToXps(string PptFilename, string XpsFilename)
        {
            var pptApp = new Microsoft.Office.Interop.PowerPoint.Application();

            var presentation = pptApp.Presentations.Open(PptFilename, MsoTriState.msoTrue, MsoTriState.msoFalse,
                MsoTriState.msoFalse);

            try
            {
                presentation.ExportAsFixedFormat(XpsFilename, Microsoft.Office.Interop.PowerPoint.PpFixedFormatType.ppFixedFormatTypeXPS);
            }
            catch (Exception ex)
            {
                LoggingService.Logger("Failed to export to XPS format: " + ex, LogLevel.Error);                
            }
            finally
            {
                presentation.Close();
                pptApp.Quit();
            }

            return new XpsDocument(XpsFilename, FileAccess.Read);
        }
        public static void ConvertPowerPointToImage(string SourcePath, string ExportPath)
        {
            try
            {
                //ppt를 이미지로 뽑아내는 기능
                var pptApp = new Microsoft.Office.Interop.PowerPoint.Application();

                var presentation = pptApp.Presentations.Open(SourcePath, MsoTriState.msoCTrue, MsoTriState.msoTriStateMixed, MsoTriState.msoFalse);
                int i = 0;
                foreach (Microsoft.Office.Interop.PowerPoint.Slide objSlide in presentation.Slides)
                {
                    //Names are generated based on timestamp.               
                    objSlide.Export(ExportPath + @"\Slide" + i + ".PNG", "PNG", 960, 960); //성공
                    i++;
                }
                presentation.Close();
                pptApp.Quit();
            }
            catch (Exception ex)
            {
                LoggingService.Logger("Failed to export to Image format: " + ex, LogLevel.Error);                
            }
        }     
        public static string? ConvertPowerPointToPDF(string ConvertType, string OrgFilePath, bool IsConvertPDF = true)
        {
            int InsertOffset = 0; //\\Temp 폴더 삽입 위치
            ContentType ConversionType;
            
            if (OrgFilePath.Contains("\\AppData\\"))
            {                 
                if (ConvertType == "PPTX" || ConvertType == "PPT")
                {
                    InsertOffset = 5;
                    ConversionType = ContentType.Temp;
                }
                else
                {
                    //PPXS 타입
                    return string.Empty;
                }
            }
            else
            {
                if (ConvertType == "PPTX")
                {
                    InsertOffset = 5;
                    ConversionType = ContentType.PPTX;
                }
                else if (ConvertType == "PPT")
                {
                    InsertOffset = 4;
                    ConversionType = ContentType.PPT;
                }
                else
                {
                    //PPXS 타입
                    return string.Empty;
                }
            }

            string? MakePDFSFilePath = OrgFilePath.Insert(OrgFilePath.LastIndexOf($"\\{ConversionType}\\") + InsertOffset, "\\PDFS");                        
            string? MakePDFSDir      = Path.GetDirectoryName(MakePDFSFilePath);
            string? ConvPdfFileName  = Path.ChangeExtension(OrgFilePath, ".pdf");
            string? OnlyPdfFileName  = Path.GetFileName(ConvPdfFileName);
           
            if (!Directory.Exists(MakePDFSDir))
            {
                Directory.CreateDirectory(MakePDFSDir!);
            }

            string? PdfFilePath = Path.Combine(MakePDFSDir!, OnlyPdfFileName);

            if (!File.Exists(PdfFilePath) && IsConvertPDF == true)
            {
                ConversionType = EnumConverterService.StringToEnum<ContentType>(ConvertType);
                ConvDocToPdfWithMsOfficeService DocToPdf = new ConvDocToPdfWithMsOfficeService();
                DocToPdf.Convert(OrgFilePath, PdfFilePath, ConversionType);                
            }
                     
            return PdfFilePath;
        }
        public static async Task<bool> ConvertPowerPointToPDFAll(IProgress<ConvertReport> ConvertProgressReport)
        {
            return await Task.Run(() =>
            {
                try
                {
                    ConvertReport ConvertReport = new ConvertReport();

                    List<string> ConvertTypes = new List<string>
                    {
                        "PPT",
                        "PPTX"
                    };

                    foreach (var ConvertType in ConvertTypes)
                    {
                        string? ExportPPTPath = Path.Combine(RootPath!, ConvertType);

                        if (!Directory.Exists(ExportPPTPath)) continue;

                        string[] PPTFiles = Directory.GetFiles(ExportPPTPath, $"*.{ConvertType}", SearchOption.AllDirectories);

                        ConvertReport.CurrentCount = 0;

                        foreach (var PPTFile in PPTFiles)
                        {
                            string? ExportPdfFilePath = ConvertPowerPointToPDF(ConvertType, PPTFile, true);

                            if (!File.Exists(ExportPdfFilePath))
                            {
                                string ConvertErrorReport;

                                if (Language == "Korean")
                                    ConvertErrorReport = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ConvertErrorReport", "NotPreview_KR.pdf");
                                else
                                    ConvertErrorReport = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ConvertErrorReport", "NotPreview_EN.pdf");

                                bool IsConvert = FileControlService.FileCopy(ConvertErrorReport, ExportPdfFilePath!);
                                if (IsConvert == false)
                                {
                                    LoggingService.Logger("PPT/PPTX to PDF Not Convert : " + ExportPdfFilePath, LogLevel.Warn);
                                }
                            }

                            ConvertReport.TotalCount = PPTFiles.Count();
                            ConvertReport.ConvertTarget = PPTFile;
                            ConvertReport.ConvertType = ConvertType;
                            ConvertProgressReport.Report(ConvertReport);
                            ConvertReport.CurrentCount++;
                        }
                    }
                }
                catch (TaskCanceledException TaskCancelEx)
                {
                    LoggingService.Logger("PPT/PPTX to PDF Converting Task Canceled Exception : " + TaskCancelEx.Message, LogLevel.Error);

                    return false;
                }
                catch (OperationCanceledException OperationCancelEx)
                {
                    LoggingService.Logger("PPT/PPTX to PDF Converting Operation Canceled Exception  : " + OperationCancelEx.Message, LogLevel.Error);

                    return false;
                }
                catch (Exception ex)
                {
                    LoggingService.Logger("PPT/PPTX to PDF Converting Error : " + ex.Message, LogLevel.Error);

                    return false;
                }
                return true;
            });
        }
        public static int CheckedNotExsistsPowerPointToPDFCount()
        {
            int NotExsistsCount = 0;

            List<string> ConvertTypes = new List<string>
            {
                "PPT",
                "PPTX"
            };

            foreach (var ConvertType in ConvertTypes)
            {                
                string? ExportPPTPath = Path.Combine(RootPath, ConvertType);

                if (Directory.Exists(ExportPPTPath) == false) continue;

                string[] PPTFiles = Directory.GetFiles(ExportPPTPath, $"*.{ConvertType}", SearchOption.AllDirectories);

                foreach (var PPTFile in PPTFiles)
                {
                    string? ExportPdfFilePath = ConvertPowerPointToPDF(ConvertType, PPTFile, false);

                    if (!File.Exists(ExportPdfFilePath))
                    {
                        NotExsistsCount = NotExsistsCount + 1;
                        LoggingService.Logger("HWP to PDF Not Convert : " + ExportPdfFilePath, LogLevel.Warn);
                    }
                }
            }

            return NotExsistsCount;
        }
        public static bool IsExsistsPowerPointDir()
        {
            bool ExsistsDirectory = false;

            List<string> ConvertTypes = new List<string>
            {
                "PPT",
                "PPTX"
            };

            foreach (var ConvertType in ConvertTypes)
            {                
                string? ExportPPTPath = Path.Combine(RootPath, ConvertType);

                if (Directory.Exists(ExportPPTPath) == true)
                {
                    ExsistsDirectory = true;
                    break;
                }
            }

            return ExsistsDirectory;
        }
        public static bool IsProtectedByPassword(string orgfilepath)
        {
            bool IsProtectedByPassword = false;
            Microsoft.Office.Interop.PowerPoint.Application? PptApplication = null;
            Microsoft.Office.Interop.PowerPoint.Presentation? PptPresentation = null;

            try
            {
                //start power point   
                PptApplication = new Microsoft.Office.Interop.PowerPoint.Application();

                //open powerpoint password checked
                PptPresentation = PptApplication!.Presentations.Open($"{(string)orgfilepath}::xopen::", MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                IsProtectedByPassword = false;
            }
            catch (Exception ex)
            {
                IsProtectedByPassword = true;

                LoggingService.Logger("PPT/PPTX Password Checke Error : " + ex.Message, LogLevel.Error);                
            }
            finally
            {
                // Close and release the Document object.  
                if (PptPresentation != null)
                {
                    PptPresentation.Close();
                    Util.ReleaseObject(PptPresentation);
                    PptPresentation = null;
                }
                else
                {
                    IsProtectedByPassword = true;
                }

                // Quit Word and release the ApplicationClass object.  
                PptApplication!.Quit();
                Util.ReleaseObject(PptApplication);
                PptApplication = null;
            }
            return IsProtectedByPassword;
        }
        public static bool IsPowerPointInstalled_V16()
        {
            // 검색할 레지스트리 경로 설정           
            List<string> RegistryPathList = new List<string>
            {
                @"SOFTWARE\WOW6432Node\Microsoft\Office", 
                @"SOFTWARE\Microsoft\Office"                
            };

            foreach(var RegistryPath in RegistryPathList)
            {
                // 검색할 버전 설정 (MS Office 2021 ,2019, 2016)
                const int officeVersion = 16;

                // 레지스트리 경로 설정
                RegistryKey? registryKey = Registry.LocalMachine.OpenSubKey(RegistryPath + @$"\{officeVersion}.0\PowerPoint\InstallRoot\");
                // 레지스트리 키가 있는지 확인
                if (registryKey != null)
                {
                    // PowerPoint.exe 파일이 있는지 확인
                    string? installRoot = registryKey.GetValue("Path") as string;

                    if (string.IsNullOrEmpty(installRoot))
                    {
                        break;
                    }

                    string powerPointPath = System.IO.Path.Combine(installRoot, "POWERPNT.EXE");
                    return System.IO.File.Exists(powerPointPath);
                }
            }            
            return false;
        }      
    }
    public class ConvHncToDocService
    {
        enum WN_Message
        {
            ToConvertPdfStartCount   = 2024,
            ToConvertPdfCurrentCount = 2025,
        }
        public static string LoadPath = string.Empty;
        private static ConvertReport ReportProgress = new ConvertReport();
        private static IProgress<ConvertReport> ?IConvertProgressReport;
        private static PresentationSource source;
        private static string RootPath = string.Empty;
        private static string Language = string.Empty;
        private static void SetLanguage(string language)
        {
            Language = language;
        }
        public static void SetRootPath(string path)
        {
            RootPath = path;
        }
        public static void SetParentHandle()
        {
            source = PresentationSource.FromVisual(Application.Current.MainWindow) as HwndSource;
        }
        public static async Task<bool> ConvertHWPToPDFAll(Progress<ConvertReport> ConvertProgressReport)
        {           
            return await Task.Run(async () =>
            {
                List<string> ConvertTypes = new List<string>
                {
                    "HWP",
                    "HWPX",
                };

                IConvertProgressReport = ConvertProgressReport;

                var hWnd_source = PresentationSource.FromVisual(Application.Current.MainWindow) as HwndSource;
                var hwnd_wndproc_hook = new HwndSourceHook(HwpToPdfConvertWndProc);
                hWnd_source!.AddHook(hwnd_wndproc_hook);
                                     
                try
                {
                    foreach (var ConvertType in ConvertTypes)
                    {
                        ReportProgress.ConvertType = ConvertType;

                        string? ExportHwpPath = Path.Combine(RootPath, ConvertType);
                        string? HwpToPdffile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "hwp_to_pdf.exe");

                        if (!Directory.Exists(ExportHwpPath)) continue;

                        ProcessStartInfo psi = new ProcessStartInfo
                        {
                            FileName = "\"" + HwpToPdffile + "\"",
                            Arguments = $"-hrp \"{ExportHwpPath}\" -phw \"{hWnd_source.Handle}\"",
                            Verb = "runas",
                            UseShellExecute = false,
                            CreateNoWindow = true,
                            RedirectStandardError = true,
                            RedirectStandardInput = true,
                            RedirectStandardOutput = true
                        };

                        Process proc = new Process
                        {
                            StartInfo = psi
                        };

                        proc.EnableRaisingEvents = true;
                        proc.OutputDataReceived += (sender, args) => Read_HwpToPdfConvertLogData(args.Data);
                        proc.ErrorDataReceived += (sender, args) => Read_HwpToPdfConvertLogData(args.Data);
                        proc.Exited += (sender, args) => HwpToPdfConvertExited();
                        proc.Start();
                        proc.BeginOutputReadLine();
                        proc.BeginErrorReadLine();

                        await proc.WaitForExitAsync();
                    }
                }
                catch (TaskCanceledException TaskCancelEx)
                {
                    LoggingService.Logger("HWP to PDF Converting Task Canceled Exception : " + TaskCancelEx.Message, LogLevel.Error);                    

                    return false;
                }
                catch (OperationCanceledException OperationCancelEx)
                {
                    LoggingService.Logger("HWP to PDF Converting Operation Canceled Exception  : " + OperationCancelEx.Message, LogLevel.Error);

                    return false;
                }
                catch (Exception ex)
                {
                    LoggingService.Logger("HWP to PDF Converting Error : " + ex.Message, LogLevel.Error);                    

                    return false;
                }
                
                hWnd_source.RemoveHook(HwpToPdfConvertWndProc);

                return true;
            }).ConfigureAwait(false);
        }
        private static IntPtr HwpToPdfConvertWndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {                       
            if (msg == (int)WN_Message.ToConvertPdfStartCount)
            {                
                ReportProgress.CurrentCount = 0;
                ReportProgress.TotalCount   = wParam.ToInt32();                
            }
            else if (msg == (int)WN_Message.ToConvertPdfCurrentCount)
            {
                ReportProgress.CurrentCount = wParam.ToInt32();
            }

            if(IConvertProgressReport != null)
            {
                IConvertProgressReport.Report(ReportProgress);
            }
          
            return IntPtr.Zero;
        }
        private static void HwpToPdfConvertExited()
        {
            if(ReportProgress != null)
            {
                LoggingService.Logger($"HWP to PDF Converting Exited :Converting Count " +
                    $"{ReportProgress.CurrentCount} / {ReportProgress.TotalCount} ", LogLevel.Info);         
            }                      
        }

        private static void Read_HwpToPdfConvertLogData(string? data)
        {
            if(data != null)
            {
                string[] parsing = data!.Split("|");
                if (parsing != null && parsing.Count() > 0 && ReportProgress != null)
                {
                    if (parsing[0] == "ToConvertPdfStartCount")
                    {
                        ReportProgress.CurrentCount = int.Parse(parsing[1]);
                        ReportProgress.TotalCount   = int.Parse(parsing[2]);
                    }
                    else
                    {
                        if(parsing.Count() > 1)
                        {
                            ReportProgress.CurrentCount = int.Parse(parsing[1]);
                        }
                    }

                    LoggingService.Logger($"{ReportProgress.ConvertType} to PDF Converting Exited :Converting Count " +
                        $"{ReportProgress.CurrentCount} / {ReportProgress.TotalCount} ", LogLevel.Info);
                }
            }            
        }

        public static string? ConvertHWPToPDF(string ConvertType, string OrgFilePath, bool IsConvertErrorReport = true)
        {
            string? PdfFilePath = "";
            int InsertOffset = 0;
            ContentType ConversionType = ContentType.HWP;
            if (OrgFilePath.Contains("\\AppData\\Local\\Temp"))
            {
                InsertOffset = 5;
                ConversionType = ContentType.Temp;
            }            
            else
            {
                if (ConvertType.Equals("HWPX", StringComparison.OrdinalIgnoreCase))
                {
                    InsertOffset = 5;
                    ConversionType = ContentType.HWPX;
                }
                else
                {
                    InsertOffset = 4;
                    ConversionType = ContentType.HWP;
                }              
            }
         
            try
            {
              
                string? MakePDFSFilePath = OrgFilePath.Insert(OrgFilePath.LastIndexOf($"\\{ConversionType}\\") + InsertOffset, "\\PDFS");
                string? MakePDFSDir      = Path.GetDirectoryName(MakePDFSFilePath);
                string? ConvPdfFileName  = Path.ChangeExtension(OrgFilePath, ".pdf");
                string? OnlyPdfFileName  = Path.GetFileName(ConvPdfFileName);

                if (!Directory.Exists(MakePDFSDir))
                {
                    Directory.CreateDirectory(MakePDFSDir!);
                }

                PdfFilePath = Path.Combine(MakePDFSDir!, OnlyPdfFileName);

                if (!File.Exists(PdfFilePath) && IsConvertErrorReport == true) // PDF가 파일이 존재 하지 않는지 한번더 체크
                {
                    if (Language == "Korean")
                        PdfFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ConvertErrorReport", "NotPreview_KR.pdf");
                    else
                        PdfFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ConvertErrorReport", "NotPreview_EN.pdf");
                }
            }
            catch(Exception ex)
            {
                LoggingService.Logger("HWP to PDF Convert Path Error : " + ex.Message, LogLevel.Warn);                
            }
        
            return PdfFilePath;
        }
        public static int CheckedNotExsistsHWPToPDFCount()
        {
            int NotExsistsCount = 0;
            List<string> ConvertTypes = new List<string>
            {
                "HWP",
                "HWPX",
            };

            foreach (var ConvertType in ConvertTypes)
            {
                string? ExportHwpPath = Path.Combine(RootPath, ConvertType);

                if (Directory.Exists(ExportHwpPath) == false) continue;

                string[] HWPFiles = Directory.GetFiles(ExportHwpPath, $"*.{ConvertType}", SearchOption.AllDirectories);

                foreach (var HWPFile in HWPFiles)
                {
                    string? ExportPdfFilePath = ConvertHWPToPDF(ConvertType, HWPFile, false);

                    if (!File.Exists(ExportPdfFilePath))
                    {
                        NotExsistsCount = NotExsistsCount + 1;
                        LoggingService.Logger("HWP(X) to PDF Not Convert : " + ExportPdfFilePath, LogLevel.Warn);

                    }
                }
            }

            return NotExsistsCount;
        }
        public static void WriteRegistryFilePathCheckerModule(bool IsSetForcedWriting = false)
        {
            bool IsHnCInstalled = ConvHncToDocService.IsHnCInstalled();

            if (IsHnCInstalled == false)
            {
                LoggingService.Logger("HWP is not installed.", LogLevel.Warn);
                
                return;
            }

            string CreateRegistryBase = @"SOFTWARE\HNC\HwpAutomation\Modules";
            string CheckerModule      = "FilePathCheckerModule";
            try
            {               
                using (var Module = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
                {
                    
                    bool ExsistValue = true;

                    using RegistryKey? OpenSubKey = Module.OpenSubKey(CreateRegistryBase);

                    if(OpenSubKey != null)
                    {
                        ExsistValue = OpenSubKey.GetValueNames().Contains(CheckerModule);

                        OpenSubKey!.Close();
                    }
                    else
                    {
                        ExsistValue = false;                      
                    }

                    if(ExsistValue == false || IsSetForcedWriting == true)
                    {
                        WriteRegistryFilePathCheckerModule(Module, CreateRegistryBase, CheckerModule);
                    }                                       
                }
            }
            catch (Exception ex)
            {
                LoggingService.Logger("Write Registry FilePathCheckerModule Error : " + ex.Message, LogLevel.Error);                
            }            
        }     
        private static void WriteRegistryFilePathCheckerModule(RegistryKey Module,string CreateRegistryBase, string CheckerModule)
        {
            if(Module != null)
            {
                RegistryKey CreateSubKey = Module.CreateSubKey(CreateRegistryBase);

                if(CreateSubKey != null)
                {
                    string FilePathCheckerModulePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"FilePathCheckerModuleExample.dll");

                    CreateSubKey.SetValue(CheckerModule, FilePathCheckerModulePath, RegistryValueKind.String);

                    CreateSubKey.Close();
                }                
            }            
        }

        public static bool ReadRegistryFilePathCheckerModule()
        {
            bool IsHnCInstalled = ConvHncToDocService.IsHnCInstalled();

            if (IsHnCInstalled == false)
            {
                LoggingService.Logger("HWP installation information is unknown.", LogLevel.Warn);                

                return false;
            }

            bool ExsistValue = true;

            try
            {              
                string CreateRegistryBase = @"SOFTWARE\HNC\HwpAutomation\Modules";
                string CheckerModule = string.Empty;
                

                using (var Module = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
                {
                    CheckerModule = "FilePathCheckerModule";
                    
                    using RegistryKey? OpenSubKey = Module.OpenSubKey(CreateRegistryBase);

                    if (OpenSubKey != null)
                    {
                        ExsistValue = OpenSubKey.GetValueNames().Contains(CheckerModule);

                        if(ExsistValue == true)
                        {
                            string ?CheckerModulePath = OpenSubKey.GetValue(CheckerModule) as string;

                            if(string.IsNullOrEmpty(CheckerModulePath) == false && File.Exists(CheckerModulePath) == true)
                            {
                                ExsistValue = true;
                            }
                            else
                            {
                                ExsistValue = false;
                            }
                        }

                        OpenSubKey!.Close();
                    }
                    else
                    {
                        ExsistValue = false;
                    }                                      
                }
            }
            catch (Exception ex)
            {
                ExsistValue = false;

                LoggingService.Logger("Read Registry FilePathCheckerModule Error : " + ex.Message, LogLevel.Error);                
            }
            return ExsistValue;
        }
        public static void Remove_gen_py()
        {
            try
            {
                //Hwp to PDF 컨버팅 안정화를 위한 gen_py 캐싱 폴더 삭제
                string LocalAppData = KnownFoldersService.GetPath(KnownFolder.LocalAppData);
                string gen_py = Path.Combine(LocalAppData, "Temp", "gen_py");
                if (Directory.Exists(gen_py))
                {
                    Directory.Delete(gen_py, true);
                }
            }
            catch(Exception ex)
            {
                LoggingService.Logger("gen_py directory delete fail : " + ex.Message, LogLevel.Error);                
            }          
        }

        public static bool IsExsistsHWPDir()
        {
            bool ExsistsDirectory = false;
            List<string> ConvertTypes = new List<string>
            {
                 "HWP",
                 "HWPX"
            };

            foreach (var ConvertType in ConvertTypes)
            {                
                string? ExportHncPath = Path.Combine(RootPath, ConvertType);

                if (Directory.Exists(ExportHncPath) == true)
                {
                    ExsistsDirectory = true;
                    break;
                }
            }

            return ExsistsDirectory;
        }
        public static bool IsHnCInstalled()
        {
            string HwpinstallPath = string.Empty;
            try
            {
                const string registryPath = @"hwp\DefaultIcon\";

                // 레지스트리 경로 설정
                RegistryKey? registryKey = Registry.ClassesRoot.OpenSubKey(registryPath);

                // 레지스트리 키가 있는지 확인
                if (registryKey == null)
                {
                    LoggingService.Logger("Not Founded Hwp installed registryPath : ", LogLevel.Warn);                    

                    return false;
                }

                // Hwp.exe 파일이 있는지 확인 
                HwpinstallPath = (registryKey.GetValue(null) as string)!;
                if (string.IsNullOrEmpty(HwpinstallPath))
                {
                    return false;
                }
            }
            catch(Exception ex) 
            {
                LoggingService.Logger("HWP installation check error : " + ex.Message, LogLevel.Warn);
                
                return false;
            }
            // 검색할 레지스트리 경로 설정
            
            return System.IO.File.Exists(HwpinstallPath);
        } 
        public static bool IsHwpToPdfConverterExsist()
        {
            bool IsHwpToPdfConverterExsist = false;

            string HwpToPdfConverter = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "hwp_to_pdf.exe");
            if (File.Exists(HwpToPdfConverter) == true)
            {
                IsHwpToPdfConverterExsist =  true;
            }
            else
            {
                IsHwpToPdfConverterExsist = false;
            }

            return IsHwpToPdfConverterExsist;
        }
    }
}