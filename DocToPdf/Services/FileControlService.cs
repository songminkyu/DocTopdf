﻿using System.Diagnostics;
using System.IO;

namespace DocToPdf.Services
{
    public class FileControlService
    {
        public static bool FileCopy(string Source, string Destination)
        {
            try
            {
                string SourceFilePath = Source;
                string DestFilePath = Destination;
                byte[] buffer = new byte[1024 * 1024]; // 1MB buffer
                var DestFileinfo = new FileInfo(DestFilePath);
                using (FileStream SourceStream = new FileStream(SourceFilePath, FileMode.Open, FileAccess.Read))
                {
                    long fileLength = SourceStream.Length;
                    if (File.Exists(DestFilePath)) { File.Delete(DestFilePath); }

                    if (!Directory.Exists(DestFileinfo.DirectoryName))
                    {
                        Directory.CreateDirectory(DestFileinfo.DirectoryName);
                    }

                    using (FileStream DestStream = new FileStream(DestFilePath, FileMode.CreateNew, FileAccess.Write))
                    {
                        long currentCopySize = 0;
                        int currentBlockSize = 0;
                        while ((currentBlockSize = SourceStream.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            currentCopySize += currentBlockSize;
                            DestStream.Write(buffer, 0, currentBlockSize);
                        }
                    }
                }
            }
            catch (Exception ex)
            {               
                return false;
            }

            return true;
        }
        public static void DocumentProcessKill()
        {
            string[] processNames = new string[] { "hwp_to_pdf.exe", "HWP.exe", "HwpConverter.exe", "POWERPNT.exe" };

            foreach (var processName in processNames)
            {
                try
                {
                    Process killProcess = new Process();
                    ProcessStartInfo taskkillStartInfo = new ProcessStartInfo
                    {
                        FileName = "taskkill",
                        Arguments = $"/F /IM {processName}",
                        Verb = "runas",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };

                    killProcess.StartInfo = taskkillStartInfo;
                    killProcess.Start();
                    killProcess.WaitForExit();
                }
                catch (Exception ex)
                {
                    
                }
            }
        }
        public static bool OpenDocumentProcesses()
        {
            string[] ProcessNames = new string[] { "POWERPNT", "HWP" };

            foreach (string ProcessName in ProcessNames)
            {
                Process[] PowerPointProcesses = Process.GetProcessesByName(ProcessName);
                if (PowerPointProcesses.Length > 0)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
