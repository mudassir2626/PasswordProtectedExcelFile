using System;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.IO;
using System.Diagnostics;
using System.Threading;

namespace PasswordProtectedZipFile
{
    public class Program
    {
        public static string path = @"";
        private static void FileSystemWatcher_Created(object sender, FileSystemEventArgs e )

        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();           
            if (e.Name.Contains("D_File_Export"))
            {
                var filename = path + e.Name;
                Console.WriteLine("File created: {0}", e.Name);
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(e.FullPath);
                xlWorkbook.Password = "123";
                xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.Quit();
                Console.WriteLine("Password-protected ZIP file created successfully.");
                stopWatch.Stop();

                TimeSpan ts = stopWatch.Elapsed;

                // Format and display the TimeSpan value.
                string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                    ts.Hours, ts.Minutes, ts.Seconds,
                    ts.Milliseconds);
                Console.WriteLine("RunTime " + elapsedTime);

                MonitorDirectory(path);
                Console.WriteLine("waiting ... ");
                Console.ReadKey();
              
            }
            else
            {
                Console.WriteLine("Wrong File :{0}," + e.Name);
            }
        

        }
        public static void Main(string[] args)
        {                    
            MonitorDirectory(path);
            Console.WriteLine("waiting ... ");
            Console.ReadKey();
        }
        private static void MonitorDirectory(string path)
        {
            FileSystemWatcher fileSystemWatcher = new FileSystemWatcher();
            fileSystemWatcher.Path = path;
            fileSystemWatcher.Created += FileSystemWatcher_Created;
            fileSystemWatcher.EnableRaisingEvents = true;
        }
    }   
}

