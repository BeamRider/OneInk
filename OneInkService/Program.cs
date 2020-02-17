using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.InteropServices;
using LibOneInk;

namespace OneInkService
{
    static class Program
    {
        private static readonly log4net.ILog Logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private const string MutexBaseId = @"Global\{8A6A585D-4EE5-4E96-97AF-7E3E4A85E04D}";

        private const string WillFileNameFilter = "*.will";

        private static string _notebookName = null;
        private static string _sectionName = null;
        private static string _dir = null;
        private static float _scaleFactor = 1.0f;
        private static float _pressureRatio = 0.25f;

        static void Main(string[] args)
        {
            int period = 10000;

            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);

            HandlerRoutine hr = new HandlerRoutine(ConsoleCtrlCheck);
            SetConsoleCtrlHandler(hr, true);
            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i];
                if (arg.Equals("--notebook"))
                {
                    i++;
                    _notebookName = args[i];
                }
                else if (arg.Equals("--section"))
                {
                    i++;
                    _sectionName = args[i];
                }
                else if (arg.Equals("--dir"))
                {
                    i++;
                    _dir = args[i];
                }
                else if (arg.Equals("--period"))
                {
                    i++;
                    int.TryParse(args[i], out period);
                }
                else if (arg.Equals("--scale"))
                {
                    i++;
                    float.TryParse(args[i], out _scaleFactor);
                }
                else if (arg.Equals("--pressureratio"))
                {
                    i++;
                    float.TryParse(args[i], out _pressureRatio);
                }
                else if (arg.Equals("--waitdebugger"))
                {
                    Logger.Warn("Waiting for debugger to attach");
                    while (!Debugger.IsAttached)
                    {
                        Thread.Sleep(100);
                    }
                    Logger.Warn("Debugger attached");
                }

            }

            SingleInstanceApplicationLock appLock = new SingleInstanceApplicationLock(MutexBaseId + "{" + _dir.GetHashCode() + "}");
            if (!appLock.TryAcquireExclusiveLock())
            {
                Console.Error.WriteLine("OneInkService is already running");
                return;
            }

            Logger.Info($"Notebook: {_notebookName}");
            Logger.Info($"Section: {_sectionName}");
            Logger.Info($"Directory: {_dir}");

            if ((_notebookName == null) || (_sectionName == null) || (_dir == null))
                return;

            Task.Run(async () =>
            {
                while (true)
                {
                    try
                    {
                        foreach (string file in Directory.EnumerateFiles(_dir, WillFileNameFilter, SearchOption.TopDirectoryOnly))
                        {
                            Logger.Info($"Will file detected on start {file}");
                            await ImportWillFile(file);
                        }
                        await Task.Delay(period);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error(ex);
                    }
                }
            });

            Thread.Sleep(Timeout.Infinite);
        }

        private static bool ConsoleCtrlCheck(CtrlTypes CtrlType)
        {
            Logger.Info($"{CtrlType} Event detected, finalizing");

            // be absolutely sure that OneNote has been disposed
            // OneNote2016 helper process will not close otherwise
            GC.Collect();
            GC.WaitForPendingFinalizers();
            return true;
        }

        private static async Task ImportWillFile(string path)
        {
            using (OneNote2016 oneNote = new OneNote2016())
            {
                await oneNote.SetNotebook(_notebookName);
                if (oneNote.NotebookId == null)
                {
                    Logger.Error($"Unable to find Notebook {_notebookName}");
                    return;
                }

                await oneNote.SetSection(_sectionName);
                if (oneNote.SectionId == null)
                {
                    Logger.Error($"Unable to find Section {_sectionName}");
                    return;
                }

                try
                {
                    WillConverter will = new WillConverter(path)
                    {
                        ScaleFactor = _scaleFactor,
                        PressureRatio = _pressureRatio
                    };

                    FileInfo fi = new FileInfo(path);
                    var oneNotePage = await oneNote.CreatePage(fi.Name.Substring(0, fi.Name.Length - fi.Extension.Length));

                    foreach (var group in will.Groups)
                    {
                        await oneNote.AddStrokeGroup(oneNotePage, group);
                    }

                    Logger.Info($"Updating page");
                    await oneNote.UpdatePageContent(oneNotePage);
                    File.Delete(path);
                }
                catch (Exception ex)
                {
                    Logger.Error(ex);
                }
            }
        }

        #region unmanaged

        /// <summary>
        /// This function sets the handler for kill events.
        /// </summary>
        /// <param name="Handler"></param>
        /// <param name="Add"></param>
        /// <returns></returns>
        [DllImport("Kernel32")]
        public static extern bool SetConsoleCtrlHandler(HandlerRoutine Handler, bool Add);

        //delegate type to be used of the handler routine
        public delegate bool HandlerRoutine(CtrlTypes CtrlType);

        // control messages
        public enum CtrlTypes
        {
            CTRL_C_EVENT = 0,
            CTRL_BREAK_EVENT,
            CTRL_CLOSE_EVENT,
            CTRL_LOGOFF_EVENT = 5,
            CTRL_SHUTDOWN_EVENT
        }

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;
        #endregion

    }
}
