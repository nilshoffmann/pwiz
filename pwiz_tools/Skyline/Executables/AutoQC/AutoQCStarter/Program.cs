using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace AutoQCStarter
{
    static class Program
    {
        private const string APP_NAME = "AutoQCStarter";
        private const string PUBLISHER_NAME = "University of Washington";
        private static string _autoqcName;
        private static string _autoqcAppPath;

        private static readonly string LOG_FILE = $"{APP_NAME}.log";
        
        private static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += Application_ThreadException;
            // Handle exceptions on the non-UI thread. 
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

            if (!InitLogging()) return;

            // https://saebamini.com/Allowing-only-one-instance-of-a-C-app-to-run/
            using (var mutex = new Mutex(false, $"{PUBLISHER_NAME} {APP_NAME}"))
            {
                if (!mutex.WaitOne(TimeSpan.Zero))
                {
                    ShowError($"{APP_NAME} is already running.");
                    return;
                }

                Log($"Starting {APP_NAME}...");

                try
                {
                    _autoqcAppPath = GetAutoQcPath(args);
                }
                catch (ArgumentException e)
                {
                    Log(e.Message);
                    Log(e.StackTrace);
                    ShowError(e.Message);
                    // Delete shortcut in the Startup folder, if it exists
                    DeleteShortcut();

                    mutex.ReleaseMutex();
                    return;
                }
                
                var stateRunning = false;
                while (true)
                {
                    if (!StartAutoQc(ref stateRunning))
                    {
                        var err = $"{_autoqcAppPath} no longer exists. Stopping.";
                        Log(err);
                        ShowError(err);
                        break;
                    }
                    Thread.Sleep(TimeSpan.FromMinutes(1)); // TODO: Change to 5 minutes?
                }
                mutex.ReleaseMutex();
            }
        }

        private static string GetAutoQcPath(IReadOnlyList<string> args)
        {
            if (args.Count == 0)
            {
                throw new ArgumentException("No arguments given.  Require one of \"daily\", \"release\" or path to AutoQC.exe or AutoQC-daily.exe.");
            }

            var isDaily = args.Count > 0 && args[0].Equals("daily", StringComparison.InvariantCultureIgnoreCase);
            var isRelease = args.Count > 0 && args[0].Equals("release", StringComparison.InvariantCultureIgnoreCase);
            var hasPath = args.Count > 0 && (!isDaily && !isRelease);

            if (hasPath)
            {
                // Assume path to the AutoQC Loader exe is given
                var path = args[0].Trim();
                if (!File.Exists(path))
                {
                    throw new ArgumentException($"Given path to AutoQC Loader executable does not exist: {path}");     
                }

                if(!path.EndsWith("AutoQC.exe") && !path.EndsWith("AutoQC-daily.exe"))
                {
                    throw new ArgumentException($"Given path is not to a AutoQC.exe or AutoQC-daily.exe: {path}");
                }

                return path;
            }

            _autoqcName = isDaily ? "AutoQC-daily" : "AutoQC";
            var apprefsName = _autoqcName + ".appref-ms";

            var appRefsPath = GetAppRefsPath(apprefsName, _autoqcName, PUBLISHER_NAME);
            if (appRefsPath == null)
            {
                throw new ArgumentException($"Could not find path to {apprefsName}.");
            }

            return appRefsPath;
        }

        private static void DeleteShortcut()
        {
            var startupDir = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            var shortcutPath = Path.Combine(startupDir, $"{APP_NAME}.lnk");
            if (File.Exists(shortcutPath))
            {
                Log($"Deleting shortcut {shortcutPath}");
                try
                {
                    File.Delete(shortcutPath);
                }
                catch (Exception e)
                {
                    Log($"Unable to delete {shortcutPath}");
                    Log(e.Message);
                }
            }
        }

        private static bool InitLogging()
        {
            var logLocation = GetLogLocation();
            var logFile = Path.Combine(logLocation, LOG_FILE);

            try
            {
                using (new FileStream(Path.Combine(logLocation, LOG_FILE), FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                }
            }
            catch (Exception e)
            {
                ShowError($"Cannot create or write to log file: {logFile}." + Environment.NewLine + $"Error was: {e.Message}");
                return false;
            }

            Trace.Listeners.Add(new TextWriterTraceListener(Path.Combine(logLocation, LOG_FILE), $"{APP_NAME} Log"));
            Trace.AutoFlush = true;
            return true;
        }

        private static string GetLogLocation()
        {
            // Why use CodeBase instead of Location?
            // CodeBase: The location of the assembly as specified originally (https://docs.microsoft.com/en-us/dotnet/api/system.reflection.assembly.codebase?)
            // Location: The location of the loaded file that contains the manifest. If the loaded file was shadow-copied, the location is that of the file after being shadow-copied (https://docs.microsoft.com/en-us/dotnet/api/system.reflection.assembly.location?)
            // Using Location can be a problem in some unit testing scenarios (https://corengen.wordpress.com/2011/08/03/assembly-location-and-codebase/)
            var file = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;

            // How to convert CodeBase to filesystem path: https://stackoverflow.com/questions/4107625/how-can-i-convert-assembly-codebase-into-a-filesystem-path-in-c
            // The code below is copied from the SkylineNightlyShim project
            if (file.StartsWith(@"file:"))
            {
                file = file.Substring(5);
            }
            while (file.StartsWith(@"/"))
            {
                file = file.Substring(1);
            }
            return Path.GetDirectoryName(file);
        }

        private static string GetAppRefsPath(string apprefsName, string productName, string publisherName)
        {
            var allProgramsPath = Environment.GetFolderPath(Environment.SpecialFolder.Programs);
            var appRefDir = Path.Combine(allProgramsPath, publisherName); //e.g. %APPDATA%\Microsoft\Windows\Start Menu\Programs\University of Washington

            Log($"Looking for application reference {apprefsName} in {appRefDir}.");
            if (!Directory.Exists(appRefDir))
            {
                appRefDir = Path.Combine(allProgramsPath, productName);//e.g. %APPDATA%\Microsoft\Windows\Start Menu\Programs\AutoQC
                Log($"Looking for application reference {apprefsName} in {appRefDir}.");
            }

            if (!Directory.Exists(appRefDir))
            {
                Log($"Could not find location of application reference {apprefsName}.");
            }
            else
            {
                var appRefPath = Path.Combine(appRefDir, apprefsName);
                if (File.Exists(appRefPath))
                {
                    return appRefPath;
                }
            }
            Log($"Application reference {apprefsName} does not exit in {appRefDir}.");
            return null;
        }

        private static bool StartAutoQc(ref bool stateRunning)
        {
            if (Process.GetProcessesByName(_autoqcName).Length == 0)
            {
                Log($"Starting {_autoqcName}.");

                if (!File.Exists(_autoqcAppPath))
                {
                    return false;
                }

                // Run AutoQC.appref-ms
                var autoQc = new Process
                {
                    StartInfo =
                    {
                        UseShellExecute =
                            true, // Set to true otherwise there is an exception: The specified executable is not a valid application for this OS platform.
                        FileName = _autoqcAppPath,
                        CreateNoWindow = true
                    }
                };

                autoQc.Start();
            }

            if (!stateRunning)
            {
                Log($"{_autoqcName} is running.");
                stateRunning = true;
            }

            return true;
        }

        private static void ShowError(string message)
        {
            MessageBox.Show(message, $"{APP_NAME} Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void Log(string message)
        {
            Trace.WriteLine($"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}: {message}");
        }

        private static void Application_ThreadException(Object sender, ThreadExceptionEventArgs e)
        {
            Log($"Unhandled exception on UI thread: {e.Exception}");
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            try
            {
                Log($"{APP_NAME} encountered an unexpected error.");
                Log(((Exception)e.ExceptionObject).Message);
                ShowError($"{APP_NAME} encountered an unexpected error. " + Environment.NewLine + $"Error was: {((Exception)e.ExceptionObject).Message}");
            }
            finally
            {
                Application.Exit();
            }
        }
    }
}
