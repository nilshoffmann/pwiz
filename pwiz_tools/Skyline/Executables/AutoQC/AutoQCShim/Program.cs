using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace AutoQCShim
{
    internal class Program
    {
        private const bool DAILY = false;
        private const string AUTOQC_NAME = DAILY ? "AutoQC-daily" : "AutoQC";
        private const string APP_NAME = DAILY ? "AutoQCShim-daily" : "AutoQCShim";
        private const string PUBLISHER_NAME = "University of Washington";
        private const string APPREF_NAME = AUTOQC_NAME + ".appref-ms";

        private static readonly string LOG_FILE = $"{APP_NAME}.log";
        private static string _autoqcAppPath;
        private static bool _logRunningMsg = true;
        
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += Application_ThreadException;

            // Handle exceptions on the non-UI thread. 
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;


            // https://stackoverflow.com/questions/3571627/show-hide-the-console-window-of-a-c-sharp-console-application
            // Look at link above if we want to make this a console application and want to hide the console

            // https://www.dotnetcurry.com/ShowArticle.aspx?ID=150
            // https://docs.microsoft.com/en-us/dotnet/api/system.threading.mutex.-ctor?view=netframework-4.7.2#System_Threading_Mutex__ctor_System_Boolean_System_String_System_Boolean__System_Security_AccessControl_MutexSecurity_
            // https://saebamini.com/Allowing-only-one-instance-of-a-C-app-to-run/
            using (var mutex = new Mutex(false, $"{PUBLISHER_NAME} {APP_NAME}"))
            {
                if (!mutex.WaitOne(TimeSpan.Zero))
                {
                    ShowError($"{APP_NAME} is already running.");
                    return;
                }

                Log("Starting AutoQCShim...");
                if (!SetWorkingDirectory())
                {
                    mutex.ReleaseMutex();
                    ShowError("Could not set working directory. Stopping.");
                    return;
                }

                _autoqcAppPath = GetAppRefsPath(APPREF_NAME, AUTOQC_NAME, PUBLISHER_NAME) ?? 
                                 GetAutoQcExePath(AUTOQC_NAME); // Look for the AutoQC.exe if the appref-ms file was not found (e.g. for an unplugged install)

                if (_autoqcAppPath == null)
                {
                    var err = $"Cannot find path to {APPREF_NAME} or {AUTOQC_NAME}.exe. Stopping.";
                    Log(err);
                    ShowError(err);

                    mutex.ReleaseMutex();
                    return;
                }

                while (true)
                {
                    if (File.Exists(_autoqcAppPath))
                    {
                        StartAutoQc();
                        Thread.Sleep(TimeSpan.FromMinutes(1)); // TODO: Change to 5 minutes?
                    }
                    else
                    {
                        var err = $"{_autoqcAppPath} does not exist. Stopping.";
                        Log(err);
                        ShowError(err);
                        break;
                    }
                }
                mutex.ReleaseMutex();
            }
        }

        private static void ShowError(string message)
        {
            MessageBox.Show(message, "AutoQCShim Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            // Console.WriteLine(message);
        }

        private static bool SetWorkingDirectory()
        {
            // Set the working directory to where the .exe is so that the log gets written there.
            var exeLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var currentWd = Directory.GetCurrentDirectory();

            Log($"Exe location: {exeLocation}");
            Log($"Current Working Directory: {currentWd}");

            var dirPath = Path.GetDirectoryName(exeLocation);
            if (!string.IsNullOrEmpty(dirPath))
            {
                if (string.Compare(Path.GetFullPath(dirPath).TrimEnd('\\'), Path.GetFullPath(currentWd).TrimEnd('\\'),
                    StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    Log($"Current working directory: {currentWd} is the same ad exe location: {dirPath}");
                }
                else
                {
                    Log($"Setting working directory to {dirPath}");
                    Directory.SetCurrentDirectory(dirPath);
                }
            }
            else
            {
                return false;
            }
            return true;
        }

        public static void StartAutoQc()
        {
            // https://stackoverflow.com/questions/4722198/checking-if-my-windows-application-is-running
            if (Process.GetProcessesByName(AUTOQC_NAME).Length == 0)
            {
                Log($"Starting {AUTOQC_NAME}.");
                // Invoke AutoQC.appref-ms OR AutoQC.exe
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
                //_logRunningMsg = true;
            }
            else
            {
                if (_logRunningMsg)
                {
                    Log($"{AUTOQC_NAME} is running.");
                }

                _logRunningMsg = false;
            }
        }

        private static string GetAutoQcExePath(string productName)
        {
            var location = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var dirPath = Path.GetDirectoryName(location);
            if (!string.IsNullOrEmpty(dirPath))
            {
                var autoqcExePath = Path.Combine(dirPath, $"{AUTOQC_NAME}.exe");
                Log($"Looking for {AUTOQC_NAME}.exe in {dirPath}.");
                if (File.Exists(autoqcExePath))
                {
                    return autoqcExePath;
                }
            }

            return null;
        }

        private static string GetAppRefsPath(string apprefsName, string productName, string publisherName)
        {
            var allProgramsPath = Environment.GetFolderPath(Environment.SpecialFolder.Programs);
            var appRefDir = Path.Combine(allProgramsPath, publisherName);

            Log($"Looking for application reference {apprefsName} in {appRefDir}.");
            if (!Directory.Exists(appRefDir))
            {
                Log($"Looking for application reference {apprefsName} in {appRefDir}.");
                appRefDir = Path.Combine(allProgramsPath, productName);
            }

            if (Directory.Exists(appRefDir))
            {
                var appRefPath = Path.Combine(appRefDir, apprefsName);
                if (File.Exists(appRefPath))
                {
                    return appRefPath;
                }

                Log($"Application reference {appRefPath} does not exit in {appRefDir}.");
            }
            else
            {
                Log($"Could not find location of application reference {apprefsName}.");
            }

            return null;
        }

        private static void Log(string message)
        {
            var now = DateTime.Now.ToLocalTime();
            try
            {
                using (var w = File.AppendText(LOG_FILE))
                {
                    w.WriteLine("{0} {1}: {2}", now.ToShortDateString(), now.ToShortTimeString(), message);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
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

                var exeLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
                MessageBox.Show($"{APP_NAME} encountered an unexpected error. " +
                                $"Error details may be found in the {LOG_FILE} file in this directory: {Path.GetDirectoryName(exeLocation)}"
                );
            }
            finally
            {
                Application.Exit();
            }
        }
    }
}
