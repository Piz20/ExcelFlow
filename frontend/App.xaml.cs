using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using WpfApplication = System.Windows.Application;
using WpfMsgBox = System.Windows.MessageBox;

namespace ExcelFlow
{
    public partial class App : WpfApplication
    {
        private Process? _backendProcess;
        private static readonly string LogFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app.log");

        public App()
        {
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            Log("Application démarrée");
            StartBackend();
        }

        private void StartBackend()
        {
            string backendPath = Path.Combine(Directory.GetCurrentDirectory(), "backend", "backend.exe");  // ✅ .exe au lieu de .dll

            if (!File.Exists(backendPath))
            {
                string msg = "Le fichier backend n'a pas été trouvé :\n" + backendPath;
                Log(msg);
                WpfMsgBox.Show(msg);
                Shutdown();
                return;
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = backendPath,        // exécutable directement
                Arguments = "",                // ou les arguments si besoin
                WorkingDirectory = Path.GetDirectoryName(backendPath),
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };


            try
            {
                _backendProcess = Process.Start(startInfo);

                if (_backendProcess != null)
                {
                    Log("Backend démarré.");
                    _backendProcess.OutputDataReceived += (s, e) =>
                    {
                        if (!string.IsNullOrWhiteSpace(e.Data))
                            Log("[Backend] " + e.Data);
                    };
                    _backendProcess.BeginOutputReadLine();
                }
                else
                {
                    Log("Échec du démarrage du backend.");
                    WpfMsgBox.Show("Impossible de démarrer le backend.");
                    Shutdown();
                }
            }
            catch (Exception ex)
            {
                Log("Erreur lors du démarrage du backend : " + ex);
                WpfMsgBox.Show("Erreur au lancement du backend :\n" + ex.Message);
                Shutdown();
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            Log("Application en cours de fermeture...");
            if (_backendProcess != null && !_backendProcess.HasExited)
            {
                try
                {
                    _backendProcess.Kill();
                    Log("Backend arrêté proprement.");
                }
                catch (Exception ex)
                {
                    Log("Erreur à l'arrêt du backend : " + ex);
                }
            }

            Log("Application terminée.");
            base.OnExit(e);
        }

        private static void Log(string message)
        {
            try
            {
                string logLine = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
                File.AppendAllText(LogFilePath, logLine + Environment.NewLine);
            }
            catch
            {
                // en cas d'échec d'écriture (verrouillage, droits), on évite une boucle d'erreur
            }
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            Log("Exception non gérée : " + e.Exception);
            WpfMsgBox.Show("Erreur critique :\n" + e.Exception.Message);
            e.Handled = true;
        }
    }
}
