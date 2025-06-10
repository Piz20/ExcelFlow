using System;
using System.Diagnostics;
using System.IO;
using System.Windows;  // Important pour StartupEventArgs, ExitEventArgs
using WpfApplication = System.Windows.Application;  // alias pour lever l'ambiguïté avec Forms
using WpfMsgBox = System.Windows.MessageBox;   // alias pour MessageBox WPF

namespace ExcelFlow
{
    public partial class App : WpfApplication

    {
        private Process? _backendProcess; // ✅ nullable pour éviter les warnings

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            StartBackend();
        }

        private void StartBackend()
        {
            string backendPath = Path.Combine(Directory.GetCurrentDirectory(), "backend", "backend.dll"); // 🔁 adapte le nom ici si différent

            if (!File.Exists(backendPath))
            {
                WpfMsgBox.Show("Le fichier backend n'a pas été trouvé :\n" + backendPath);
                Shutdown();
                return;
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = $"\"{backendPath}\"",
                WorkingDirectory = Path.GetDirectoryName(backendPath),
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            _backendProcess = Process.Start(startInfo);

            if (_backendProcess != null)
            {
                _backendProcess.OutputDataReceived += (s, e) => Console.WriteLine(e.Data);
                _backendProcess.BeginOutputReadLine();
            }
            else
            {
                WpfMsgBox.Show("Impossible de démarrer le backend.");
                Shutdown();
            }

        }

        protected override void OnExit(ExitEventArgs e)
        {
            base.OnExit(e);

            if (_backendProcess != null && !_backendProcess.HasExited)
            {
                _backendProcess.Kill();
            }

        }
    }
}
