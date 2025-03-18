using AudioSwitcher.AudioApi;
using AudioSwitcher.AudioApi.CoreAudio;
using Microsoft.Win32;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace Microchange
{
    class Program
    {
        private static NotifyIcon _notifyIcon;
        private static CoreAudioController _audioController;
        private static readonly string AppName = "FullMicrochangeCS";
        private static readonly string Version = "2.0.2c";
        private static readonly string LogFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "FullMicrochangeCS_error_log.txt");
        private static readonly string StartupKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private static bool _isRussian;

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Определяем локаль
            _isRussian = System.Globalization.CultureInfo.CurrentCulture.Name.StartsWith("ru", StringComparison.OrdinalIgnoreCase);

            try
            {
                // Инициализация контроллера аудиоустройств
                _audioController = new CoreAudioController();

                // Создание иконки в трее
                _notifyIcon = new NotifyIcon
                {
                    Icon = FullMicrochangeCS.Properties.Resources.FullMicrochange,
                    Visible = true,
                    Text = AppName
                };
                _notifyIcon.Click += (s, e) => UpdateTrayMenu();

                // Обновляем меню при старте
                UpdateTrayMenu();

                // Запускаем цикл сообщений без формы
                Application.Run(new TrayApplicationContext());
            }
            catch (Exception ex)
            {
                LogError($"Initialization error: {ex}");
                _notifyIcon?.Dispose();
                Environment.Exit(1);
            }
        }

        private static void UpdateTrayMenu()
        {
            try
            {
                var contextMenu = new ContextMenuStrip();

                // Проверка наличия устройств
                var playbackDevices = _audioController.GetPlaybackDevices()
                    .Where(d => d.State == DeviceState.Active)
                    .ToList();
                var recordingDevices = _audioController.GetCaptureDevices()
                    .Where(d => d.State == DeviceState.Active)
                    .ToList();

                if (!playbackDevices.Any() && !recordingDevices.Any())
                {
                    contextMenu.Items.Add(new ToolStripMenuItem(_isRussian ? "Нет активных устройств" : "No active devices"));
                    _notifyIcon.ContextMenuStrip = contextMenu;
                    return;
                }

                // Устройства воспроизведения (Playback)
                var playbackItems = playbackDevices.Select(device => new ToolStripMenuItem
                {
                    Text = device.Name,
                    Checked = device.IsDefaultDevice,
                    Tag = device,
                    CheckOnClick = true
                }).ToArray();

                foreach (var item in playbackItems)
                {
                    item.Click += (s, e) =>
                    {
                        var selectedDevice = (CoreAudioDevice)item.Tag;
                        selectedDevice.SetAsDefault();
                        UpdateTrayMenu();
                    };
                }
                contextMenu.Items.Add(new ToolStripMenuItem(_isRussian ? "Устройства воспроизведения" : "Playback Devices", null, playbackItems));

                // Устройства записи (Recording)
                var recordingItems = recordingDevices.Select(device => new ToolStripMenuItem
                {
                    Text = device.Name,
                    Checked = device.IsDefaultDevice,
                    Tag = device,
                    CheckOnClick = true
                }).ToArray();

                foreach (var item in recordingItems)
                {
                    item.Click += (s, e) =>
                    {
                        var selectedDevice = (CoreAudioDevice)item.Tag;
                        selectedDevice.SetAsDefault();
                        UpdateTrayMenu();
                    };
                }
                contextMenu.Items.Add(new ToolStripMenuItem(_isRussian ? "Устройства записи" : "Recording Devices", null, recordingItems));

                // Разделитель
                contextMenu.Items.Add(new ToolStripSeparator());

                // Автозапуск
                var autostartItem = new ToolStripMenuItem(_isRussian ? "Автозапуск с системой" : "Startup with system")
                {
                    Checked = IsStartupEnabled(),
                    CheckOnClick = true
                };
                autostartItem.Click += (s, e) => ToggleStartup();
                contextMenu.Items.Add(autostartItem);

                // Разделитель
                contextMenu.Items.Add(new ToolStripSeparator());

                // Закрыть
                contextMenu.Items.Add(new ToolStripMenuItem(_isRussian ? "Закрыть программу" : "Close application", null, (s, e) =>
                {
                    _notifyIcon.Visible = false;
                    _notifyIcon.Dispose();
                    Application.Exit();
                }));

                // О программе
                var aboutItem = new ToolStripMenuItem(_isRussian ? "О программе" : "About");
                aboutItem.Click += (s, e) => ShowAboutDialog();
                contextMenu.Items.Add(aboutItem);

                _notifyIcon.ContextMenuStrip = contextMenu;

            }
            catch (Exception ex)
            {
                LogError($"Error updating tray menu: {ex}");
            }
        }

        private static void ShowAboutDialog()
        {
            try
            {
                string title = _isRussian ? "О программе" : "About";
                string message = _isRussian
                    ? $"FullMicrochangeCS\n\nТрей-ориентированная программа для изменения стандартных динамиков и микрофона в системе.\n\nАвтор: Dark Hacker (MooDuck Games)\n\nВерсия: {Version}\n\nКонтакты:\n@PythonistHarry (Telegram)\ndarkhacker (Discord)"
                    : $"FullMicrochangeCS\n\nTray-oriented program for changing default speakers and microphone in the system.\n\nAuthor: Dark Hacker (MooDuck Games)\n\nVersion: {Version}\n\nContacts:\n@PythonistHarry (Telegram)\ndarkhacker (Discord)";

                MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                LogError($"Error showing about dialog: {ex}");
            }
        }

        private static void CheckAndSetupStartup()
        {
            try
            {
                if (IsStartupEnabled()) return;

                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(StartupKeyPath, true))
                {
                    key?.SetValue(AppName, $"\"{Assembly.GetExecutingAssembly().Location}\"", RegistryValueKind.String);
                }
            }
            catch (Exception ex)
            {
                LogError($"Error setting up startup: {ex}");
            }
        }

        private static bool IsStartupEnabled()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(StartupKeyPath, false))
                {
                    if (key == null) return false;
                    string currentValue = key.GetValue(AppName) as string;
                    if (string.IsNullOrEmpty(currentValue)) return false;
                    string currentExecutablePath = Assembly.GetExecutingAssembly().Location;
                    return currentValue.Equals($"\"{currentExecutablePath}\"", StringComparison.OrdinalIgnoreCase);
                }
            }
            catch (Exception ex)
            {
                LogError($"Error checking startup: {ex}");
                return false;
            }
        }

        private static void ToggleStartup()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(StartupKeyPath, true))
                {
                    if (IsStartupEnabled())
                        key?.DeleteValue(AppName);
                    else
                        key?.SetValue(AppName, $"\"{Assembly.GetExecutingAssembly().Location}\"", RegistryValueKind.String);
                }
                UpdateTrayMenu();
            }
            catch (Exception ex)
            {
                LogError($"Error toggling startup: {ex}");
            }
        }

        private static void LogError(string message)
        {
            try
            {
                // Логируем ошибку в файл
                File.AppendAllText(LogFilePath, $"{DateTime.Now} - ERROR - {message}{Environment.NewLine}");

                // Показываем уведомление
                _notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
                _notifyIcon.BalloonTipTitle = AppName;
                _notifyIcon.BalloonTipText = _isRussian
                    ? $"Произошла ошибка, подробности смотрите в {Path.GetFileName(LogFilePath)} на рабочем столе."
                    : $"An error occurred, see details in {Path.GetFileName(LogFilePath)} on your desktop.";
                _notifyIcon.ShowBalloonTip(10000); // Показываем уведомление на 10 секунд
            }
            catch
            {
                // Игнорируем ошибки логирования и уведомления
            }
        }

        private class TrayApplicationContext : ApplicationContext
        {
            public TrayApplicationContext()
            {
                // Ничего не делаем, так как иконка уже инициализирована
            }

            protected override void Dispose(bool disposing)
            {
                _notifyIcon?.Dispose();
                base.Dispose(disposing);
            }
        }
    }
}