using AudioSwitcher.AudioApi;
using AudioSwitcher.AudioApi.CoreAudio;
using Microsoft.Win32;
using System;
using System.Drawing;
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
        private static readonly string AppName = "FullMicrochange";
        private static readonly string Version = "2.0c";
        private static readonly string LogFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Microchange_error_log.txt");
        private static readonly string StartupKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private static bool _isRussian;

        [STAThread]
        static void Main(string[] args)
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
                    Icon = FullMicrochange.Properties.Resources.FullMicrochange, //new Icon("FullMicrochange.ico"),
                    Visible = true,
                    Text = AppName
                };
                _notifyIcon.Click += (s, e) => UpdateTrayMenu();

                // Обновляем меню при старте
                UpdateTrayMenu();

                // Проверяем и настраиваем автозапуск при первом запуске программы (Пока что без надобности)
                // CheckAndSetupStartup();

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

                // Устройства воспроизведения (Playback) — только активные
                var playbackDevices = _audioController.GetPlaybackDevices()
                    .Where(d => d.State == DeviceState.Active) // Фильтруем только активные устройства
                    .ToList();
                //.Where(d => !d.Name.Contains("NVIDIA") && !d.Name.Contains("Rift")) // Исключаем "мусор" (можно настроить под свои нужды)


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
                        UpdateTrayMenu(); // Обновляем меню после изменения
                    };
                }
                contextMenu.Items.Add(new ToolStripMenuItem(_isRussian ? "Устройства воспроизведения" : "Playback Devices", null, playbackItems));

                // Разделитель
                contextMenu.Items.Add(new ToolStripSeparator());

                // Устройства записи (Recording) — только активные
                var recordingDevices = _audioController.GetCaptureDevices()
                    .Where(d => d.State == DeviceState.Active) // Фильтруем только активные устройства
                    //.Where(d => !d.Name.Contains("NVIDIA") && !d.Name.Contains("Rift")) // Исключаем "мусор" (можно настроить под свои нужды)
                    .ToList();

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
                        UpdateTrayMenu(); // Обновляем меню после изменения
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

                // Разделитель
                contextMenu.Items.Add(new ToolStripSeparator());

                // Версия
                contextMenu.Items.Add(new ToolStripMenuItem(_isRussian ? $"Версия {Version}" : $"Version {Version}", null, (s, e) => { }));

                _notifyIcon.ContextMenuStrip = contextMenu;
            }
            catch (Exception ex)
            {
                LogError($"Error updating tray menu: {ex}");
            }
        }

        private static void CheckAndSetupStartup()
        {
            try
            {
                if (IsStartupEnabled())
                {
                    return;
                }

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
                    if (key == null)
                    {
                        return false;
                    }

                    string currentValue = key.GetValue(AppName) as string;
                    if (string.IsNullOrEmpty(currentValue))
                    {
                        return false;
                    }

                    // Сравниваем текущий путь с путем к текущему запущенному файлу
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
                    {
                        key?.DeleteValue(AppName);
                    }
                    else
                    {
                        key?.SetValue(AppName, $"\"{Assembly.GetExecutingAssembly().Location}\"", RegistryValueKind.String);
                    }
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
                File.AppendAllText(LogFilePath, $"{DateTime.Now} - ERROR - {message}{Environment.NewLine}");
            }
            catch
            {
                // Игнорируем ошибки логирования
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