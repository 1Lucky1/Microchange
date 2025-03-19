using AudioSwitcher.AudioApi;
using AudioSwitcher.AudioApi.CoreAudio;
using AudioSwitcher.AudioApi.Observables;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
        private static readonly string Version = "2.2.0c";
        private static readonly string LogFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "FullMicrochangeCS_error_log.txt");
        private static readonly string StartupKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private static bool _isRussian;
        private static IDisposable _deviceChangeSubscription;
        private static ToolStripMenuItem _playbackMenuItem;
        private static ToolStripMenuItem _recordingMenuItem;
        private static ContextMenuStrip _contextMenu;

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            _isRussian = System.Globalization.CultureInfo.CurrentCulture.Name.StartsWith("ru", StringComparison.OrdinalIgnoreCase);

            try
            {
                _audioController = new CoreAudioController();

                _notifyIcon = new NotifyIcon
                {
                    Icon = FullMicrochangeCS.Properties.Resources.FullMicrochange,
                    Visible = true,
                    Text = AppName
                };
                _notifyIcon.Click += (s, e) => ShowTrayMenu();

                InitializeTrayMenu();
                SubscribeToDeviceChanges();

                Application.Run(new TrayApplicationContext());
            }
            catch (Exception ex)
            {
                LogError($"Initialization error: {ex}");
                _notifyIcon?.Dispose();
                Environment.Exit(1);
            }
        }

        private static void InitializeTrayMenu()
        {
            _contextMenu = new ContextMenuStrip();

            _playbackMenuItem = new ToolStripMenuItem(_isRussian ? "Устройства воспроизведения" : "Playback Devices");
            _contextMenu.Items.Add(_playbackMenuItem);

            _recordingMenuItem = new ToolStripMenuItem(_isRussian ? "Устройства записи" : "Recording Devices");
            _contextMenu.Items.Add(_recordingMenuItem);

            _contextMenu.Items.Add(new ToolStripSeparator());

            var autostartItem = new ToolStripMenuItem(_isRussian ? "Автозапуск с системой" : "Startup with system")
            {
                Checked = IsStartupEnabled(),
                CheckOnClick = true
            };
            autostartItem.Click += (s, e) => ToggleStartup();
            _contextMenu.Items.Add(autostartItem);

            _contextMenu.Items.Add(new ToolStripSeparator());

            _contextMenu.Items.Add(new ToolStripMenuItem(_isRussian ? "Закрыть программу" : "Close application", null, (s, e) =>
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
                Application.Exit();
            }));

            var aboutItem = new ToolStripMenuItem(_isRussian ? "О программе" : "About");
            aboutItem.Click += (s, e) => ShowAboutDialog();
            _contextMenu.Items.Add(aboutItem);

            _notifyIcon.ContextMenuStrip = _contextMenu;

            UpdateDeviceLists();
        }

        private static void ShowTrayMenu()
        {
            UpdateDeviceLists();
            _notifyIcon.ContextMenuStrip.Show(Cursor.Position);
        }

        private static void SubscribeToDeviceChanges()
        {
            try
            {
                _deviceChangeSubscription = _audioController.AudioDeviceChanged.Subscribe(OnDeviceChanged);
            }
            catch (Exception ex)
            {
                LogError($"Error subscribing to device changes: {ex}");
            }
        }

        private static void OnDeviceChanged(DeviceChangedArgs args)
        {
            try
            {
                // Сохраняем Id устройства, чтобы найти его позже
                Guid? deviceId = args.Device?.Id;
                bool isActive = args.Device != null && args.Device.State == DeviceState.Active;

                // Убиваем старый экземпляр и создаём новый
                _deviceChangeSubscription?.Dispose();
                _audioController?.Dispose();
                _audioController = new CoreAudioController();
                SubscribeToDeviceChanges(); // Подписываемся заново

                UpdateDeviceLists();

                // Если устройство активно, ищем его в новом списке по Id и показываем уведомление
                if (isActive && deviceId.HasValue)
                {
                    // Ищем устройство в новом списке
                    var allDevices = _audioController.GetPlaybackDevices(DeviceState.Active)
                        .Cast<IDevice>()
                        .Concat(_audioController.GetCaptureDevices(DeviceState.Active).Cast <IDevice>())
                        .ToList();

                    var newDevice = allDevices.FirstOrDefault(d => d.Id == deviceId.Value);
                    if (newDevice != null)
                    {
                        string deviceName = GetDeviceName(newDevice);
                        _notifyIcon.BalloonTipIcon = ToolTipIcon.Info;
                        _notifyIcon.BalloonTipTitle = AppName;
                        _notifyIcon.BalloonTipText = _isRussian
                            ? $"Обнаружено новое устройство: {deviceName}!"
                            : $"New device detected: {deviceName}!";
                        _notifyIcon.ShowBalloonTip(5000);
                    }
                    else
                    {
                        LogError($"Device with ID {deviceId} not found in new list.");
                    }
                }
            }
            catch (Exception ex)
            {
                LogError($"Error handling device change: {ex}");
            }
        }

        private static void UpdateDeviceLists()
        {
            try
            {
                var playbackDevices = _audioController.GetPlaybackDevices(DeviceState.Active).Cast<IDevice>().ToList();
                var recordingDevices = _audioController.GetCaptureDevices(DeviceState.Active).Cast<IDevice>().ToList();

                UpdatePlaybackMenuItems(playbackDevices);
                UpdateRecordingMenuItems(recordingDevices);

                if (!playbackDevices.Any() && !recordingDevices.Any())
                {
                    _playbackMenuItem.DropDownItems.Clear();
                    _playbackMenuItem.DropDownItems.Add(new ToolStripMenuItem(_isRussian ? "Нет активных устройств" : "No active devices"));
                    _recordingMenuItem.DropDownItems.Clear();
                }
            }
            catch (Exception ex)
            {
                LogError($"Error updating device lists: {ex}");
            }
        }

        private static void UpdatePlaybackMenuItems(List<IDevice> devices)
        {
            _playbackMenuItem.DropDownItems.Clear();
            var playbackItems = devices.Select(device =>
            {
                string deviceName = GetDeviceName(device);
                var item = new ToolStripMenuItem
                {
                    Text = deviceName,
                    Checked = device.IsDefaultDevice,
                    Tag = device,
                    CheckOnClick = true
                };
                item.Click += (s, e) =>
                {
                    var selectedDevice = (IDevice)item.Tag;
                    LogError($"Setting default device: {selectedDevice.FullName}");
                    selectedDevice.SetAsDefault();
                    UpdateDeviceLists();
                };
                return item;
            }).ToArray();
            _playbackMenuItem.DropDownItems.AddRange(playbackItems);
        }

        private static void UpdateRecordingMenuItems(List<IDevice> devices)
        {
            _recordingMenuItem.DropDownItems.Clear();
            var recordingItems = devices.Select(device =>
            {
                string deviceName = GetDeviceName(device);
                var item = new ToolStripMenuItem
                {
                    Text = deviceName,
                    Checked = device.IsDefaultDevice,
                    Tag = device,
                    CheckOnClick = true
                };
                item.Click += (s, e) =>
                {
                    var selectedDevice = (IDevice)item.Tag;
                    LogError($"Setting default device: {selectedDevice.FullName}");
                    selectedDevice.SetAsDefault();
                    UpdateDeviceLists();
                };
                return item;
            }).ToArray();
            _recordingMenuItem.DropDownItems.AddRange(recordingItems);
        }

        private static string GetDeviceName(IDevice device)
        {
            try
            {
                string name = device.FullName;
                if (string.IsNullOrEmpty(name) || name == "Unknown")
                {
                    name = device.Name;
                    if (string.IsNullOrEmpty(name) || name == "Unknown")
                    {
                        name = device.InterfaceName;
                    }
                }
                return string.IsNullOrEmpty(name) ? "Unknown Device" : name;
            }
            catch (Exception ex)
            {
                LogError($"Error getting device name: {ex}");
                return "Unknown Device";
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
                foreach (ToolStripItem item in _contextMenu.Items)
                {
                    if (item.Text == (_isRussian ? "Автозапуск с системой" : "Startup with system"))
                    {
                        ((ToolStripMenuItem)item).Checked = IsStartupEnabled();
                        break;
                    }
                }
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
                _notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
                _notifyIcon.BalloonTipTitle = AppName;
                _notifyIcon.BalloonTipText = _isRussian
                    ? $"Произошла ошибка, подробности смотрите в {Path.GetFileName(LogFilePath)} на рабочем столе."
                    : $"An error occurred, see details in {Path.GetFileName(LogFilePath)} on your desktop.";
                _notifyIcon.ShowBalloonTip(10000);
            }
            catch
            {
            }
        }

        private class TrayApplicationContext : ApplicationContext
        {
            public TrayApplicationContext()
            {
            }

            protected override void Dispose(bool disposing)
            {
                _deviceChangeSubscription?.Dispose();
                _audioController?.Dispose();
                _notifyIcon?.Dispose();
                base.Dispose(disposing);
            }
        }
    }
}