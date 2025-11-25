using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

#region THÔNG BÁO PHIÊN BẢN MỚI

public static class UpdateManager
{
    #region CẤU HÌNH

    // ⚙️ CẤU HÌNH CƠ BẢN (Có thể tùy chỉnh)
    private static int CHECK_INTERVAL_HOURS = 12;           // Kiểm tra mỗi 12 giờ
    private static readonly int HTTP_TIMEOUT_SECONDS = 8;            // Timeout khi tải file
    private static bool ENABLE_EPS_UNLOCK = false;          // Bật/tắt tính năng unlock EPS
    private static string[] ALLOWED_IP_PREFIXES = new[]     // Dải IP được phép unlock
    {
        "107.126.",
        "107.115."
    };
    private static string[] BLOCKED_IP_PREFIXES = new[]     // Dải IP bị chặn unlock
    {
        "107.125."
    };

    // ⚙️ CẤU HÌNH ĐƯỜNG DẪN UNLOCK
    private static readonly string UNLOCK_BAT_BASE_URL = "http://107.126.41.111:8888/unlock/"; //Đường dẫn chứa file .bat
    #endregion

    #region BIẾN NỘI BỘ

    private static System.Windows.Forms.Timer _updateCheckTimer;
    private static DateTime _lastCheckTime = DateTime.MinValue;

    #endregion

    #region PUBLIC API

    /// <summary>
    /// Khởi tạo tự động kiểm tra cập nhật
    /// </summary>
    /// <param name="exeName">Tên file .exe (vd: "MyApp.exe")</param>
    /// <param name="httpServers">Danh sách HTTP servers</param>
    /// <param name="checkIntervalHours">Kiểm tra mỗi bao nhiêu giờ (mặc định 12)</param>
    /// <param name="enableEpsUnlock">Có bật unlock EPS không (mặc định false)</param>
    public static void InitializeAutoCheck(
        string exeName,
        string[] httpServers,
        int checkIntervalHours = 12,
        bool enableEpsUnlock = false,
        string unlockBatBaseUrl = null)
    {
        if (unlockBatBaseUrl is null)
        {
            throw new ArgumentNullException(nameof(unlockBatBaseUrl));
        }

        CHECK_INTERVAL_HOURS = checkIntervalHours;
        ENABLE_EPS_UNLOCK = enableEpsUnlock;

        StopAutoCheck();
        CheckForUpdates(exeName, httpServers);
        _lastCheckTime = DateTime.Now;

        _updateCheckTimer = new System.Windows.Forms.Timer
        {
            Interval = CHECK_INTERVAL_HOURS * 60 * 60 * 1000
        };

        _updateCheckTimer.Tick += (s, e) =>
        {
            try
            {
                TimeSpan timeSinceLastCheck = DateTime.Now - _lastCheckTime;
                if (timeSinceLastCheck.TotalHours >= CHECK_INTERVAL_HOURS)
                {
                    Debug.WriteLine($"[Auto Check] Đã {timeSinceLastCheck.TotalHours:F1} giờ, kiểm tra cập nhật...");
                    CheckForUpdates(exeName, httpServers);
                    _lastCheckTime = DateTime.Now;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Auto Check] Lỗi trong timer: {ex.Message}");
                InitializeAutoCheck(exeName, httpServers, checkIntervalHours, enableEpsUnlock);
            }
        };

        _updateCheckTimer.Start();
        Debug.WriteLine($"[Auto Check] Timer đã khởi động - kiểm tra mỗi {CHECK_INTERVAL_HOURS} giờ");
    }

    /// <summary>
    /// Cấu hình dải IP cho unlock EPS
    /// </summary>
    public static void ConfigureIPRanges(string[] allowedPrefixes, string[] blockedPrefixes = null)
    {
        ALLOWED_IP_PREFIXES = allowedPrefixes ?? new string[0];
        BLOCKED_IP_PREFIXES = blockedPrefixes ?? new string[0];
    }

    /// <summary>
    /// Khởi động lại timer nếu bị dừng
    /// </summary>
    public static void RestartTimerIfStopped(string exeName, string[] httpServers)
    {
        if (_updateCheckTimer == null || !_updateCheckTimer.Enabled)
        {
            Debug.WriteLine("[Auto Check] Timer bị dừng, khởi động lại...");
            InitializeAutoCheck(exeName, httpServers, CHECK_INTERVAL_HOURS, ENABLE_EPS_UNLOCK);
        }
    }

    /// <summary>
    /// Dừng kiểm tra tự động
    /// </summary>
    public static void StopAutoCheck()
    {
        if (_updateCheckTimer != null)
        {
            _updateCheckTimer.Stop();
            _updateCheckTimer.Dispose();
            _updateCheckTimer = null;
            Debug.WriteLine("[Auto Check] Timer đã dừng");
        }
    }

    #endregion

    #region KIỂM TRA CẬP NHẬT

    /// <summary>
    /// Kiểm tra phiên bản mới
    /// </summary>
    public static async void CheckForUpdates(string exeName, string[] httpServers)
    {
        try
        {
            string currentVersion = Application.ProductVersion;
            string latestVersion = null;
            string changelog = "";
            string workingServerUrl = null;

            Debug.WriteLine($"[Cập nhật] Phiên bản hiện tại: {currentVersion}");

            // Kiểm tra version qua HTTP
            var httpResult = await TryCheckVersionViaHTTP(httpServers);
            if (!httpResult.Success)
            {
                Debug.WriteLine("[Cập nhật] ❌ Không thể kết nối đến server cập nhật!");
                return;
            }

            latestVersion = httpResult.Version;
            workingServerUrl = httpResult.ServerUrl;
            Debug.WriteLine($"[Cập nhật] ✅ HTTP thành công! Phiên bản: {latestVersion}");

            // Lấy changelog
            changelog = await GetChangelogSafe(workingServerUrl);

            // So sánh version
            if (string.Compare(latestVersion, currentVersion, StringComparison.OrdinalIgnoreCase) > 0)
            {
                Debug.WriteLine($"[Cập nhật] Đã có phiên bản mới: {latestVersion} > {currentVersion}");
                ShowUpdatePrompt(latestVersion, changelog, workingServerUrl, exeName);
            }
            else
            {
                Debug.WriteLine($"[Cập nhật] Đã cập nhật: {currentVersion}");
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[Cập nhật bị lỗi] {ex.Message}");
        }
    }

    private static async Task<(bool Success, string Version, string ServerUrl)> TryCheckVersionViaHTTP(string[] servers)
    {
        using (var client = new HttpClient { Timeout = TimeSpan.FromSeconds(HTTP_TIMEOUT_SECONDS) })
        {
            foreach (var server in servers)
            {
                try
                {
                    string url = server.TrimEnd('/') + "/version.txt";
                    Debug.WriteLine($"[HTTP] Đang thử: {url}");
                    string version = (await client.GetStringAsync(url)).Trim();
                    return (true, version, server.TrimEnd('/'));
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[HTTP] Lỗi {server}: {ex.Message}");
                }
            }
        }
        return (false, null, null);
    }

    private static async Task<string> GetChangelogSafe(string serverUrl)
    {
        try
        {
            using (var client = new HttpClient { Timeout = TimeSpan.FromSeconds(HTTP_TIMEOUT_SECONDS) })
            {
                return await client.GetStringAsync(serverUrl + "/changelog.txt");
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[Changelog] Lỗi: {ex.Message}");
        }
        return "(Không có thông tin thay đổi)";
    }

    #endregion

    #region EPS UNLOCK (OPTIONAL)
    /// <summary>
    /// Tìm và tải tất cả file .bat unlock qua HTTP
    /// </summary>
    private static async Task<List<string>> FindAllUnlockBatAsync(IProgress<string> progress, string serverUrl = null)
    {
        var batFiles = new List<string>();

        // ========================================
        // PHẦN 1: QUÉT QUA HTTP (ƯU TIÊN)
        // ========================================
        string unlockBaseUrl = !string.IsNullOrEmpty(UNLOCK_BAT_BASE_URL)
            ? UNLOCK_BAT_BASE_URL
            : (!string.IsNullOrEmpty(serverUrl) ? serverUrl.TrimEnd('/') + "/unlock/" : null);

        if (!string.IsNullOrEmpty(unlockBaseUrl))
        {
            try
            {
                progress?.Report($"🌐 Đang quét HTTP: {unlockBaseUrl}");

                using (var client = new HttpClient { Timeout = TimeSpan.FromSeconds(HTTP_TIMEOUT_SECONDS) })
                {
                    string listUrl = unlockBaseUrl.TrimEnd('/') + "/list.txt";

                    try
                    {
                        string fileList = await client.GetStringAsync(listUrl);
                        var httpBatFiles = fileList
                            .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                            .Where(f => f.EndsWith(".bat", StringComparison.OrdinalIgnoreCase))
                            .Select(f => f.Trim())
                            .ToList();

                        if (httpBatFiles.Any())
                        {
                            progress?.Report($"✅ Tìm thấy {httpBatFiles.Count} file .bat trên HTTP");

                            foreach (var fileName in httpBatFiles)
                            {
                                // ✅ SỬA LẠI: Dùng unlockBaseUrl thay vì serverUrl
                                string downloadUrl = unlockBaseUrl.TrimEnd('/') + "/" + fileName;
                                string tempBatPath = Path.Combine(Path.GetTempPath(), fileName);

                                try
                                {
                                    progress?.Report($"⬇️ Đang tải: {fileName}");
                                    byte[] batContent = await client.GetByteArrayAsync(downloadUrl);
                                    File.WriteAllBytes(tempBatPath, batContent);
                                    batFiles.Add(tempBatPath);
                                    progress?.Report($"✅ Đã tải: {fileName}");
                                }
                                catch (Exception ex)
                                {
                                    progress?.Report($"⚠️ Lỗi tải {fileName}: {ex.Message}");
                                }
                            }
                        }
                    }
                    catch (HttpRequestException)
                    {
                        progress?.Report("ℹ️ Không tìm thấy list.txt, thử tải file mặc định...");

                        var defaultBatFiles = new[]
                        {
                        "unlock_eps.bat",
                        "disable_eps.bat",
                        "unblock_printer.bat",
                        "unlock.bat"
                    };

                        foreach (var fileName in defaultBatFiles)
                        {
                            // ✅ SỬA LẠI: Dùng unlockBaseUrl thay vì serverUrl
                            string downloadUrl = unlockBaseUrl.TrimEnd('/') + "/" + fileName;
                            string tempBatPath = Path.Combine(Path.GetTempPath(), fileName);

                            try
                            {
                                byte[] batContent = await client.GetByteArrayAsync(downloadUrl);
                                File.WriteAllBytes(tempBatPath, batContent);
                                batFiles.Add(tempBatPath);
                                progress?.Report($"✅ Đã tải: {fileName}");
                            }
                            catch
                            {
                                // Bỏ qua nếu file không tồn tại
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                progress?.Report($"⚠️ Lỗi quét HTTP: {ex.Message}");
            }
        }

        // ========================================
        // PHẦN 2: QUÉT LOCAL (DỰ PHÒNG)
        // ========================================
        if (!batFiles.Any())
        {
            await Task.Run(() =>
            {
                var localPaths = new List<string>
            {
                Application.StartupPath,
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), "EPS"),
            };

                foreach (var localPath in localPaths)
                {
                    if (!Directory.Exists(localPath)) continue;

                    try
                    {
                        progress?.Report($"🔍 Đang quét local: {localPath}");
                        var foundFiles = Directory.GetFiles(localPath, "*.bat", SearchOption.AllDirectories)
                            .Where(f =>
                            {
                                string name = Path.GetFileName(f).ToLower();
                                return name.Contains("unlock") || name.Contains("eps") ||
                                       name.Contains("disable") || name.Contains("unblock");
                            })
                            .ToList();

                        if (foundFiles.Any())
                        {
                            progress?.Report($"✅ Tìm thấy {foundFiles.Count} file .bat local");
                            batFiles.AddRange(foundFiles);
                        }
                    }
                    catch (Exception ex)
                    {
                        progress?.Report($"⚠️ Lỗi quét {localPath}: {ex.Message}");
                    }
                }
            });
        }

        if (!batFiles.Any())
        {
            progress?.Report("ℹ️ Không tìm thấy file .bat unlock");
            return batFiles;
        }

        batFiles = batFiles.OrderByDescending(f =>
        {
            string name = Path.GetFileName(f).ToLower();
            if (name.Contains("unlock") && name.Contains("eps")) return 2;
            if (name.Contains("unlock") || name.Contains("eps")) return 1;
            return 0;
        }).ToList();

        return batFiles;
    }

    private static bool IsAllowedIPRange()
    {
        try
        {
            var host = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName());
            foreach (var ip in host.AddressList)
            {
                if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                {
                    string ipString = ip.ToString();

                    // Kiểm tra blocked trước
                    foreach (var blocked in BLOCKED_IP_PREFIXES)
                    {
                        if (ipString.StartsWith(blocked))
                        {
                            Debug.WriteLine($"[IP Check] ❌ IP bị chặn: {ipString}");
                            return false;
                        }
                    }

                    // Kiểm tra allowed
                    foreach (var allowed in ALLOWED_IP_PREFIXES)
                    {
                        if (ipString.StartsWith(allowed))
                        {
                            Debug.WriteLine($"[IP Check] ✅ IP được phép: {ipString}");
                            return true;
                        }
                    }
                }
            }

            Debug.WriteLine("[IP Check] ⚠️ IP không trong danh sách");
            return false;
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[IP Check] ⚠️ Lỗi: {ex.Message}");
            return false;
        }
    }

    private static async Task<bool> RunBatFileAsync(string batFilePath, IProgress<string> progress)
    {
        try
        {
            progress?.Report($"▶️ Đang chạy: {Path.GetFileName(batFilePath)}");

            var processInfo = new ProcessStartInfo
            {
                FileName = batFilePath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WorkingDirectory = Path.GetDirectoryName(batFilePath)
            };

            using (var process = new Process { StartInfo = processInfo })
            {
                process.OutputDataReceived += (s, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                        progress?.Report($"  📄 {e.Data}");
                };

                process.ErrorDataReceived += (s, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                        progress?.Report($"  ⚠️ {e.Data}");
                };

                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();

                bool completed = await Task.Run(() => process.WaitForExit(30000));

                if (!completed)
                {
                    progress?.Report("⏱️ Timeout - Dừng process");
                    try { process.Kill(); } catch { }
                    return false;
                }

                bool success = process.ExitCode == 0;
                progress?.Report(success
                    ? "✅ Unlock thành công!"
                    : $"⚠️ Exit code: {process.ExitCode}");

                return success;
            }
        }
        catch (Exception ex)
        {
            progress?.Report($"❌ Lỗi chạy .bat: {ex.Message}");
            return false;
        }
    }

    #endregion

    #region GIAO DIỆN THÔNG BÁO

    private static void ShowUpdatePrompt(string latestVersion, string changelog, string serverUrl, string exeName)
    {
        int cornerRadius = 20;
        var updateForm = new Form
        {
            Text = "Cập nhật phần mềm",
            Size = new Size(450, 300),
            StartPosition = FormStartPosition.Manual,
            FormBorderStyle = FormBorderStyle.None,
            TopMost = true,
            BackColor = Color.White,
            Icon = Application.OpenForms.Count > 0 ? Application.OpenForms[0].Icon : SystemIcons.Application
        };

        updateForm.Location = new Point(
            Screen.PrimaryScreen.WorkingArea.Right - updateForm.Width - 20,
            Screen.PrimaryScreen.WorkingArea.Bottom - updateForm.Height - 20
        );

        IntPtr hRgn = CreateRoundRectRgn(0, 0, updateForm.Width, updateForm.Height, cornerRadius, cornerRadius);
        updateForm.Region = Region.FromHrgn(hRgn);

        int val = 2;
        DwmSetWindowAttribute(updateForm.Handle, 2, ref val, 4);
        MARGINS margins = new MARGINS() { cxLeftWidth = 1, cxRightWidth = 1, cyTopHeight = 1, cyBottomHeight = 1 };
        DwmExtendFrameIntoClientArea(updateForm.Handle, ref margins);

        updateForm.Paint += (s, e) =>
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var pen = new Pen(Color.FromArgb(128, Color.LightGray)))
            {
                pen.Width = 1f;
                e.Graphics.DrawRectangle(pen, 0.5f, 0.5f, updateForm.Width - 1, updateForm.Height - 1);
            }
        };

        // Icon
        var assembly = Assembly.GetExecutingAssembly();
        string resourceName = "ITCONFIRM.src.update_icon.png";
        Image iconImage = SystemIcons.Shield.ToBitmap();

        try
        {
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream != null && stream.Length > 0)
                {
                    using (var tempImage = new Bitmap(stream))
                    {
                        iconImage = new Bitmap(tempImage);
                    }
                }
            }
        }
        catch { }

        var picIcon = new PictureBox
        {
            Size = new Size(40, 40),
            Location = new Point(20, 20),
            Image = iconImage,
            SizeMode = PictureBoxSizeMode.Zoom,
            BackColor = Color.Transparent
        };

        string appName = exeName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)
            ? exeName.Substring(0, exeName.Length - 4)
            : exeName;

        string currentVersion = Application.ProductVersion;
        var lblVersion = new Label
        {
            Text = $"{appName} đã có phiên bản mới: {latestVersion}\nPhiên bản hiện tại: {currentVersion}",
            Location = new Point(70, 15),
            Width = updateForm.Width - 90,
            Height = 40,
            TextAlign = ContentAlignment.TopLeft,
            Font = new Font("Segoe UI", 10, FontStyle.Bold)
        };

        var rtbChangelog = new RichTextBox
        {
            Text = changelog,
            Location = new Point(50, 60),
            Width = updateForm.Width - 65,
            Height = 170,
            BorderStyle = BorderStyle.None,
            BackColor = Color.White,
            Font = new Font("Segoe UI", 9),
            ScrollBars = RichTextBoxScrollBars.Vertical,
            ReadOnly = true,
            WordWrap = true,
        };

        var txtLog = new TextBox
        {
            Location = new Point(20, 80),
            Width = updateForm.Width - 40,
            Height = 150,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            ReadOnly = true,
            Font = new Font("Consolas", 8),
            BackColor = Color.FromArgb(240, 240, 240),
            Visible = false
        };

        var panelButtons = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 70,
            BackColor = Color.White,
            Padding = new Padding(0, 0, 0, 10)
        };

        var btnUpdate = new Button
        {
            Text = "Cập nhật",
            Width = 100,
            Height = 35,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI Semibold", 10F),
            BackColor = Color.FromArgb(0, 120, 212),
            ForeColor = Color.White,
            Cursor = Cursors.Hand
        };
        btnUpdate.FlatAppearance.BorderSize = 0;
        btnUpdate.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 150, 255);
        btnUpdate.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, btnUpdate.Width, btnUpdate.Height, 10, 10));

        var btnSkip = new Button
        {
            Text = "Bỏ qua",
            Width = 100,
            Height = 35,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI Semibold", 10F),
            BackColor = Color.FromArgb(200, 200, 200),
            ForeColor = Color.Black,
            Cursor = Cursors.Hand
        };
        var lblWarning = new Label
        {
            Text = "Nếu báo lỗi không tự động cập nhật\r\nHãy Unlock EPS trước khi bấm cập nhật ứng dụng",
            ForeColor = Color.Red,
            Font = new Font("Segoe UI", 8, FontStyle.Italic),
            Width = panelButtons.Width,
            Height = 30,
            TextAlign = ContentAlignment.MiddleCenter,
            Location = new Point(panelButtons.Width - 70, 40) // đặt dưới 2 nút (10 + 35 = 45)
        };
        btnSkip.FlatAppearance.BorderSize = 0;
        btnSkip.FlatAppearance.MouseOverBackColor = Color.FromArgb(170, 170, 170);
        btnSkip.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, btnSkip.Width, btnSkip.Height, 10, 10));

        btnUpdate.Location = new Point(70, 5);
        btnSkip.Location = new Point(panelButtons.Width - btnSkip.Width - 70, 5);
        btnSkip.Anchor = AnchorStyles.Top | AnchorStyles.Right;

        btnSkip.Click += (s, e) => updateForm.Close();

        btnUpdate.Click += async (s, e) =>
        {
            btnUpdate.Enabled = false;
            btnSkip.Enabled = false;
            txtLog.Visible = true;

            var progress = new Progress<string>(msg =>
            {
                if (updateForm.InvokeRequired)
                {
                    updateForm.Invoke(new Action(() =>
                    {
                        txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}");
                        txtLog.SelectionStart = txtLog.Text.Length;
                        txtLog.ScrollToCaret();
                    }));
                }
                else
                {
                    txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}");
                    txtLog.SelectionStart = txtLog.Text.Length;
                    txtLog.ScrollToCaret();
                }
            });

            await DownloadAndUpdateAsync(serverUrl, exeName, btnUpdate, updateForm, progress);
        };

        panelButtons.Controls.Add(btnUpdate);
        panelButtons.Controls.Add(btnSkip);
        updateForm.Controls.Add(picIcon);
        updateForm.Controls.Add(lblVersion);
        updateForm.Controls.Add(txtLog);
        updateForm.Controls.Add(rtbChangelog);
        panelButtons.Controls.Add(lblWarning);
        updateForm.Controls.Add(panelButtons);
        updateForm.Show();
    }

    #endregion

    #region TẢI VÀ CẬP NHẬT

    private static async Task DownloadAndUpdateAsync(
        string serverUrl, string exeName,
        Button btnUpdate, Form updateForm, IProgress<string> progress)
    {
        string baseName = Path.GetFileNameWithoutExtension(exeName);
        string tempFile = Path.Combine(Path.GetTempPath(), baseName + "_Update.exe");

        try
        {
            // BƯỚC 1: UNLOCK EPS (NẾU ĐƯỢC BẬT)
            if (ENABLE_EPS_UNLOCK && IsAllowedIPRange())
            {
                progress?.Report("🔓 Bắt đầu unlock...");
                var unlockBatFiles = await FindAllUnlockBatAsync(progress, serverUrl);

                if (unlockBatFiles != null && unlockBatFiles.Any())
                {
                    progress?.Report($"📋 Tìm thấy {unlockBatFiles.Count} file unlock");

                    foreach (var batPath in unlockBatFiles)
                    {
                        bool success = await RunBatFileAsync(batPath, progress);
                        if (success)
                        {
                            progress?.Report($"✅ Unlock thành công!");
                            await Task.Delay(1000);
                            break;
                        }
                    }
                }
            }

            // BƯỚC 2: TẢI FILE CẬP NHẬT
            progress?.Report("📥 Đang tải bản cập nhật...");
            bool downloadSuccess = await DownloadUpdateViaHTTP(serverUrl, exeName, tempFile, btnUpdate, progress);

            if (!downloadSuccess)
                throw new Exception("Không thể tải file cập nhật");

            await Task.Delay(1000);

            // BƯỚC 3: TẠO BATCH SCRIPT VÀ KHỞI ĐỘNG LẠI
            progress?.Report("🔄 Đang chuẩn bị cập nhật...");

            string currentExe = Application.ExecutablePath;
            string currentVersion = Application.ProductVersion.Replace(".", "_"); // Chuyển 1.0.0 thành 1_0_0
            string exeDirectory = Path.GetDirectoryName(currentExe);
            string exeBaseName = Path.GetFileNameWithoutExtension(currentExe);
            string backupExePath = Path.Combine(exeDirectory, exeBaseName + "_v" + currentVersion + ".exe");

            // XÓA CÁC BACKUP CŨ - CHỈ GIỮ 2 PHIÊN BẢN GẦN NHẤT
            try
            {
                var backupFiles = Directory.GetFiles(exeDirectory, exeBaseName + "_v*.exe")
                    .Select(f => new FileInfo(f))
                    .OrderByDescending(f => f.LastWriteTime)
                    .ToList();

                if (backupFiles.Count >= 2)
                {
                    // Xóa các file backup cũ, chỉ giữ lại file mới nhất
                    var filesToDelete = backupFiles.Skip(2).ToList();
                    foreach (var oldBackup in filesToDelete)
                    {
                        try
                        {
                            File.Delete(oldBackup.FullName);
                            progress?.Report($"🗑️ Đã xóa backup cũ: {oldBackup.Name}");
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[Cleanup] Không thể xóa {oldBackup.Name}: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Cleanup] Lỗi khi dọn dẹp backup: {ex.Message}");
            }

            string batFile = Path.Combine(Path.GetTempPath(), "update.bat");

            string batContent = $@"@echo off
chcp 65001 >nul
echo Đang chờ ứng dụng đóng...
:waitloop
timeout /t 3 /nobreak >nul
tasklist /FI ""IMAGENAME eq {Path.GetFileName(currentExe)}"" 2>NUL | find /I /N ""{Path.GetFileName(currentExe)}"">NUL
if ""%ERRORLEVEL%""==""0"" goto waitloop

echo Bắt đầu cập nhật...
timeout /t 3 /nobreak >nul

if exist ""{currentExe}"" ren ""{currentExe}"" ""{Path.GetFileName(backupExePath)}""

copy /y ""{tempFile}"" ""{currentExe}""
timeout /t 3 /nobreak >nul

if exist ""{currentExe}"" (
    start """" ""{currentExe}""
    del /f /q ""{tempFile}"" 2>nul
    del /f /q ""%~f0""
)";

            File.WriteAllText(batFile, batContent, new System.Text.UTF8Encoding(false));

            btnUpdate.Text = "Hoàn tất ✔";
            progress?.Report("🎉 Cập nhật hoàn tất! Đang khởi động lại...");

            await Task.Delay(2000);

            Process.Start(new ProcessStartInfo
            {
                FileName = batFile,
                UseShellExecute = true,
                CreateNoWindow = false
            });

            await Task.Delay(1000);
            updateForm.Close();
            Application.Exit();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[Lỗi cập nhật] {ex.Message}");
            progress?.Report($"❌ LỖI: {ex.Message}");
            MessageBox.Show($"Cập nhật thất bại:\n\n{ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            btnUpdate.Text = "Cập nhật";
            btnUpdate.Enabled = true;
        }
    }

    private static async Task<bool> DownloadUpdateViaHTTP(string serverUrl, string exeName, string tempFile, Button btnUpdate, IProgress<string> progress)
    {
        try
        {
            using (var client = new HttpClient { Timeout = TimeSpan.FromMinutes(10) })
            using (var response = await client.GetAsync(serverUrl + "/" + exeName, HttpCompletionOption.ResponseHeadersRead))
            {
                response.EnsureSuccessStatusCode();
                var total = response.Content.Headers.ContentLength ?? -1L;

                using (var input = await response.Content.ReadAsStreamAsync())
                using (var output = new FileStream(tempFile, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true))
                {
                    var buffer = new byte[8192];
                    long totalRead = 0;
                    int read;
                    int lastPercent = 0;

                    do
                    {
                        read = await input.ReadAsync(buffer, 0, buffer.Length);
                        if (read > 0)
                        {
                            await output.WriteAsync(buffer, 0, read);
                            totalRead += read;

                            if (total != -1)
                            {
                                int percent = (int)(totalRead * 100 / total);

                                if (btnUpdate.InvokeRequired)
                                    btnUpdate.Invoke(new Action(() => btnUpdate.Text = $"Đang tải... {percent}%"));
                                else
                                    btnUpdate.Text = $"Đang tải... {percent}%";

                                if (percent != lastPercent && percent % 10 == 0)
                                {
                                    progress?.Report($"📊 Tiến độ: {percent}%");
                                    lastPercent = percent;
                                }
                            }
                        }
                    } while (read > 0);
                }
            }

            progress?.Report("✅ Tải xuống hoàn tất!");
            return true;
        }
        catch (Exception ex)
        {
            progress?.Report($"❌ Lỗi tải: {ex.Message}");
            return false;
        }
    }

    #endregion

    #region WIN32 API

    [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
    private static extern IntPtr CreateRoundRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect,
        int nWidthEllipse, int nHeightEllipse);

    [DllImport("dwmapi.dll")]
    private static extern int DwmExtendFrameIntoClientArea(IntPtr hWnd, ref MARGINS pMarInset);

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

    private struct MARGINS
    {
        public int cxLeftWidth;
        public int cxRightWidth;
        public int cyTopHeight;
        public int cyBottomHeight;
    }

    #endregion
}
#endregion