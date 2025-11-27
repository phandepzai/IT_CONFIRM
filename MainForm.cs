using ITCONFIRM;
using ITCONFIRM.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ITCONFIRM
{
    public partial class MainForm : Form
    {
        #region KHAI BÁO CÁC BIẾN
        private readonly System.Windows.Forms.Timer timer;
        private TextBox currentTextBox;
        private readonly ToolTip validationToolTip;
        private string _lastSavedFilePath; // Biến mới để lưu đường dẫn file        
        private readonly ToolTip statusToolTip;// Biến mới cho ToolTip của thông báo trạng thái     
        private readonly Timer rainbowTimer;// Các biến cho hiệu ứng chuyển màu cầu vồng mượt mà
        private bool isRainbowActive = false;
        private Color originalCopyrightColor;
        private double rainbowPhase = 0;
        private readonly Dictionary<string, Color> originalColors = new Dictionary<string, Color>();// Sử dụng Dictionary để lưu màu gốc của tất cả các nút
        private System.Windows.Forms.ComboBox cboModel;
        // Biến cho NAS       
        private string nasDirectoryPath; // readonly, chỉ gán trong constructor
        private Color defaultTextBoxBackColor; // Biến lưu trữ màu nền mặc định
        #endregion

        #region FORM KHỞI TẠO UI
        public MainForm()
        {
            InitializeComponent();
            InitializeFocusColor(); // Gọi hàm khởi tạo màu nền khi focus
            string eqpid = ReadEQPIDFromIniFile();
            this.Text = "IT CONFIRM" + (string.IsNullOrEmpty(eqpid) ? "" : "_" + eqpid + "");
            InitializeKeyboardEvents();
            TxtSAPN.MaxLength = 89;
            validationToolTip = new ToolTip
            {
                AutoPopDelay = 3000,
                InitialDelay = 500,
                ReshowDelay = 500,
                ShowAlways = true
            };
            statusToolTip = new ToolTip();//Gợi ý bấm vào để mở thư mục           
            this.LblStatus.Text = "Sẵn sàng nhập dữ liệu...";
            UpdateSavedSAPNCount();

            // Khởi tạo ComboBox với danh sách lỗi
            cboErrorType.Items.AddRange(new string[] { "B-SPOT", "WHITE SPOT", "ĐỐM SPIN", "ĐỐM PANEL", "ĐỐM ĐƯỜNG DỌC", "-" });
            cboErrorType.SelectedIndex = -1; // Mặc định chọn "ĐỐM" =-1 KHÔNG CHỌN GÌ

            // Khởi tạo timer cho hiệu ứng cầu vồng
            this.rainbowTimer = new Timer
            {
                Interval = 20 // Cập nhật màu mỗi 20ms để mượt hơn
            };
            this.rainbowTimer.Tick += new EventHandler(this.RainbowTimer_Tick);

            timer = new System.Windows.Forms.Timer { Interval = 500 };
            timer.Tick += Timer_Tick;
            timer.Start();

            // Gắn sự kiện cho LblCopyright
            LblCopyright.MouseEnter += LblCopyright_MouseEnter;
            LblCopyright.MouseLeave += LblCopyright_MouseLeave;

            // Gán sự kiện Click và thay đổi con trỏ chuột cho LblStatus
            this.LblStatus.Click += new EventHandler(this.LblStatus_Click);
            this.LblStatus.Cursor = Cursors.Hand;

            // Gán sự kiện Click và thay đổi con trỏ chuột cho LblStatus
            this.LblSAPNCount.Click += new EventHandler(this.LblSAPNCount_Click);
            this.LblSAPNCount.Cursor = Cursors.Hand;

            // Khởi tạo NAS
            var nasCredentials = ReadNASCredentialsFromIniFile(); // Gọi hàm để đọc credentials và tạo file NASConfig.ini
            AppendToLog($"Cấu hình NAS đã được đọc thành công. Path: {@"C:\IT_CONFIRM\Config\NASConfig.ini"}", System.Drawing.Color.Green);

            // Tạo đường dẫn file dựa trên ngày hiện tại
            SetFilePath();

            // Khởi tạo hiệu ứng cho các nút
            InitializeButtonEffects();
            //KIÊM TRA TỰ ĐỘNG CẬP NHẬT
            UpdateManager.InitializeAutoCheck(
                "IT CONFIRM.exe",
                new[]
                {
                    //"http://107.125.221.79:8888/update/ITCONF/",
                    //"http://107.126.41.111:8888/update/ITCONF/",
                    "http://192.168.111.101:8888/update/ITCONF/"
                },
                checkIntervalHours: 12,
                enableEpsUnlock: false,
                unlockBatBaseUrl: "http://107.126.41.111:8888/unlock/"
            );
            // --- Đọc config cho loại lỗi ---
            try
            {
                List<string> errorTypesList = ReadErrorTypesFromIniFile(); // Gọi phương thức mới

                // Cập nhật danh sách trong ComboBox loại lỗi dựa trên file config INI
                if (errorTypesList != null && errorTypesList.Count > 0)
                {
                    cboErrorType.Items.Clear(); // Xóa các mục cũ (nếu có)
                    cboErrorType.Items.AddRange(errorTypesList.ToArray());
                    // Không chọn mục nào mặc định, để người dùng chọn
                }
                else
                {
                    // Nếu không đọc được config hoặc config rỗng, có thể thêm các mục mặc định
                    // hoặc để ComboBox trống. Ở đây, để trống.
                    // Đảm bảo không null nếu cần thiết ở nơi khác, nhưng List<string> rỗng cũng ổn
                }
            }
            catch (Exception ex)
            {
                // LblStatus.ForeColor = System.Drawing.Color.Red; // Xóa
                // LblStatus.Text = $"Lỗi không mong đợi khi đọc cấu hình loại lỗi: {ex.Message}"; // Xóa
                AppendToLog($"Lỗi không mong đợi khi đọc cấu hình loại lỗi: {ex.Message}", System.Drawing.Color.Red);
            }
            // Gán sự kiện TextChanged cho các TextBox tọa độ
            TxtSx1.TextChanged += CoordinateTextBox_TextChanged;
            TxtSy1.TextChanged += CoordinateTextBox_TextChanged;
            TxtEx1.TextChanged += CoordinateTextBox_TextChanged;
            TxtEy1.TextChanged += CoordinateTextBox_TextChanged;
            TxtSx2.TextChanged += CoordinateTextBox_TextChanged;
            TxtSy2.TextChanged += CoordinateTextBox_TextChanged;
            TxtEx2.TextChanged += CoordinateTextBox_TextChanged;
            TxtEy2.TextChanged += CoordinateTextBox_TextChanged;
            TxtSx3.TextChanged += CoordinateTextBox_TextChanged;
            TxtSy3.TextChanged += CoordinateTextBox_TextChanged;
            TxtEx3.TextChanged += CoordinateTextBox_TextChanged;
            TxtEy3.TextChanged += CoordinateTextBox_TextChanged;
            TxtX1.TextChanged += CoordinateTextBox_TextChanged;
            TxtY1.TextChanged += CoordinateTextBox_TextChanged;
            TxtX2.TextChanged += CoordinateTextBox_TextChanged;
            TxtY2.TextChanged += CoordinateTextBox_TextChanged;
            TxtX3.TextChanged += CoordinateTextBox_TextChanged;
            TxtY3.TextChanged += CoordinateTextBox_TextChanged;
        }
        #endregion

        #region AUTO FOCUS KHI NHẬP ĐỦ 3 KÝ TỰ & MÀU NỀN KHI FOCUS
        private void CoordinateTextBox_TextChanged(object sender, EventArgs e)
        {
            TextBox currentTextBox = (TextBox)sender;

            // Kiểm tra xem TextBox hiện tại có phải là một trong các ô tọa độ không
            // và có phải là ô có giới hạn 3 ký tự (loại trừ ô sAPN có 89 ký tự)
            // Danh sách các TextBox tọa độ có giới hạn 3 ký tự
            var coordinateTextBoxes = new TextBox[] {
                TxtSx1, TxtSy1, TxtEx1, TxtEy1,
                TxtSx2, TxtSy2, TxtEx2, TxtEy2,
                TxtSx3, TxtSy3, TxtEx3, TxtEy3,
                TxtX1, TxtY1, TxtX2, TxtY2, TxtX3
            };

            if (coordinateTextBoxes.Contains(currentTextBox))
            {
                // Kiểm tra độ dài và đảm bảo đủ 3 ký tự
                if (currentTextBox.Text.Length == 3)
                {
                    // Chuyển sang ô tiếp theo trong danh sách
                    // SelectNextControl sẽ chọn ô tiếp theo theo thứ tự tab index
                    // Nếu bạn muốn thứ tự cụ thể, bạn cần tự quản lý danh sách và logic chuyển tiếp
                    this.SelectNextControl(currentTextBox, true, true, true, true);
                }
            }        
        }

        private void InitializeFocusColor()
        {
            // Lưu màu nền mặc định của một TextBox bất kỳ, hoặc sử dụng SystemColors.Window
            // Nếu bạn không thay đổi màu nền mặc định trong Designer, SystemColors.Window là an toàn.
            defaultTextBoxBackColor = SystemColors.Window; // Hoặc Color.FromKnownColor(KnownColor.Window);

            // Gán sự kiện Enter và Leave cho các TextBox tọa độ
            // Bạn có thể thêm TxtSAPN vào đây nếu muốn áp dụng cho cả ô sAPN
            TxtSx1.Enter += TextBox_Enter;
            TxtSx1.Leave += TextBox_Leave;
            TxtSy1.Enter += TextBox_Enter;
            TxtSy1.Leave += TextBox_Leave;
            TxtEx1.Enter += TextBox_Enter;
            TxtEx1.Leave += TextBox_Leave;
            TxtEy1.Enter += TextBox_Enter;
            TxtEy1.Leave += TextBox_Leave;
            TxtSx2.Enter += TextBox_Enter;
            TxtSx2.Leave += TextBox_Leave;
            TxtSy2.Enter += TextBox_Enter;
            TxtSy2.Leave += TextBox_Leave;
            TxtEx2.Enter += TextBox_Enter;
            TxtEx2.Leave += TextBox_Leave;
            TxtEy2.Enter += TextBox_Enter;
            TxtEy2.Leave += TextBox_Leave;
            TxtSx3.Enter += TextBox_Enter;
            TxtSx3.Leave += TextBox_Leave;
            TxtSy3.Enter += TextBox_Enter;
            TxtSy3.Leave += TextBox_Leave;
            TxtEx3.Enter += TextBox_Enter;
            TxtEx3.Leave += TextBox_Leave;
            TxtEy3.Enter += TextBox_Enter;
            TxtEy3.Leave += TextBox_Leave;
            TxtX1.Enter += TextBox_Enter;
            TxtX1.Leave += TextBox_Leave;
            TxtY1.Enter += TextBox_Enter;
            TxtY1.Leave += TextBox_Leave;
            TxtX2.Enter += TextBox_Enter;
            TxtX2.Leave += TextBox_Leave;
            TxtY2.Enter += TextBox_Enter;
            TxtY2.Leave += TextBox_Leave;
            TxtX3.Enter += TextBox_Enter;
            TxtX3.Leave += TextBox_Leave;
            TxtY3.Enter += TextBox_Enter;
            TxtY3.Leave += TextBox_Leave;

            // Nếu muốn áp dụng cho TxtSAPN:
             TxtSAPN.Enter += TextBox_Enter;
             TxtSAPN.Leave += TextBox_Leave;
        }
        private void TextBox_Enter(object sender, EventArgs e)
        {
            if (sender is TextBox textBox)
            {
                // Đặt màu nền khi nhận focus (ví dụ: màu vàng nhạt)
                textBox.BackColor = Color.PaleGoldenrod;
            }
        }

        private void TextBox_Leave(object sender, EventArgs e)
        {
            if (sender is TextBox textBox)
            {
                // Khôi phục màu nền mặc định khi mất focus
                textBox.BackColor = defaultTextBoxBackColor;
            }
        }
        #endregion

        #region CẬP NHẬT LOG STATUS
        // Thêm phương thức này vào class MainForm
        private void AppendToLog(string message, System.Drawing.Color color)
        {
            if (txtLog.InvokeRequired)
            {
                txtLog.Invoke(new Action(() => AppendToLog(message, color)));
                return;
            }

            // Lấy thời gian hiện tại theo định dạng mong muốn (ví dụ: [HH:mm:ss])
            string timestamp = $"[{DateTime.Now:dd/MM/yyyy | HH:mm:ss}] ";

            // Di chuyển con trỏ đến cuối văn bản
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.SelectionLength = 0; // Không chọn gì

            // --- Thêm phần thời gian với màu mặc định ---
            // (Màu mặc định là màu hiện tại của ForeColor, thường là đen hoặc màu bạn thiết lập)
            // Vì SelectionColor đang là màu mặc định (hoặc màu từ lần trước nếu không reset), nên ta có thể append luôn
            // Nhưng để chắc chắn, ta reset về màu ForeColor trước khi thêm timestamp
            txtLog.SelectionColor = txtLog.ForeColor; // Đặt màu về mặc định
            txtLog.AppendText(timestamp);

            // --- Thêm phần nội dung log với màu được chỉ định ---
            txtLog.SelectionColor = color; // Đặt màu theo tham số truyền vào
            txtLog.AppendText(message);

            // --- Thêm dấu ngắt dòng ---
            txtLog.SelectionColor = txtLog.ForeColor; // Dấu ngắt dòng cũng nên là màu mặc định
            txtLog.AppendText("\n"); // hoặc Environment.NewLine

            // Cuộn xuống cuối cùng
            txtLog.ScrollToCaret();
        }
        #endregion

        #region ĐỒNG HỒ CHẠY THEO GMT +7
        // Sự kiện cho đồng hồ
        private void Timer_Tick(object sender, EventArgs e)
        {
            try
            {
                // Lấy múi giờ GMT+7
                TimeZoneInfo vietnamZone = TimeZoneInfo.FindSystemTimeZoneById("SE Asia Standard Time");
                // Lấy thời gian hiện tại theo UTC (GMT+0)
                DateTime utcTime = DateTime.UtcNow;
                // Chuyển đổi thời gian UTC sang thời gian Việt Nam (GMT+7)
                DateTime vietnamTime = TimeZoneInfo.ConvertTimeFromUtc(utcTime, vietnamZone);

                // Hiển thị thời gian đã chuyển đổi
                LabelTime.Text = vietnamTime.ToString("HH:mm:ss");
                LabelDate.Text = vietnamTime.ToString("dd/MM/yyyy");
            }
            catch (TimeZoneNotFoundException ex)
            {
                // Xử lý ngoại lệ nếu không tìm thấy múi giờ "SE Asia Standard Time"
                LblStatus.ForeColor = System.Drawing.Color.DarkRed;
                LblStatus.Text = $"Lỗi múi giờ: {ex.Message}";
            }
            catch (Exception ex)
            {
                // Xử lý các lỗi khác
                LblStatus.ForeColor = System.Drawing.Color.DarkRed;
                LblStatus.Text = $"Lỗi cập nhật đồng hồ: {ex.Message}";
            }
        }
        #endregion

        #region ĐÓNG ỨNG DỤNG  
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Dừng auto update
            UpdateManager.StopAutoCheck();

            // Dừng và giải phóng timer cầu vồng
            if (rainbowTimer != null)
            {
                rainbowTimer.Stop();
                rainbowTimer.Dispose();
            }

            // Giải phóng tooltip
            validationToolTip?.Dispose();
            statusToolTip?.Dispose();

            base.OnFormClosing(e);
        }
        #endregion

        #region HIỆU ỨNG CHO CÁC NÚT BẤM
        private void InitializeButtonEffects()
        {
            Color keyboardBaseColor = System.Drawing.ColorTranslator.FromHtml("#FFF");

            // Xử lý nút SAVE (giữ nguyên màu ban đầu)
            originalColors.Add("BtnSave", BtnSave.BackColor);
            BtnSave.MouseEnter += Button_MouseEnter;
            BtnSave.MouseLeave += Button_MouseLeave;
            BtnSave.MouseDown += Button_MouseDown;
            BtnSave.MouseUp += Button_MouseUp;

            // Xử lý nút RESET (giữ nguyên màu ban đầu)
            originalColors.Add("BtnReset", BtnReset.BackColor);
            BtnReset.MouseEnter += Button_MouseEnter;
            BtnReset.MouseLeave += Button_MouseLeave;
            BtnReset.MouseDown += Button_MouseDown;
            BtnReset.MouseUp += Button_MouseUp;

            // Đặt lại màu cho nút ALL và xử lý riêng biệt
            BtnAll.BackColor = System.Drawing.ColorTranslator.FromHtml("#97FFFF");
            originalColors.Add("BtnAll", BtnAll.BackColor);
            BtnAll.MouseEnter += Button_MouseEnter;
            BtnAll.MouseLeave += Button_MouseLeave;
            BtnAll.MouseDown += Button_MouseDown;
            BtnAll.MouseUp += Button_MouseUp;

            // Áp dụng màu nền mới cho các nút trên bàn phím ảo (trừ nút ALL và BtnBack)
            Button[] keyboardButtons = { Btn0, Btn1, Btn2, Btn3, Btn4, Btn5, Btn6, Btn7, Btn8, Btn9, BtnBack };
            foreach (Button Btn in keyboardButtons)
            {
                // Chỉ áp dụng màu nền mới cho các nút số
                if (Btn.Name != "BtnBack")
                {
                    Btn.BackColor = keyboardBaseColor;
                }

                // Lưu màu nền hiện tại (đã được đổi) vào từ điển
                originalColors.Add(Btn.Name, Btn.BackColor);

                Btn.MouseEnter += Button_MouseEnter;
                Btn.MouseLeave += Button_MouseLeave;
                Btn.MouseDown += Button_MouseDown;
                Btn.MouseUp += Button_MouseUp;
            }
        }

        // Hiệu ứng khi di chuột vào nút
        private void Button_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button Btn)
            {
                // Áp dụng hiệu ứng di chuột cho TẤT CẢ các nút
                if (originalColors.ContainsKey(Btn.Name))
                {
                    Color originalColor = originalColors[Btn.Name];
                    int r = Math.Min(255, originalColor.R + 50);
                    int g = Math.Min(255, originalColor.G + 50);
                    int b = Math.Min(255, originalColor.B + 50);
                    Btn.BackColor = Color.FromArgb(r, g, b);
                }
            }
        }

        private void Button_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button Btn && originalColors.ContainsKey(Btn.Name))
            {
                // Khôi phục màu nền ban đầu
                Btn.BackColor = originalColors[Btn.Name];
            }
        }

        private void Button_MouseDown(object sender, EventArgs e)
        {
            if (sender is Button Btn)
            {
                // Áp dụng hiệu ứng khi nhấn chuột xuống cho TẤT CẢ các nút
                if (originalColors.ContainsKey(Btn.Name))
                {
                    Color originalColor = originalColors[Btn.Name];
                    int r = Math.Max(0, originalColor.R - 100);
                    int g = Math.Max(0, originalColor.G - 100);
                    int b = Math.Max(0, originalColor.B - 100);
                    Btn.BackColor = Color.FromArgb(r, g, b);
                }
            }
        }

        private void Button_MouseUp(object sender, EventArgs e)
        {
            if (sender is Button Btn)
            {
                if (Btn.ClientRectangle.Contains(Btn.PointToClient(Cursor.Position)))
                {
                    // Nếu nhả chuột trong vùng nút, khôi phục hiệu ứng di chuột vào
                    if (originalColors.ContainsKey(Btn.Name))
                    {
                        Color originalColor = originalColors[Btn.Name];
                        int r = Math.Min(255, originalColor.R + 30);
                        int g = Math.Min(255, originalColor.G + 30);
                        int b = Math.Min(255, originalColor.B + 30);
                        Btn.BackColor = Color.FromArgb(r, g, b);
                    }
                }
                else
                {
                    // Nếu nhả chuột ngoài vùng nút, khôi phục màu ban đầu
                    if (originalColors.ContainsKey(Btn.Name))
                    {
                        Btn.BackColor = originalColors[Btn.Name];
                    }
                }
            }
        }
        #endregion

        #region XỬ LÝ BÀN PHÍM ẢO VÀ KIỂM TRA DỮ LIỆU
        // Hàm này được gọi trong constructor
        private void InitializeKeyboardEvents()
        {
            // Gán sự kiện Click để xác định TextBox hiện tại
            TxtSAPN.Click += TextBox_Click;
            TxtSx1.Click += TextBox_Click;
            TxtSy1.Click += TextBox_Click;
            TxtEx1.Click += TextBox_Click;
            TxtEy1.Click += TextBox_Click;
            TxtSx2.Click += TextBox_Click;
            TxtSy2.Click += TextBox_Click;
            TxtEx2.Click += TextBox_Click;
            TxtEy2.Click += TextBox_Click;
            TxtSx3.Click += TextBox_Click;
            TxtSy3.Click += TextBox_Click;
            TxtEx3.Click += TextBox_Click;
            TxtEy3.Click += TextBox_Click;
            TxtX1.Click += TextBox_Click;
            TxtY1.Click += TextBox_Click;
            TxtX2.Click += TextBox_Click;
            TxtY2.Click += TextBox_Click;
            TxtX3.Click += TextBox_Click;
            TxtY3.Click += TextBox_Click;

            // Gán sự kiện KeyDown cho TxtSAPN
            TxtSAPN.KeyDown += TxtSAPN_KeyDown;
            // Gán sự kiện KeyPress cho TxtSAPN
            TxtSAPN.KeyPress += TxtSAPN_KeyPress;
            // Gán sự kiện TextChanged để cảnh báo sớm
            TxtSAPN.TextChanged += TxtSAPN_TextChanged;
            // Gán sự kiện KeyPress riêng cho TxtSx1 (cho phép ALL)
            TxtSx1.KeyPress += TxtSx1_KeyPress;

            // Gán sự kiện KeyPress chung cho các ô tọa độ còn lại
            TxtSy1.KeyPress += CoordinateTextBox_KeyPress;
            TxtEx1.KeyPress += CoordinateTextBox_KeyPress;
            TxtEy1.KeyPress += CoordinateTextBox_KeyPress;
            TxtSx2.KeyPress += CoordinateTextBox_KeyPress;
            TxtSy2.KeyPress += CoordinateTextBox_KeyPress;
            TxtEx2.KeyPress += CoordinateTextBox_KeyPress;
            TxtEy2.KeyPress += CoordinateTextBox_KeyPress;
            TxtSx3.KeyPress += CoordinateTextBox_KeyPress;
            TxtSy3.KeyPress += CoordinateTextBox_KeyPress;
            TxtEx3.KeyPress += CoordinateTextBox_KeyPress;
            TxtEy3.KeyPress += CoordinateTextBox_KeyPress;
            TxtX1.KeyPress += CoordinateTextBox_KeyPress;
            TxtY1.KeyPress += CoordinateTextBox_KeyPress;
            TxtX2.KeyPress += CoordinateTextBox_KeyPress;
            TxtY2.KeyPress += CoordinateTextBox_KeyPress;
            TxtX3.KeyPress += CoordinateTextBox_KeyPress;
            TxtY3.KeyPress += CoordinateTextBox_KeyPress;
        }

        private void TxtSAPN_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox currentTextBox = (TextBox)sender;

            // Kiểm tra nếu độ dài văn bản >= 89 và không phải phím điều khiển
            if (!char.IsControl(e.KeyChar) && currentTextBox.Text.Length >= 89)
            {
                e.Handled = true; // Ngăn nhập thêm ký tự
                validationToolTip.ToolTipTitle = "Lỗi nhập liệu";
                validationToolTip.Show("Chỉ cho phép nhập tối đa 89 ký tự", currentTextBox, 0, currentTextBox.Height, 3000);
                return;
            }
        }

        private void TxtSAPN_TextChanged(object sender, EventArgs e)
        {
            int remaining = 89 - TxtSAPN.Text.Length;
            if (remaining <= 5 && remaining >= 0)
            {
                validationToolTip.Show($"Còn {remaining} ký tự", TxtSAPN, 0, TxtSAPN.Height, 1000);
            }
            // Kiểm tra nếu độ dài văn bản đạt 89 ký tự
            if (TxtSAPN.Text.Length == 89)
            {
                // Chuyển focus sang TxtSx1
                TxtSx1.Focus();
                // (Tùy chọn) Nếu bạn cũng muốn chọn (select) toàn bộ văn bản trong TxtSx1 khi focus,
                // bạn có thể thêm: TxtSx1.SelectAll();
            }
        }

        // Khi chuột di vào nhãn, kích hoạt hiệu ứng cầu vồng
        private void LblCopyright_MouseEnter(object sender, EventArgs e)
        {
            if (!isRainbowActive)
            {
                isRainbowActive = true;
                originalCopyrightColor = LblCopyright.ForeColor;
                rainbowTimer.Start();
            }
        }

        // Khi chuột rời nhãn, tắt hiệu ứng và khôi phục màu gốc
        private void LblCopyright_MouseLeave(object sender, EventArgs e)
        {
            if (isRainbowActive)
            {
                isRainbowActive = false;
                rainbowTimer.Stop();
                LblCopyright.ForeColor = originalCopyrightColor;
            }
        }
        #endregion

        #region THAY ĐỔI MÀU SẮC KHI DI CHUỘT VÀO TÊN TÁC GIẢ
        // Sự kiện Tick của timer, cập nhật màu sắc
        private void RainbowTimer_Tick(object sender, EventArgs e)
        {
            rainbowPhase += 0.05; // Giảm tốc độ thay đổi để màu chuyển từ từ hơn

            Color newColor = CalculateRainbowColor(rainbowPhase);
            LblCopyright.ForeColor = newColor;
        }

        // Tính toán màu sắc cầu vồng dựa trên giai đoạn
        private Color CalculateRainbowColor(double phase)
        {
            double red = Math.Sin(phase) * 127 + 128;
            double green = Math.Sin(phase + 2 * Math.PI / 3) * 127 + 128;
            double blue = Math.Sin(phase + 4 * Math.PI / 3) * 127 + 128;

            red = Math.Max(0, Math.Min(255, red));
            green = Math.Max(0, Math.Min(255, green));
            blue = Math.Max(0, Math.Min(255, blue));

            return Color.FromArgb((int)red, (int)green, (int)blue);
        }
        private void TextBox_Click(object sender, EventArgs e)
        {
            currentTextBox = (TextBox)sender;
            currentTextBox.Focus(); // Đảm bảo TextBox được focus
        }
        #endregion

        #region KIỂM TRA ĐÃ NHẬP DỮ LIÊU HAY CHƯA
        // Phương thức xử lý sự kiện Click cho LblStatus
        private void LblStatus_Click(object sender, EventArgs e)
        {
            // Kiểm tra xem đã có đường dẫn file được lưu chưa
            if (!string.IsNullOrEmpty(_lastSavedFilePath))
            {
                try
                {
                    // Lấy đường dẫn thư mục chứa file
                    string folderPath = Path.GetDirectoryName(_lastSavedFilePath);
                    // Mở thư mục bằng Windows Explorer
                    Process.Start("explorer.exe", folderPath);
                }
                catch (Exception ex)
                {
                    // Báo lỗi nếu không thể mở thư mục
                    MessageBox.Show($"Không thể mở thư mục. Lỗi: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void LblSAPNCount_Click(object sender, EventArgs e)
        {
            // Kiểm tra xem đã có đường dẫn file được lưu chưa
            if (!string.IsNullOrEmpty(_lastSavedFilePath))
            {
                try
                {
                    // Lấy đường dẫn thư mục chứa file
                    string folderPath = Path.GetDirectoryName(_lastSavedFilePath);
                    // Mở thư mục bằng Windows Explorer
                    Process.Start("explorer.exe", folderPath);
                }
                catch (Exception ex)
                {
                    // Báo lỗi nếu không thể mở thư mục
                    MessageBox.Show($"Không thể mở thư mục. Lỗi: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // Phương thức xử lý sự kiện KeyDown của TxtSAPN
        private void TxtSAPN_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                TxtSx1.Focus();
                currentTextBox = TxtSx1; // Cập nhật currentTextBox để bàn phím ảo tương tác với TxtSx1
                e.SuppressKeyPress = true;
            }
        }

        // Xử lý các nút số trên bàn phím ảo
        private void BtnNumber_Click(object sender, EventArgs e)
        {
            if (currentTextBox != null)
            {
                Button Btn = (Button)sender;
                string buttonText = Btn.Text;

                // Kiểm tra nếu ô nhập là TxtSAPN và đã đạt giới hạn 89 ký tự
                if (currentTextBox == TxtSAPN && currentTextBox.Text.Length >= 89)
                {
                    validationToolTip.ToolTipTitle = "Lỗi nhập liệu";
                    validationToolTip.Show("Chỉ cho phép nhập tối đa 89 ký tự", currentTextBox, 0, currentTextBox.Height, 3000);
                    return;
                }

                // Kiểm tra nếu ô nhập đã đạt giới hạn 3 ký tự
                if (currentTextBox.Text.Length >= 3 && !buttonText.Equals("All", StringComparison.OrdinalIgnoreCase))
                {
                    // Trừ trường hợp người dùng xóa hoặc nhập "ALL"
                    validationToolTip.ToolTipTitle = "Lỗi nhập liệu";
                    validationToolTip.Show("Chỉ cho phép nhập tối đa 3 số", currentTextBox, 0, currentTextBox.Height, 2000);
                    return;
                }

                // Chỉ cho phép nhập "All" vào TxtSx1
                if (buttonText.Equals("All", StringComparison.OrdinalIgnoreCase))
                {
                    if (currentTextBox != TxtSx1 || !string.IsNullOrWhiteSpace(currentTextBox.Text))
                    {
                        // Hiển thị tooltip nếu cố gắng nhập "ALL" vào ô khác
                        if (currentTextBox != TxtSx1)
                        {
                            validationToolTip.ToolTipTitle = "Lỗi nhập liệu";
                            validationToolTip.Show("Chỉ cho phép nhập ALL vào ô Sx1", currentTextBox, 0, currentTextBox.Height, 2000);
                        }
                        return; // Không làm gì nếu không phải TxtSx1 hoặc ô đã có nội dung
                    }
                    currentTextBox.Text = buttonText;
                }
                // Cho phép nhập số vào các ô
                else
                {
                    currentTextBox.Text += buttonText;
                }
            }
        }

        // Xử lý nút xóa (Back)
        private void BtnBack_Click(object sender, EventArgs e)
        {
            if (currentTextBox != null && currentTextBox.Text.Length > 0)
            {
                currentTextBox.Text = currentTextBox.Text.Substring(0, currentTextBox.Text.Length - 1);
            }
        }

        // Kiểm tra dữ liệu sAPN và ít nhất một trường tọa độ
        private bool IsDataValid()
        {
            if (string.IsNullOrWhiteSpace(TxtSAPN.Text))
            {
                return false;
            }

            // Kiểm tra xem ít nhất một trong các ô tọa độ có dữ liệu không
            if (string.IsNullOrWhiteSpace(TxtSx1.Text) && string.IsNullOrWhiteSpace(TxtSy1.Text) && string.IsNullOrWhiteSpace(TxtEx1.Text) && string.IsNullOrWhiteSpace(TxtEy1.Text) &&
                string.IsNullOrWhiteSpace(TxtSx2.Text) && string.IsNullOrWhiteSpace(TxtSy2.Text) && string.IsNullOrWhiteSpace(TxtEx2.Text) && string.IsNullOrWhiteSpace(TxtEy2.Text) &&
                string.IsNullOrWhiteSpace(TxtSx3.Text) && string.IsNullOrWhiteSpace(TxtSy3.Text) && string.IsNullOrWhiteSpace(TxtEx3.Text) && string.IsNullOrWhiteSpace(TxtEy3.Text) &&
                string.IsNullOrWhiteSpace(TxtX1.Text) && string.IsNullOrWhiteSpace(TxtY1.Text) && string.IsNullOrWhiteSpace(TxtX2.Text) && string.IsNullOrWhiteSpace(TxtY2.Text) &&
                string.IsNullOrWhiteSpace(TxtX3.Text) && string.IsNullOrWhiteSpace(TxtY3.Text))
            {
                return false;
            }
            return true;
        }

        #endregion

        #region SƯ KIỆN BẤM NÚT SAVE VÀ RESET
        // Xử lý nút SAVE
        private void BtnSave_Click(object sender, EventArgs e)
        {
            // Kiểm tra xem đã chọn loại lỗi chưa
            if (cboErrorType.SelectedIndex == -1)
            {
                validationToolTip.ToolTipIcon = ToolTipIcon.Warning;
                validationToolTip.ToolTipTitle = "Lỗi";
                validationToolTip.Show("Vui lòng chọn loại lỗi!", cboErrorType, 0, cboErrorType.Height, 5000);
                return;
            }
            // Kiểm tra xem có mục nào được chọn không
            if (cboModel.SelectedIndex == -1 || cboModel.SelectedItem == null)
            {
                AppendToLog("Lỗi: Chưa chọn model.", System.Drawing.Color.Red);
                validationToolTip.ToolTipIcon = ToolTipIcon.Warning;
                validationToolTip.ToolTipTitle = "Lỗi";
                validationToolTip.Show("Vui lòng chọn model!", cboModel, 0, cboModel.Height, 5000);
                return; // Thoát nếu không có model nào được chọn
            }
            // Kiểm tra dữ liệu sAPN trước khi lưu
            if (!IsDataValid())
            {
                validationToolTip.ToolTipIcon = ToolTipIcon.Warning;
                validationToolTip.ToolTipTitle = "Lỗi";
                validationToolTip.Show("Vui lòng nhập sAPN và ít nhất một trong số các ô tọa độ!", TxtSAPN, 0, TxtSAPN.Height, 5000);
                return;
            }
            
            // Lấy múi giờ GMT+7
            TimeZoneInfo vietnamZone = TimeZoneInfo.FindSystemTimeZoneById("SE Asia Standard Time");
            DateTime vietnamTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, vietnamZone);

            // Xác định ngày và ca làm việc
            string dateString;
            string shift;

            // Nếu thời gian từ 20:00 hôm nay đến trước 08:00 hôm sau, sử dụng ngày bắt đầu ca đêm
            if (vietnamTime.Hour >= 20 || vietnamTime.Hour < 8)
            {
                // Nếu thời gian từ 00:00 đến 07:59:59, sử dụng ngày hôm trước
                if (vietnamTime.Hour < 8)
                {
                    dateString = vietnamTime.AddDays(-1).ToString("yyyyMMdd");
                }
                else
                {
                    dateString = vietnamTime.ToString("yyyyMMdd");
                }
                shift = "NIGHT";
            }
            // Nếu thời gian từ 08:00 đến trước 20:00, sử dụng ngày hiện tại và ca ngày
            else
            {
                dateString = vietnamTime.ToString("yyyyMMdd");
                shift = "DAY";
            }

            string eqpid = ReadEQPIDFromIniFile(); // Lấy EQPID từ file ini
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string appFolderPath = Path.Combine(desktopPath, "IT_CONFIRM");
            string fileName = string.IsNullOrEmpty(eqpid) ? $"IT_{dateString}_{shift}.csv" : $"IT_{eqpid}_{dateString}_{shift}.csv";
            string filePath = Path.Combine(appFolderPath, fileName);
            string timestamp = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            try
            {
                if (!Directory.Exists(appFolderPath))
                {
                    Directory.CreateDirectory(appFolderPath);
                }

                bool fileExists = File.Exists(filePath);
                if (!fileExists)
                {
                    string header = "MODEL,sAPN,DESCRIPTION,Sx1,Sy1,Ex1,Ey1,Sx2,Sy2,Ex2,Ey2,Sx3,Sy3,Ex3,Ey3,X1,Y1,X2,Y2,X3,Y3,EVENT_TIME";
                    File.AppendAllText(filePath, header + Environment.NewLine, System.Text.Encoding.UTF8);
                }

                // Lấy model đã chọn từ ComboBox
                string selectedModel = cboModel.SelectedItem.ToString();
                // Lấy loại lỗi đã chọn từ ComboBox
                string selectedErrorType = cboErrorType.SelectedItem.ToString();

                string csvData = $"{selectedModel},{TxtSAPN.Text},{selectedErrorType},{TxtSx1.Text},{TxtSy1.Text},{TxtEx1.Text},{TxtEy1.Text}," +
                                 $"{TxtSx2.Text},{TxtSy2.Text},{TxtEx2.Text},{TxtEy2.Text}," +
                                 $"{TxtSx3.Text},{TxtSy3.Text},{TxtEx3.Text},{TxtEy3.Text}," +
                                 $"{TxtX1.Text},{TxtY1.Text},{TxtX2.Text},{TxtY2.Text},{TxtX3.Text},{TxtY3.Text},{timestamp}";

                File.AppendAllText(filePath, csvData + Environment.NewLine, System.Text.Encoding.UTF8);
                // Cập nhật thông báo thành công
                _lastSavedFilePath = filePath; // Lưu đường dẫn file vào biến toàn cục
                // Cập nhật thông báo thành công
                LblStatus.ForeColor = System.Drawing.Color.Green;
                LblStatus.Text = $"Lưu thành công! \nDữ liệu đã được ghi lại lúc: {timestamp}\nVị trí lưu: {filePath}";
                AppendToLog($"Dữ liệu đã được ghi lại lúc: {timestamp} \nVị trí lưu: {filePath}", System.Drawing.Color.DarkGreen);
                // Thiết lập tooltip cho LblStatus
                statusToolTip.SetToolTip(LblStatus, "BẤM VÀO ĐÂY ĐỂ MỞ VỊ TRÍ LƯU FILE");

                // Xóa nội dung của tất cả các TextBox sau khi lưu thành công
                TxtSAPN.Clear();
                TxtSx1.Clear();
                TxtSy1.Clear();
                TxtEx1.Clear();
                TxtEy1.Clear();
                TxtSx2.Clear();
                TxtSy2.Clear();
                TxtEx2.Clear();
                TxtEy2.Clear();
                TxtSx3.Clear();
                TxtSy3.Clear();
                TxtEx3.Clear();
                TxtEy3.Clear();
                TxtX1.Clear();
                TxtY1.Clear();
                TxtX2.Clear();
                TxtY2.Clear();
                TxtX3.Clear();
                TxtY3.Clear();
                //cboErrorType.SelectedIndex = 0; // Mặc định chọn "ĐỐM" =-1 KHÔNG CHỌN GÌ

                // Đặt focus lại cho ô đầu tiên
                TxtSAPN.Focus();
                // Cập nhật bộ đếm sau khi lưu
                UpdateSavedSAPNCount();

                // Chạy lưu NAS async (background) để không block UI
                _ = Task.Run(() =>
                {
                    bool nasSaved = false;
                    string nasError = null;
                    string successfulNasPath = null;

                    // Lấy danh sách thông tin xác thực NAS
                    var nasCredentialsList = ReadNASCredentialsFromIniFile();

                    // Thử từng server NAS trong danh sách
                    foreach (var (nasPath, credentials) in nasCredentialsList)
                    {
                        try
                        {
                            using (var connection = new NetworkConnection(nasPath, credentials))
                            {
                                string nasSubDirectory = string.IsNullOrEmpty(ReadEQPIDFromIniFile()) ? "NoEQPID" : ReadEQPIDFromIniFile();
                                string nasSubDirectoryPath = Path.Combine(nasPath, nasSubDirectory);
                                string nasFilePath = Path.Combine(nasSubDirectoryPath, Path.GetFileName(_lastSavedFilePath));

                                // Đảm bảo thư mục NAS chính tồn tại
                                if (!Directory.Exists(nasPath))
                                {
                                    Directory.CreateDirectory(nasPath);
                                }

                                // Đảm bảo thư mục con EQPID tồn tại
                                if (!Directory.Exists(nasSubDirectoryPath))
                                {
                                    Directory.CreateDirectory(nasSubDirectoryPath);
                                }

                                // Tạo hoặc thêm vào file NAS
                                if (!File.Exists(nasFilePath))
                                {
                                    string header = "MODEL,sAPN,DESCRIPTION,Sx1,Sy1,Ex1,Ey1,Sx2,Sy2,Ex2,Ey2,Sx3,Sy3,Ex3,Ey3,X1,Y1,X2,Y2,X3,Y3,EVENT_TIME";
                                    File.AppendAllText(nasFilePath, header + Environment.NewLine, System.Text.Encoding.UTF8);
                                }
                                File.AppendAllText(nasFilePath, csvData + Environment.NewLine, System.Text.Encoding.UTF8);

                                // Nếu lưu thành công, đánh dấu và lưu đường dẫn
                                nasSaved = true;
                                successfulNasPath = nasFilePath;
                                break; // Thoát vòng lặp khi lưu thành công
                            }
                        }
                        catch (Win32Exception winEx)
                        {
                            nasError = $"Lỗi {winEx.NativeErrorCode} ({nasPath}): {winEx.Message}";
                        }
                        catch (Exception ex)
                        {
                            nasError = $"Lỗi ({nasPath}): {ex.Message}";
                        }
                    }

                    // Cập nhật UI trên luồng chính
                    this.Invoke(new Action(() =>
                    {
                        if (nasSaved)
                        {
                            LblStatusNas.ForeColor = System.Drawing.Color.Green;
                            LblStatusNas.Text = $"NAS Server: {successfulNasPath}";
                            AppendToLog($"NAS Server: {successfulNasPath}", System.Drawing.Color.DarkCyan);
                            statusToolTip.SetToolTip(LblStatusNas, "BẤM VÀO ĐÂY ĐỂ MỞ VỊ TRÍ LƯU FILE");
                        }
                        else
                        {
                            LblStatusNas.ForeColor = System.Drawing.Color.Red;
                            LblStatusNas.Text = $"NAS Server: {nasError}";
                            AppendToLog($"NAS Server: {nasError}", System.Drawing.Color.DarkOrchid);
                            statusToolTip.SetToolTip(LblStatusNas, "Kiểm tra lại NASConfig.ini hoặc tài khoản NAS");
                        }
                    }));
                });
            }
            catch (IOException)
            {
                // Cập nhật thông báo lỗi cụ thể khi file đang được mở
                LblStatus.ForeColor = System.Drawing.Color.Red;
                LblStatus.Text = "File đang được mở bởi ứng dụng khác hoặc không thể ghi dữ liệu.\nHãy đóng file đang mở trước khi bấm Save";
                AppendToLog($"File đang được mở bởi ứng dụng khác hoặc không thể ghi dữ liệu.Hãy đóng file đang mở trước khi bấm Save", System.Drawing.Color.Green);
                // Xóa tooltip khi có lỗi
                statusToolTip.SetToolTip(LblStatus, "");
            }
            catch (Exception ex)
            {
                // Báo lỗi chung nếu có lỗi khác
                LblStatus.ForeColor = System.Drawing.Color.Red;
                LblStatus.Text = $"Đã xảy ra lỗi: {ex.Message}";
                AppendToLog($"Đã xảy ra lỗi: {ex.Message}", System.Drawing.Color.Red);
                // Xóa tooltip khi có lỗi
                statusToolTip.SetToolTip(LblStatus, "");
            }
        }

        // Xử lý nút RESET
        private void BtnReset_Click(object sender, EventArgs e)
        {
            // Xóa nội dung của tất cả các TextBox
            TxtSAPN.Clear();
            TxtSx1.Clear();
            TxtSy1.Clear();
            TxtEx1.Clear();
            TxtEy1.Clear();
            TxtSx2.Clear();
            TxtSy2.Clear();
            TxtEx2.Clear();
            TxtEy2.Clear();
            TxtSx3.Clear();
            TxtSy3.Clear();
            TxtEx3.Clear();
            TxtEy3.Clear();
            TxtX1.Clear();
            TxtY1.Clear();
            TxtX2.Clear();
            TxtY2.Clear();
            TxtX3.Clear();
            TxtY3.Clear();
            //cboErrorType.SelectedIndex = -1; // Mặc định chọn "ĐỐM" =0 CHỌN DÒNG ĐẦU TIÊN

            // Cập nhật thông báo
            LblStatus.ForeColor = System.Drawing.Color.DarkOrange;
            LblStatus.Text = "Đã khởi tạo lại ứng dụng.";
            AppendToLog($"Đã khởi tạo lại ứng dụng.", System.Drawing.Color.Chocolate);
            // Xóa tooltip khi reset ứng dụng
            statusToolTip.SetToolTip(LblStatus, "");
            // Reset trạng thái NAS
            LblStatusNas.ForeColor = System.Drawing.Color.DarkOrange;
            LblStatusNas.Text = "";
            statusToolTip.SetToolTip(LblStatusNas, "");
            // Đặt focus lại cho ô đầu tiên
            TxtSAPN.Focus();
        }

        private void CoordinateTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox currentTextBox = (TextBox)sender;

            // Kiểm tra và ngăn không cho nhập nếu độ dài đã đạt 3
            if (!char.IsControl(e.KeyChar) && currentTextBox.Text.Length >= 3)
            {
                e.Handled = true;
                validationToolTip.ToolTipTitle = "Lỗi nhập liệu";
                validationToolTip.Show("Chỉ cho phép nhập tối đa 3 số", currentTextBox, 0, currentTextBox.Height, 3000);
                return;
            }

            // Cho phép nhập số và các ký tự điều khiển (như Backspace)
            // Ngăn các ký tự khác (chữ cái, ký tự đặc biệt)
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
                validationToolTip.ToolTipTitle = "Lỗi nhập liệu";
                validationToolTip.Show("Chỉ cho phép nhập số!", currentTextBox, 0, currentTextBox.Height, 3000);
            }
        }

        private void TxtSx1_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox currentTextBox = (TextBox)sender;
            string currentText = currentTextBox.Text;

            // Nếu người dùng nhập "ALL" và muốn thêm ký tự, hãy ngăn chặn
            if (currentText.ToUpper() == "ALL" && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
                validationToolTip.ToolTipTitle = "Lỗi nhập liệu";
                validationToolTip.Show("Không thể nhập thêm sau 'ALL'", currentTextBox, 0, currentTextBox.Height, 2000);
                return;
            }

            // Kiểm tra giới hạn 3 ký tự
            if (!char.IsControl(e.KeyChar) && currentText.Length >= 3)
            {
                // Nếu ký tự mới là số và độ dài đã đạt 3, ngăn chặn
                if (char.IsDigit(e.KeyChar))
                {
                    e.Handled = true;
                    validationToolTip.ToolTipTitle = "Lỗi nhập liệu";
                    validationToolTip.Show("Chỉ cho phép nhập tối đa 3 số", currentTextBox, 0, currentTextBox.Height, 2000);
                    return;
                }
            }

            // Cho phép nhập số, Backspace, và các ký tự 'a', 'A', 'l', 'L'
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && "ALL".IndexOf(char.ToUpper(e.KeyChar)) < 0)
            {
                e.Handled = true;
                validationToolTip.ToolTipTitle = "Lỗi nhập liệu";
                validationToolTip.Show("Chỉ cho phép nhập số hoặc ALL", currentTextBox, 0, currentTextBox.Height, 2000);
            }
        }

        // Phương thức mới để đếm và cập nhật số lượng sAPN đã lưu
        private void UpdateSavedSAPNCount()
        {
            // Sử dụng đường dẫn file cuối cùng được lưu
            string filePath = _lastSavedFilePath;

            // Nếu chưa có file nào được lưu (ví dụ: lần đầu chạy), thì thoát
            if (string.IsNullOrEmpty(filePath))
            {
                LblSAPNCount.Text = "Số lượng APN đã lưu: 0";
                AppendToLog("Chưa có file nào được lưu, không thể đếm.", System.Drawing.Color.Gray);
                return;
            }

            // Kiểm tra xem file có tồn tại không (tránh lỗi nếu file bị xóa)
            if (!File.Exists(filePath))
            {
                LblSAPNCount.Text = "Số lượng APN đã lưu: 0";
                AppendToLog($"File CSV không tồn tại, không thể đếm: {filePath}", System.Drawing.Color.Gray);
                return;
            }

            // Kiểm tra xem file có bị khóa bởi tiến trình khác không
            if (!CanReadFile(filePath)) // Giả sử bạn có hàm CanReadFile như trước
            {
                LblSAPNCount.Text = "Số lượng APN đã lưu: 0";
                AppendToLog($"Không thể đọc file CSV (có thể đang mở): {filePath}", System.Drawing.Color.Sienna);
                return;
            }

            int count = 0;
            try
            {
                using (var reader = new StreamReader(filePath, System.Text.Encoding.UTF8))
                {
                    string headerLine = reader.ReadLine();
                    if (headerLine == null)
                    {
                        LblSAPNCount.Text = "Số lượng APN đã lưu: 0";
                        AppendToLog($"File CSV trống, không thể đếm: {filePath}", System.Drawing.Color.Gray);
                        return; // Thoát nếu file rỗng
                    }

                    // Xác định định dạng dựa trên tiêu đề (có thể vẫn cần nếu xử lý định dạng cũ/khác sau này)
                    string[] headers = headerLine.Split(',');
                    bool isNewFormat = headers.Any(h => h.Contains("\"X1,Y1\""));

                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        var parts = line.Split(',');
                        // --- CHỈ KIỂM TRA CỘT MODEL và sAPN ---
                        if (parts.Length >= 2 && !string.IsNullOrWhiteSpace(parts[0]) && !string.IsNullOrWhiteSpace(parts[1]))
                        {
                            count++; // Tăng bộ đếm nếu cả MODEL và sAPN đều có dữ liệu
                        }
                        // Bỏ qua các dòng không đủ điều kiện trên
                    }
                }

                LblSAPNCount.Text = $"Số lượng APN đã lưu: {count}";
            }
            catch (IOException ex) // Bắt lỗi cụ thể liên quan đến I/O
            {
                LblSAPNCount.Text = "Số lượng APN đã lưu: 0";
                AppendToLog($"Lỗi I/O khi đếm APN từ file {filePath}: {ex.Message}", System.Drawing.Color.Sienna);
            }
            catch (Exception ex) // Bắt các lỗi khác
            {
                LblSAPNCount.Text = "Số lượng APN đã lưu: 0";
                AppendToLog($"Lỗi không mong đợi khi đếm APN từ file {filePath}: {ex.Message}", System.Drawing.Color.Violet);
            }
        }

        // Phương thức kiểm tra quyền đọc file
        private bool CanReadFile(string filePath)
        {
            try
            {
                using (FileStream fs = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }
        #endregion

        #region TẠO ĐƯỜNG DẪN FILE CỤC BỘ VÀ NAS
        // Phương thức tạo đường dẫn file cục bộ và NAS
        private void SetFilePath()
        {
            // Lấy múi giờ GMT+7
            TimeZoneInfo vietnamZone = TimeZoneInfo.FindSystemTimeZoneById("SE Asia Standard Time");
            DateTime vietnamTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, vietnamZone);

            // Xác định ngày và ca làm việc
            string dateString;
            string shift;

            if (vietnamTime.Hour >= 20 || vietnamTime.Hour < 8)
            {
                if (vietnamTime.Hour < 8)
                {
                    dateString = vietnamTime.AddDays(-1).ToString("yyyyMMdd");
                }
                else
                {
                    dateString = vietnamTime.ToString("yyyyMMdd");
                }
                shift = "NIGHT";
            }
            else
            {
                dateString = vietnamTime.ToString("yyyyMMdd");
                shift = "DAY";
            }

            string eqpid = ReadEQPIDFromIniFile();
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string directoryPath = Path.Combine(desktopPath, "IT_CONFIRM");
            string fileName = string.IsNullOrEmpty(eqpid) ? $"IT_{dateString}_{shift}.csv" : $"IT_{eqpid}_{dateString}_{shift}.csv";
            _lastSavedFilePath = Path.Combine(directoryPath, fileName);

            // Đảm bảo thư mục cục bộ tồn tại
            try
            {
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }
            }
            catch (Exception ex)
            {
                LblStatus.ForeColor = System.Drawing.Color.Red;
                LblStatus.Text = $"Lỗi tạo thư mục cục bộ: {ex.Message}";
            }
        }
        #endregion

        #region ĐỌC THÔNG TIN TỪ FILE INI
        /// Đọc giá trị EQPID từ file MachineParam.ini
        private string ReadEQPIDFromIniFile()
        {
            string filePath = @"C:\samsung\Debug\Config\MachineParam.ini";
            if (File.Exists(filePath))
            {
                try
                {
                    string[] lines = File.ReadAllLines(filePath);
                    foreach (string line in lines)
                    {
                        if (line.StartsWith("EQPID="))
                        {
                            return line.Substring("EQPID=".Length);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Xử lý lỗi khi đọc file, ví dụ: không có quyền truy cập
                    //MessageBox.Show("Lỗi khi đọc file MachineParam.ini: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    AppendToLog($"Lỗi khi đọc file MachineParam.ini:   {ex.Message}", System.Drawing.Color.Green);
                }
            }
            return "";
        }
        //Đọc thông tin NAS từ file NASConfig.ini
        private List<(string NasPath, NetworkCredential Credentials)> ReadNASCredentialsFromIniFile()
        {
            string filePath = @"C:\IT_CONFIRM\Config\NASConfig.ini";
            List<(string NasPath, NetworkCredential Credentials)> nasCredentialsList = new List<(string, NetworkCredential)>();

            // Mặc định cho ba khối NAS SERVER
            var defaultNasServers = new[]
            {
                new
                {
                    NasPath = @"\\107.126.41.111\IT_CONFIRM",
                    NasUser = "admin",
                    NasPassword = "insp2019@",
                    NasDomain = ""
                },
                new
                {
                    NasPath = @"",
                    NasUser = "",
                    NasPassword = "",
                    NasDomain = ""
                },
                new
                {
                    NasPath = @"",
                    NasUser = "",
                    NasPassword = "",
                    NasDomain = ""
                }
            };
            try
            {
                // Đảm bảo thư mục tồn tại
                string directoryPath = Path.GetDirectoryName(filePath);
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }

                // Kiểm tra quyền ghi vào thư mục
                if (!IsDirectoryWritable(directoryPath))
                {
                    LblStatus.ForeColor = System.Drawing.Color.Red;
                    LblStatus.Text = $"Không có quyền ghi vào thư mục {directoryPath} để tạo NASConfig.ini";
                    AppendToLog($"Không có quyền ghi vào thư mục {directoryPath} để tạo NASConfig.ini", System.Drawing.Color.Red);
                    // Trả về danh sách mặc định nếu không thể tạo file
                    foreach (var server in defaultNasServers)
                    {
                        nasCredentialsList.Add((server.NasPath, new NetworkCredential(server.NasUser, server.NasPassword, server.NasDomain)));
                    }
                    nasDirectoryPath = defaultNasServers[0].NasPath; // Gán mặc định
                    return nasCredentialsList;
                }

                // Nếu file chưa tồn tại, tạo file với ba khối [NAS SERVER]
                if (!File.Exists(filePath))
                {
                    StringBuilder iniContent = new StringBuilder();
                    for (int i = 1; i <= 3; i++)
                    {
                        var server = defaultNasServers[i - 1];
                        iniContent.AppendLine($"[NAS SERVER {i}]");
                        iniContent.AppendLine($"NASPATH={server.NasPath}");
                        iniContent.AppendLine($"NASUSER={server.NasUser}");
                        iniContent.AppendLine($"NASPASSWORD={server.NasPassword}");
                        iniContent.AppendLine($"NASDOMAIN={server.NasDomain}");
                        iniContent.AppendLine();
                    }
                    File.WriteAllText(filePath, iniContent.ToString(), Encoding.UTF8);
                    AppendToLog($"File cấu hình NAS mặc định đã được tạo tại: {filePath}", System.Drawing.Color.Blue);
                }

                // Đọc file
                string[] lines = File.ReadAllLines(filePath);
                string currentNasPath = null;
                string currentNasUser = null;
                string currentNasPassword = null;
                string currentNasDomain = null;
                string currentSection = null;

                foreach (string line in lines)
                {
                    string trimmedLine = line.Trim();
                    if (string.IsNullOrEmpty(trimmedLine))
                        continue;

                    // Kiểm tra section
                    if (trimmedLine.StartsWith("[NAS SERVER"))
                    {
                        // Nếu đã có dữ liệu từ section trước, thêm vào danh sách
                        if (currentNasPath != null && currentNasUser != null && currentNasPassword != null)
                        {
                            nasCredentialsList.Add((currentNasPath, new NetworkCredential(currentNasUser, currentNasPassword, currentNasDomain ?? "")));
                        }
                        currentSection = trimmedLine;
                        currentNasPath = null;
                        currentNasUser = null;
                        currentNasPassword = null;
                        currentNasDomain = null;
                        continue;
                    }

                    // Đọc các khóa trong section [NAS SERVER]
                    if (currentSection != null && currentSection.StartsWith("[NAS SERVER"))
                    {
                        if (trimmedLine.StartsWith("NASPATH="))
                            currentNasPath = trimmedLine.Substring("NASPATH=".Length);
                        else if (trimmedLine.StartsWith("NASUSER="))
                            currentNasUser = trimmedLine.Substring("NASUSER=".Length);
                        else if (trimmedLine.StartsWith("NASPASSWORD="))
                            currentNasPassword = trimmedLine.Substring("NASPASSWORD=".Length);
                        else if (trimmedLine.StartsWith("NASDOMAIN="))
                            currentNasDomain = trimmedLine.Substring("NASDOMAIN=".Length);
                    }
                }

                // Thêm section cuối cùng nếu có dữ liệu
                if (currentNasPath != null && currentNasUser != null && currentNasPassword != null)
                {
                    nasCredentialsList.Add((currentNasPath, new NetworkCredential(currentNasUser, currentNasPassword, currentNasDomain ?? "")));
                }

                // Nếu không tìm thấy section nào hợp lệ, thêm các server mặc định
                if (nasCredentialsList.Count == 0)
                {
                    foreach (var server in defaultNasServers)
                    {
                        nasCredentialsList.Add((server.NasPath, new NetworkCredential(server.NasUser, server.NasPassword, server.NasDomain)));
                    }
                }

                // Gán nasDirectoryPath cho server đầu tiên trong danh sách
                nasDirectoryPath = nasCredentialsList[0].NasPath;               
                return nasCredentialsList;
            }
            catch (Exception ex)
            {
                LblStatus.ForeColor = System.Drawing.Color.Red;
                LblStatus.Text = $"Lỗi khi đọc/tạo file NASConfig.ini: {ex.Message}";
                AppendToLog($"Lỗi khi đọc/tạo file NASConfig.ini: {ex.Message}", System.Drawing.Color.Red);
                // Trả về danh sách mặc định nếu có lỗi
                foreach (var server in defaultNasServers)
                {
                    nasCredentialsList.Add((server.NasPath, new NetworkCredential(server.NasUser, server.NasPassword, server.NasDomain)));
                }
                nasDirectoryPath = defaultNasServers[0].NasPath; // Gán mặc định
                return nasCredentialsList;
            }
        }

        // Phương thức kiểm tra quyền ghi thư mục
        private bool IsDirectoryWritable(string directoryPath)
        {
            try
            {
                using (FileStream fs = File.Create(Path.Combine(directoryPath, Path.GetRandomFileName()), 1, FileOptions.DeleteOnClose))
                { }
                return true;
            }
            catch
            {
                return false;
            }
        }
        #endregion

        #region ĐỌC LOẠI LỖI TỪ FILE INI
        // Thêm phương thức này vào class MainForm
        private List<string> ReadErrorTypesFromIniFile()
        {
            string filePath = @"C:\IT_CONFIRM\Config\TENLOI.ini";
            List<string> errorTypes = new List<string>();

            // Kiểm tra xem file có tồn tại không
            if (!File.Exists(filePath))
            {
                // Nếu không tồn tại, tạo thư mục nếu cần và tạo file với nội dung mặc định
                string directoryPath = Path.GetDirectoryName(filePath);
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath); // Tạo thư mục nếu chưa có
                }

                // Nội dung mặc định
                string defaultContent = @"[DEFECT_NAME]
B-SPOT
WHITE SPOT
ĐỐM SPIN
ĐỐM PANEL
ĐỐM ĐƯỜNG DỌC";

                try
                {
                    File.WriteAllText(filePath, defaultContent, System.Text.Encoding.UTF8);
                    // LblStatus.ForeColor = System.Drawing.Color.Blue; // Xóa
                    // LblStatus.Text = $"File cấu hình loại lỗi mặc định đã được tạo tại: {filePath}"; // Xóa
                    AppendToLog($"File cấu hình loại lỗi mặc định đã được tạo tại: {filePath}", System.Drawing.Color.Blue);
                }
                catch (Exception ex)
                {
                    // LblStatus.ForeColor = System.Drawing.Color.Red; // Xóa
                    // LblStatus.Text = $"Không thể tạo file cấu hình loại lỗi mặc định: {ex.Message}"; // Xóa
                    AppendToLog($"Không thể tạo file cấu hình loại lỗi mặc định: {ex.Message}", System.Drawing.Color.Red);
                    return new List<string> { "B-SPOT", "WHITE SPOT", "ĐỐM SPIN", "ĐỐM PANEL", "ĐỐM ĐƯỜNG DỌC" };
                }
            }

            // Bây giờ file chắc chắn tồn tại, đọc nội dung
            try
            {
                string[] lines = File.ReadAllLines(filePath);
                bool inErrorTypesSection = false;

                foreach (string line in lines)
                {
                    string trimmedLine = line.Trim();

                    // Bỏ qua dòng trống và comment (bắt đầu bằng ; hoặc #)
                    if (string.IsNullOrEmpty(trimmedLine) || trimmedLine.StartsWith(";") || trimmedLine.StartsWith("#"))
                    {
                        continue;
                    }

                    // Kiểm tra section [ERROR_TYPES]
                    if (trimmedLine.StartsWith("[") && trimmedLine.EndsWith("]"))
                    {
                        inErrorTypesSection = trimmedLine.Equals("[DEFECT_NAME]", StringComparison.OrdinalIgnoreCase);
                        continue;
                    }

                    // Nếu đang trong section [ERROR_TYPES], đọc key=value
                    if (inErrorTypesSection)
                    {
                        int equalsIndex = trimmedLine.IndexOf('=');
                        if (equalsIndex > 0) // Đảm bảo có dấu =
                        {
                            string valueStr = trimmedLine.Substring(equalsIndex + 1).Trim();
                            if (!string.IsNullOrEmpty(valueStr))
                            {
                                errorTypes.Add(valueStr);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // LblStatus.ForeColor = System.Drawing.Color.Red; // Xóa
                // LblStatus.Text = $"Lỗi khi đọc file cấu hình TENLOI.ini: {ex.Message}. Sử dụng giá trị mặc định."; // Xóa
                AppendToLog($"Lỗi khi đọc file cấu hình TENLOI.ini: {ex.Message}. Sử dụng giá trị mặc định.", System.Drawing.Color.Red);
                return new List<string> { "B-SPOT", "WHITE SPOT", "ĐỐM SPIN", "ĐỐM PANEL", "ĐỐM ĐƯỜNG DỌC" };
            }

            // Nếu đọc thành công và không có lỗi, cập nhật status (tuỳ chọn)
            // if (errorTypes.Count > 0)
            // {
            //     // LblStatus.ForeColor = System.Drawing.Color.Green;
            //     // LblStatus.Text = "Cấu hình loại lỗi đã được đọc thành công.";
            // }
            AppendToLog($"Cấu hình tên lỗi đã được đọc thành công. Path: {filePath}", System.Drawing.Color.Green);
            return errorTypes;
        }
        #endregion

        #region TIP HƯỚNG DẪN
        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // Tạo form hướng dẫn mới
            MAP mapForm = new MAP();

            // Hiển thị form dưới dạng Dialog
            mapForm.ShowDialog();
        }
        #endregion
    }
}