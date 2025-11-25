using ITCONFIRM;
using ITCONFIRM.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
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
        private TextBox currentTextBox;
        private readonly ToolTip validationToolTip;
        private string _lastSavedFilePath; // Biến mới để lưu đường dẫn file        
        private readonly ToolTip statusToolTip;// Biến mới cho ToolTip của thông báo trạng thái     
        private readonly Timer rainbowTimer;// Các biến cho hiệu ứng chuyển màu cầu vồng mượt mà
        private bool isRainbowActive = false;
        private Color originalCopyrightColor;
        private double rainbowPhase = 0;
        private readonly Dictionary<string, Color> originalColors = new Dictionary<string, Color>();// Sử dụng Dictionary để lưu màu gốc của tất cả các nút

        // Biến cho NAS       
        private string nasDirectoryPath; // readonly, chỉ gán trong constructor
        #endregion

        #region FORM KHỞI TẠO UI
        public MainForm()
        {
            InitializeComponent();
            string eqpid = ReadEQPIDFromIniFile();
            this.Text = "IT CONFIRM" + (string.IsNullOrEmpty(eqpid) ? "" : "_" + eqpid + "");
            InitializeKeyboardEvents();
            txtSAPN.MaxLength = 89;
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
            var nasCredentials = ReadNASCredentialsFromIniFile(); // Gọi hàm để đọc credentials và tạo file NAS.ini       
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
                    int r = Math.Min(255, originalColor.R + 30);
                    int g = Math.Min(255, originalColor.G + 30);
                    int b = Math.Min(255, originalColor.B + 30);
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
                    int r = Math.Max(0, originalColor.R - 30);
                    int g = Math.Max(0, originalColor.G - 30);
                    int b = Math.Max(0, originalColor.B - 30);
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
            txtSAPN.Click += TextBox_Click;
            txtSx1.Click += TextBox_Click;
            txtSy1.Click += TextBox_Click;
            txtEx1.Click += TextBox_Click;
            txtEy1.Click += TextBox_Click;
            txtSx2.Click += TextBox_Click;
            txtSy2.Click += TextBox_Click;
            txtEx2.Click += TextBox_Click;
            txtEy2.Click += TextBox_Click;
            txtSx3.Click += TextBox_Click;
            txtSy3.Click += TextBox_Click;
            txtEx3.Click += TextBox_Click;
            txtEy3.Click += TextBox_Click;
            txtX1.Click += TextBox_Click;
            txtY1.Click += TextBox_Click;
            txtX2.Click += TextBox_Click;
            txtY2.Click += TextBox_Click;
            txtX3.Click += TextBox_Click;
            txtY3.Click += TextBox_Click;

            // Gán sự kiện KeyDown cho TxtSAPN
            txtSAPN.KeyDown += TxtSAPN_KeyDown;
            // Gán sự kiện KeyPress cho TxtSAPN
            txtSAPN.KeyPress += TxtSAPN_KeyPress;
            // Gán sự kiện TextChanged để cảnh báo sớm
            txtSAPN.TextChanged += TxtSAPN_TextChanged;
            // Gán sự kiện KeyPress riêng cho TxtSx1 (cho phép ALL)
            txtSx1.KeyPress += TxtSx1_KeyPress;

            // Gán sự kiện KeyPress chung cho các ô tọa độ còn lại
            txtSy1.KeyPress += CoordinateTextBox_KeyPress;
            txtEx1.KeyPress += CoordinateTextBox_KeyPress;
            txtEy1.KeyPress += CoordinateTextBox_KeyPress;
            txtSx2.KeyPress += CoordinateTextBox_KeyPress;
            txtSy2.KeyPress += CoordinateTextBox_KeyPress;
            txtEx2.KeyPress += CoordinateTextBox_KeyPress;
            txtEy2.KeyPress += CoordinateTextBox_KeyPress;
            txtSx3.KeyPress += CoordinateTextBox_KeyPress;
            txtSy3.KeyPress += CoordinateTextBox_KeyPress;
            txtEx3.KeyPress += CoordinateTextBox_KeyPress;
            txtEy3.KeyPress += CoordinateTextBox_KeyPress;
            txtX1.KeyPress += CoordinateTextBox_KeyPress;
            txtY1.KeyPress += CoordinateTextBox_KeyPress;
            txtX2.KeyPress += CoordinateTextBox_KeyPress;
            txtY2.KeyPress += CoordinateTextBox_KeyPress;
            txtX3.KeyPress += CoordinateTextBox_KeyPress;
            txtY3.KeyPress += CoordinateTextBox_KeyPress;
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
            int remaining = 89 - txtSAPN.Text.Length;
            if (remaining <= 5 && remaining >= 0)
            {
                validationToolTip.Show($"Còn {remaining} ký tự", txtSAPN, 0, txtSAPN.Height, 1000);
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

        // Phương thức xử lý sự kiện KeyDown của txtSAPN
        private void TxtSAPN_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtSx1.Focus();
                currentTextBox = txtSx1; // Cập nhật currentTextBox để bàn phím ảo tương tác với txtSx1
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
                if (currentTextBox == txtSAPN && currentTextBox.Text.Length >= 89)
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

                // Chỉ cho phép nhập "All" vào txtSx1
                if (buttonText.Equals("All", StringComparison.OrdinalIgnoreCase))
                {
                    if (currentTextBox != txtSx1 || !string.IsNullOrWhiteSpace(currentTextBox.Text))
                    {
                        // Hiển thị tooltip nếu cố gắng nhập "ALL" vào ô khác
                        if (currentTextBox != txtSx1)
                        {
                            validationToolTip.ToolTipTitle = "Lỗi nhập liệu";
                            validationToolTip.Show("Chỉ cho phép nhập ALL vào ô Sx1", currentTextBox, 0, currentTextBox.Height, 2000);
                        }
                        return; // Không làm gì nếu không phải txtSx1 hoặc ô đã có nội dung
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
            if (string.IsNullOrWhiteSpace(txtSAPN.Text))
            {
                return false;
            }

            // Kiểm tra xem ít nhất một trong các ô tọa độ có dữ liệu không
            if (string.IsNullOrWhiteSpace(txtSx1.Text) && string.IsNullOrWhiteSpace(txtSy1.Text) && string.IsNullOrWhiteSpace(txtEx1.Text) && string.IsNullOrWhiteSpace(txtEy1.Text) &&
                string.IsNullOrWhiteSpace(txtSx2.Text) && string.IsNullOrWhiteSpace(txtSy2.Text) && string.IsNullOrWhiteSpace(txtEx2.Text) && string.IsNullOrWhiteSpace(txtEy2.Text) &&
                string.IsNullOrWhiteSpace(txtSx3.Text) && string.IsNullOrWhiteSpace(txtSy3.Text) && string.IsNullOrWhiteSpace(txtEx3.Text) && string.IsNullOrWhiteSpace(txtEy3.Text) &&
                string.IsNullOrWhiteSpace(txtX1.Text) && string.IsNullOrWhiteSpace(txtY1.Text) && string.IsNullOrWhiteSpace(txtX2.Text) && string.IsNullOrWhiteSpace(txtY2.Text) &&
                string.IsNullOrWhiteSpace(txtX3.Text) && string.IsNullOrWhiteSpace(txtY3.Text))
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
            //Kiểm tra xem đã chọn model chưa
            if (!rdoI251.Checked && !rdoI252.Checked)
            {
                validationToolTip.ToolTipIcon = ToolTipIcon.Warning;
                validationToolTip.ToolTipTitle = "Lỗi";
                validationToolTip.Show("Vui lòng chọn model (I251 hoặc I252)!", rdoI251, 0, rdoI251.Height, 5000);
                return;
            }

            // Kiểm tra dữ liệu sAPN trước khi lưu
            if (!IsDataValid())
            {
                validationToolTip.ToolTipIcon = ToolTipIcon.Warning;
                validationToolTip.ToolTipTitle = "Lỗi";
                validationToolTip.Show("Vui lòng nhập sAPN và ít nhất một trong số các ô tọa độ!", txtSAPN, 0, txtSAPN.Height, 5000);
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

                // Lấy model đã chọn từ button được chọn
                string selectedModel = rdoI251.Checked ? "I251" : "I252";
                // Lấy loại lỗi đã chọn từ ComboBox
                string selectedErrorType = cboErrorType.SelectedItem.ToString();

                string csvData = $"{selectedModel},{txtSAPN.Text},{selectedErrorType},{txtSx1.Text},{txtSy1.Text},{txtEx1.Text},{txtEy1.Text}," +
                                 $"{txtSx2.Text},{txtSy2.Text},{txtEx2.Text},{txtEy2.Text}," +
                                 $"{txtSx3.Text},{txtSy3.Text},{txtEx3.Text},{txtEy3.Text}," +
                                 $"{txtX1.Text},{txtY1.Text},{txtX2.Text},{txtY2.Text},{txtX3.Text},{txtY3.Text},{timestamp}";

                File.AppendAllText(filePath, csvData + Environment.NewLine, System.Text.Encoding.UTF8);
                // Cập nhật thông báo thành công
                _lastSavedFilePath = filePath; // Lưu đường dẫn file vào biến toàn cục
                // Cập nhật thông báo thành công
                LblStatus.ForeColor = System.Drawing.Color.Green;
                LblStatus.Text = $"Lưu thành công! \nDữ liệu đã được ghi lại lúc: {timestamp}\nVị trí lưu: {filePath}";
                // Thiết lập tooltip cho LblStatus
                statusToolTip.SetToolTip(LblStatus, "BẤM VÀO ĐÂY ĐỂ MỞ VỊ TRÍ LƯU FILE");

                // Xóa nội dung của tất cả các TextBox sau khi lưu thành công
                txtSAPN.Clear();
                txtSx1.Clear();
                txtSy1.Clear();
                txtEx1.Clear();
                txtEy1.Clear();
                txtSx2.Clear();
                txtSy2.Clear();
                txtEx2.Clear();
                txtEy2.Clear();
                txtSx3.Clear();
                txtSy3.Clear();
                txtEx3.Clear();
                txtEy3.Clear();
                txtX1.Clear();
                txtY1.Clear();
                txtX2.Clear();
                txtY2.Clear();
                txtX3.Clear();
                txtY3.Clear();
                //cboErrorType.SelectedIndex = 0; // Mặc định chọn "ĐỐM" =-1 KHÔNG CHỌN GÌ

                // Đặt focus lại cho ô đầu tiên
                txtSAPN.Focus();
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
                            statusToolTip.SetToolTip(LblStatusNas, "BẤM VÀO ĐÂY ĐỂ MỞ VỊ TRÍ LƯU FILE");
                        }
                        else
                        {
                            LblStatusNas.ForeColor = System.Drawing.Color.Red;
                            LblStatusNas.Text = $"NAS Server: {nasError}";
                            statusToolTip.SetToolTip(LblStatusNas, "Kiểm tra lại NAS.ini hoặc tài khoản NAS");
                        }
                    }));
                });
            }
            catch (IOException)
            {
                // Cập nhật thông báo lỗi cụ thể khi file đang được mở
                LblStatus.ForeColor = System.Drawing.Color.Red;
                LblStatus.Text = "File đang được mở bởi ứng dụng khác hoặc không thể ghi dữ liệu.\nHãy đóng file đang mở trước khi bấm Save";
                // Xóa tooltip khi có lỗi
                statusToolTip.SetToolTip(LblStatus, "");
            }
            catch (Exception ex)
            {
                // Báo lỗi chung nếu có lỗi khác
                LblStatus.ForeColor = System.Drawing.Color.Red;
                LblStatus.Text = $"Đã xảy ra lỗi: {ex.Message}";
                // Xóa tooltip khi có lỗi
                statusToolTip.SetToolTip(LblStatus, "");
            }
        }

        // Xử lý nút RESET
        private void BtnReset_Click(object sender, EventArgs e)
        {
            // Xóa nội dung của tất cả các TextBox
            txtSAPN.Clear();
            txtSx1.Clear();
            txtSy1.Clear();
            txtEx1.Clear();
            txtEy1.Clear();
            txtSx2.Clear();
            txtSy2.Clear();
            txtEx2.Clear();
            txtEy2.Clear();
            txtSx3.Clear();
            txtSy3.Clear();
            txtEx3.Clear();
            txtEy3.Clear();
            txtX1.Clear();
            txtY1.Clear();
            txtX2.Clear();
            txtY2.Clear();
            txtX3.Clear();
            txtY3.Clear();
            cboErrorType.SelectedIndex = -1; // Mặc định chọn "ĐỐM" =0 CHỌN DÒNG ĐẦU TIÊN

            // Cập nhật thông báo
            LblStatus.ForeColor = System.Drawing.Color.DarkOrange;
            LblStatus.Text = "Đã khởi tạo lại ứng dụng.";
            // Xóa tooltip khi reset ứng dụng
            statusToolTip.SetToolTip(LblStatus, "");
            // Reset trạng thái NAS
            LblStatusNas.ForeColor = System.Drawing.Color.DarkOrange;
            LblStatusNas.Text = "";
            statusToolTip.SetToolTip(LblStatusNas, "");
            // Đặt focus lại cho ô đầu tiên
            txtSAPN.Focus();
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

            string eqpid = ReadEQPIDFromIniFile(); // Lấy EQPID từ file ini
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string appFolderPath = Path.Combine(desktopPath, "IT_CONFIRM");
            string fileName = string.IsNullOrEmpty(eqpid) ? $"IT_{dateString}_{shift}.csv" : $"IT_{eqpid}_{dateString}_{shift}.csv";
            string filePath = Path.Combine(appFolderPath, fileName);

            int count = 0;

            try
            {
                // Kiểm tra quyền đọc file
                if (!CanReadFile(filePath))
                {
                    LblStatus.ForeColor = System.Drawing.Color.ForestGreen;                   
                    LblSAPNCount.Text = "Số lượng APN đã lưu: 0";
                    return;
                }

                // Kiểm tra file tồn tại
                if (!File.Exists(filePath))
                {
                    LblSAPNCount.Text = "Số lượng APN đã lưu: 0";
                    return;
                }

                // Đọc file bằng StreamReader để tối ưu hiệu suất
                using (var reader = new StreamReader(filePath, System.Text.Encoding.UTF8))
                {
                    // Bỏ qua dòng tiêu đề
                    string header = reader.ReadLine();
                    if (header == null)
                    {
                        LblSAPNCount.Text = "Số lượng APN đã lưu: 0";
                        return;
                    }

                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        var parts = line.Split(',');
                        // Kiểm tra dòng có đủ cột
                        if (parts.Length >= 22 && !string.IsNullOrWhiteSpace(parts[1]))
                        {
                            // Kiểm tra có ít nhất một tọa độ không rỗng (từ cột Sx1 đến Y3)
                            bool hasCoordinates = false;
                            for (int j = 3; j < parts.Length - 1 && j <= 20; j++)
                            {
                                if (!string.IsNullOrWhiteSpace(parts[j]))
                                {
                                    hasCoordinates = true;
                                    break;
                                }
                            }
                            if (hasCoordinates)
                            {
                                count++;
                            }
                        }
                    }
                }

                // Cập nhật số lượng lên nhãn
                LblSAPNCount.Text = $"Số lượng APN đã lưu: {count}";
            }
            catch (IOException ex)
            {
                // File đang bị khóa hoặc không thể truy cập
                LblStatus.ForeColor = System.Drawing.Color.Red;
                LblStatus.Text = $"Lỗi khi đọc file CSV (có thể đang mở): {ex.Message}";
                LblSAPNCount.Text = "Số lượng APN đã lưu: 0";
            }
            catch (Exception ex)
            {
                // Xử lý các lỗi khác
                LblStatus.ForeColor = System.Drawing.Color.Red;
                LblStatus.Text = $"Lỗi khi đọc file CSV ({filePath}): {ex.Message}";
                LblSAPNCount.Text = "Số lượng APN đã lưu: 0";
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
                    MessageBox.Show("Lỗi khi đọc file MachineParam.ini: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            return "";
        }
        //Đọc thông tin NAS từ file NAS.ini
        private List<(string NasPath, NetworkCredential Credentials)> ReadNASCredentialsFromIniFile()
        {
            string filePath = @"C:\IT_CONFIRM\Config\NAS.ini";
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
                    NasPath = @"\\107.126.41.112\IT_CONFIRM",
                    NasUser = "admin2",
                    NasPassword = "password2",
                    NasDomain = ""
                },
                new
                {
                    NasPath = @"\\107.126.41.113\IT_CONFIRM",
                    NasUser = "admin3",
                    NasPassword = "password3",
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
                    LblStatus.Text = $"Không có quyền ghi vào thư mục {directoryPath} để tạo NAS.ini";
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
                LblStatus.Text = $"Lỗi khi đọc/tạo file NAS.ini: {ex.Message}";
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