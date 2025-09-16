using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using IT_CONFIRM.Properties;

namespace IT_CONFIRM
{
    public partial class MainForm : Form
    {
        #region KHAI BÁO CÁC BIẾN
        private TextBox currentTextBox;
        private ToolTip validationToolTip;
        private string _lastSavedFilePath; // Biến mới để lưu đường dẫn file        
        private ToolTip statusToolTip;// Biến mới cho ToolTip của thông báo trạng thái     
        private Timer rainbowTimer;// Các biến cho hiệu ứng chuyển màu cầu vồng mượt mà
        private bool isRainbowActive = false;
        private Color originalCopyrightColor;
        private double rainbowPhase = 0;       
        private Dictionary<string, Color> originalColors = new Dictionary<string, Color>();// Sử dụng Dictionary để lưu màu gốc của tất cả các nút
        #endregion

        #region FORM KHỞI TẠO UI
        public MainForm()
        {
            InitializeComponent();
            InitializeKeyboardEvents();
            txtSAPN.MaxLength = 300;
            validationToolTip = new ToolTip();//Thông báo yêu cầu nhập dữ liệu
            statusToolTip = new ToolTip();//Gợi ý bấm vào để mở thư mục
            this.lblStatus.Text = "Sẵn sàng nhập dữ liệu...";
            UpdateSavedSAPNCount();

            // Khởi tạo ComboBox với danh sách lỗi
            cboErrorType.Items.AddRange(new string[] { "B-SPOT", "WHITE SPOT", "ĐỐM SPIN", "ĐỐM PANEL", "ĐỐM ĐƯỜNG DỌC", "-" });
            cboErrorType.SelectedIndex = -1; // Mặc định chọn "ĐỐM" =-1 KHÔNG CHỌN GÌ

            // Khởi tạo timer cho hiệu ứng cầu vồng
            this.rainbowTimer = new Timer();
            this.rainbowTimer.Interval = 20; // Cập nhật màu mỗi 20ms để mượt hơn
            this.rainbowTimer.Tick += new EventHandler(this.RainbowTimer_Tick);

            // Gắn sự kiện cho lblCopyright
            lblCopyright.MouseEnter += LblCopyright_MouseEnter;
            lblCopyright.MouseLeave += LblCopyright_MouseLeave;

            // Gán sự kiện Click và thay đổi con trỏ chuột cho lblStatus
            this.lblStatus.Click += new EventHandler(this.lblStatus_Click);
            this.lblStatus.Cursor = Cursors.Hand;

            // Khởi tạo hiệu ứng cho các nút
            InitializeButtonEffects();
        }
        #endregion

        #region HIỆU ỨNG CHO CÁC NÚT BẤM
        private void InitializeButtonEffects()
        {
            Color keyboardBaseColor = System.Drawing.ColorTranslator.FromHtml("#FFF");

            // Xử lý nút SAVE (giữ nguyên màu ban đầu)
            originalColors.Add("btnSave", btnSave.BackColor);
            btnSave.MouseEnter += Button_MouseEnter;
            btnSave.MouseLeave += Button_MouseLeave;
            btnSave.MouseDown += Button_MouseDown;
            btnSave.MouseUp += Button_MouseUp;

            // Xử lý nút RESET (giữ nguyên màu ban đầu)
            originalColors.Add("btnReset", btnReset.BackColor);
            btnReset.MouseEnter += Button_MouseEnter;
            btnReset.MouseLeave += Button_MouseLeave;
            btnReset.MouseDown += Button_MouseDown;
            btnReset.MouseUp += Button_MouseUp;

            // Đặt lại màu cho nút ALL và xử lý riêng biệt
            btnAll.BackColor = System.Drawing.ColorTranslator.FromHtml("#97FFFF");
            originalColors.Add("btnAll", btnAll.BackColor);
            btnAll.MouseEnter += Button_MouseEnter;
            btnAll.MouseLeave += Button_MouseLeave;
            btnAll.MouseDown += Button_MouseDown;
            btnAll.MouseUp += Button_MouseUp;

            // Áp dụng màu nền mới cho các nút trên bàn phím ảo (trừ nút ALL và btnBack)
            Button[] keyboardButtons = { btn0, btn1, btn2, btn3, btn4, btn5, btn6, btn7, btn8, btn9, btnBack };
            foreach (Button btn in keyboardButtons)
            {
                // Chỉ áp dụng màu nền mới cho các nút số
                if (btn.Name != "btnBack")
                {
                    btn.BackColor = keyboardBaseColor;
                }

                // Lưu màu nền hiện tại (đã được đổi) vào từ điển
                originalColors.Add(btn.Name, btn.BackColor);

                btn.MouseEnter += Button_MouseEnter;
                btn.MouseLeave += Button_MouseLeave;
                btn.MouseDown += Button_MouseDown;
                btn.MouseUp += Button_MouseUp;
            }
        }

        // Hiệu ứng khi di chuột vào nút
        private void Button_MouseEnter(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            if (btn != null)
            {
                // Áp dụng hiệu ứng di chuột cho TẤT CẢ các nút
                if (originalColors.ContainsKey(btn.Name))
                {
                    Color originalColor = originalColors[btn.Name];
                    int r = Math.Min(255, originalColor.R + 30);
                    int g = Math.Min(255, originalColor.G + 30);
                    int b = Math.Min(255, originalColor.B + 30);
                    btn.BackColor = Color.FromArgb(r, g, b);
                }
            }
        }

        private void Button_MouseLeave(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            if (btn != null && originalColors.ContainsKey(btn.Name))
            {
                // Khôi phục màu nền ban đầu
                btn.BackColor = originalColors[btn.Name];
            }
        }

        private void Button_MouseDown(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            if (btn != null)
            {
                // Áp dụng hiệu ứng khi nhấn chuột xuống cho TẤT CẢ các nút
                if (originalColors.ContainsKey(btn.Name))
                {
                    Color originalColor = originalColors[btn.Name];
                    int r = Math.Max(0, originalColor.R - 30);
                    int g = Math.Max(0, originalColor.G - 30);
                    int b = Math.Max(0, originalColor.B - 30);
                    btn.BackColor = Color.FromArgb(r, g, b);
                }
            }
        }

        private void Button_MouseUp(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            if (btn != null)
            {
                if (btn.ClientRectangle.Contains(btn.PointToClient(Cursor.Position)))
                {
                    // Nếu nhả chuột trong vùng nút, khôi phục hiệu ứng di chuột vào
                    if (originalColors.ContainsKey(btn.Name))
                    {
                        Color originalColor = originalColors[btn.Name];
                        int r = Math.Min(255, originalColor.R + 30);
                        int g = Math.Min(255, originalColor.G + 30);
                        int b = Math.Min(255, originalColor.B + 30);
                        btn.BackColor = Color.FromArgb(r, g, b);
                    }
                }
                else
                {
                    // Nếu nhả chuột ngoài vùng nút, khôi phục màu ban đầu
                    if (originalColors.ContainsKey(btn.Name))
                    {
                        btn.BackColor = originalColors[btn.Name];
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

            // Gán sự kiện KeyDown cho txtSAPN
            txtSAPN.KeyDown += txtSAPN_KeyDown;

            // Gán sự kiện KeyPress riêng cho txtSx1 (cho phép ALL)
            txtSx1.KeyPress += txtSx1_KeyPress;

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

        // Khi chuột di vào nhãn, kích hoạt hiệu ứng cầu vồng
        private void LblCopyright_MouseEnter(object sender, EventArgs e)
        {
            if (!isRainbowActive)
            {
                isRainbowActive = true;
                originalCopyrightColor = lblCopyright.ForeColor;
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
                lblCopyright.ForeColor = originalCopyrightColor;
            }
        }
        #endregion

        #region THAY ĐỔI MÀU SẮC KHI DI CHUỘT VÀO TÊN TÁC GIẢ
        // Sự kiện Tick của timer, cập nhật màu sắc
        private void RainbowTimer_Tick(object sender, EventArgs e)
        {
            rainbowPhase += 0.05; // Giảm tốc độ thay đổi để màu chuyển từ từ hơn

            Color newColor = CalculateRainbowColor(rainbowPhase);
            lblCopyright.ForeColor = newColor;
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
        // Phương thức xử lý sự kiện Click cho lblStatus
        private void lblStatus_Click(object sender, EventArgs e)
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
        private void txtSAPN_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtSx1.Focus();
                currentTextBox = txtSx1; // Cập nhật currentTextBox để bàn phím ảo tương tác với txtSx1
                e.SuppressKeyPress = true;
            }
        }

        // Xử lý các nút số trên bàn phím ảo
        private void btnNumber_Click(object sender, EventArgs e)
        {
            if (currentTextBox != null)
            {
                Button btn = (Button)sender;
                string buttonText = btn.Text;

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
        private void btnBack_Click(object sender, EventArgs e)
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
        private void btnSave_Click(object sender, EventArgs e)
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

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string appFolderPath = Path.Combine(desktopPath, "IT_CONFIRM");
            string fileName = $"TOA DO_{DateTime.Now:yyyyMMdd}.csv";
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
                lblStatus.ForeColor = System.Drawing.Color.Green;
                lblStatus.Text = $"Lưu thành công! Dữ liệu đã được ghi lại lúc: {timestamp}\nDữ liệu được lưu tại: {filePath}";
                // Thiết lập tooltip cho lblStatus
                statusToolTip.SetToolTip(lblStatus, "Bấm vào đây để mở thư mục lưu file");

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
            }
            catch (IOException)
            {
                // Cập nhật thông báo lỗi cụ thể khi file đang được mở
                lblStatus.ForeColor = System.Drawing.Color.Red;
                lblStatus.Text = "File đang được mở bởi ứng dụng khác hoặc không thể ghi dữ liệu.\nHãy đóng file đang mở trước khi bấm Save";
                // Xóa tooltip khi có lỗi
                statusToolTip.SetToolTip(lblStatus, "");
            }
            catch (Exception ex)
            {
                // Báo lỗi chung nếu có lỗi khác
                lblStatus.ForeColor = System.Drawing.Color.Red;
                lblStatus.Text = $"Đã xảy ra lỗi: {ex.Message}";
                // Xóa tooltip khi có lỗi
                statusToolTip.SetToolTip(lblStatus, "");
            }
        }

        // Xử lý nút RESET
        private void btnReset_Click(object sender, EventArgs e)
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
            lblStatus.ForeColor = System.Drawing.Color.DarkOrange;
            lblStatus.Text = "Đã khởi tạo lại ứng dụng.";
            // Xóa tooltip khi reset ứng dụng
            statusToolTip.SetToolTip(lblStatus, "");
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
                validationToolTip.Show("Chỉ cho phép nhập tối đa 3 số", currentTextBox, 0, currentTextBox.Height, 2000);
                return;
            }

            // Cho phép nhập số và các ký tự điều khiển (như Backspace)
            // Ngăn các ký tự khác (chữ cái, ký tự đặc biệt)
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
                validationToolTip.ToolTipTitle = "Lỗi nhập liệu";
                validationToolTip.Show("Chỉ cho phép nhập số!", currentTextBox, 0, currentTextBox.Height, 2000);
            }
        }

        private void txtSx1_KeyPress(object sender, KeyPressEventArgs e)
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
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string appFolderPath = Path.Combine(desktopPath, "IT_CONFIRM");
            string fileName = $"TOA DO_{DateTime.Now:yyyyMMdd}.csv";
            string filePath = Path.Combine(appFolderPath, fileName);

            int count = 0;
            if (File.Exists(filePath))
            {
                try
                {
                    var lines = File.ReadAllLines(filePath, System.Text.Encoding.UTF8);
                    // Bỏ qua dòng tiêu đề
                    for (int i = 1; i < lines.Length; i++)
                    {
                        var parts = lines[i].Split(',');
                        if (parts.Length > 1 && !string.IsNullOrWhiteSpace(parts[0]))
                        {
                            // Kiểm tra xem có ít nhất một tọa độ không rỗng
                            bool hasCoordinates = false;
                            for (int j = 1; j < parts.Length - 1; j++)
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
                catch (IOException)
                {
                    // Bỏ qua lỗi nếu file đang được mở, không cập nhật bộ đếm
                    return;
                }
                catch (Exception ex)
                {
                    // Xử lý các lỗi khác nếu có
                    lblStatus.ForeColor = System.Drawing.Color.Red;
                    lblStatus.Text = $"Lỗi khi đọc file đếm số lượng: {ex.Message}";
                    return;
                }
            }
            lblSAPNCount.Text = $"Số lượng APN đã lưu: {count}";
        }
        #endregion

        #region TIP HƯỚNG DẪN
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // Tạo form hướng dẫn mới
            MAP mapForm = new MAP();

            // Hiển thị form dưới dạng Dialog
            mapForm.ShowDialog();
        }
        #endregion
    }
}