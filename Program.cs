using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading; // Thêm using này để sử dụng Mutex

namespace ITCONFIRM
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Tạo Mutex với tên unique (global)
            using (Mutex mutex = new Mutex(true, "ITConfirmSingleInstanceMutex", out bool createdNew))
            {
                if (createdNew)
                {
                    // Nếu Mutex mới được tạo (chưa có instance nào chạy), tiếp tục chạy ứng dụng
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new MainForm());
                }
                else
                {
                    // Nếu Mutex đã tồn tại (instance khác đang chạy), hiển thị thông báo và thoát
                    MessageBox.Show(
                        "Ứng dụng đã đang chạy. Không thể mở thêm phiên bản mới.",
                        "Thông báo",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    // Thoát ứng dụng ngay lập tức
                    Environment.Exit(0);
                }
            }
        }
    }
}