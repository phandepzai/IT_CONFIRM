using System;
using System.ComponentModel;
using System.Net;
using System.Runtime.InteropServices;



namespace IT_CONFIRM
{
    public class NetworkConnection : IDisposable
    {
        private readonly string _networkName;

        public NetworkConnection(string networkName, NetworkCredential credentials)
        {
            _networkName = networkName;

            var netResource = new NetResource
            {
                Scope = ResourceScope.GlobalNetwork,
                ResourceType = ResourceType.Disk,
                DisplayType = ResourceDisplaytype.Share,
                RemoteName = networkName
            };

            var userName = string.IsNullOrEmpty(credentials.Domain)
                ? credentials.UserName
                : string.Format(@"{0}\\{1}", credentials.Domain, credentials.UserName);

            // Hủy kết nối cũ nếu tồn tại để tránh lỗi 1219/1327
            WNetCancelConnection2(networkName, 0, true);

            var result = WNetAddConnection2(
                netResource,
                credentials.Password,
                userName,
                0);

            if (result != 0)
            {
                string errorMessage = GetWin32ErrorMessage(result);
                throw new Win32Exception(result, $"Lỗi khi kết nối tới {networkName}: {errorMessage} (Mã {result})");
            }
        }

        private string GetWin32ErrorMessage(int errorCode)
        {
            switch (errorCode)
            {
                case 5: return "Truy cập bị từ chối";
                case 67: return "Không tìm thấy tên mạng";
                case 86: return "Mật khẩu mạng được chỉ định không đúng";
                case 1219: return "Không cho phép nhiều kết nối đến một máy chủ từ cùng một người dùng bằng nhiều tên người dùng khác nhau";
                case 1326: return "Đăng nhập thất bại: tên người dùng không tồn tại hoặc mật khẩu sai";
                case 1327: return "Hạn chế tài khoản đang ngăn người dùng này đăng nhập";
                default: return $"Mã lỗi {errorCode}";
            }
        }

        ~NetworkConnection()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            WNetCancelConnection2(_networkName, 0, true);
        }

        [DllImport("mpr.dll")]
        private static extern int WNetAddConnection2(NetResource netResource,
            string password, string username, int flags);

        [DllImport("mpr.dll")]
        private static extern int WNetCancelConnection2(string name, int flags,
            bool force);

        [StructLayout(LayoutKind.Sequential)]
        public class NetResource
        {
            public ResourceScope Scope;
            public ResourceType ResourceType;
            public ResourceDisplaytype DisplayType;
            public int Usage;
            public string LocalName;
            public string RemoteName;
            public string Comment;
            public string Provider;
        }

        public enum ResourceScope
        {
            Connected = 1,
            GlobalNetwork,
            Remembered,
            Recent,
            Context
        }

        public enum ResourceType
        {
            Any = 0,
            Disk = 1,
            Print = 2,
            Reserved = 8
        }

        public enum ResourceDisplaytype
        {
            Generic = 0x0,
            Domain = 0x01,
            Server = 0x02,
            Share = 0x03,
            File = 0x04,
            Group = 0x05,
            Network = 0x06,
            Root = 0x07,
            Shareadmin = 0x08,
            Directory = 0x09,
            Tree = 0x0a,
            Ndscontainer = 0x0b
        }
    }
}