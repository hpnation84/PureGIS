using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;
using System.Management;
using System.Timers;
using Newtonsoft.Json;
using System.Collections.Specialized;

namespace PureGIS_Geo_QC.Licensing
{
    public class LicenseManager
    {
        private static readonly string LICENSE_SERVER_URL = "https://www.jindigo.kr/puregis/license_api.php";
        private static readonly int HEARTBEAT_INTERVAL = 120000; // 2분
        private static LicenseManager _instance;
        private static readonly object _lock = new object();

        private Timer _heartbeatTimer;
        private string _sessionToken;
        private string _licenseKey;
        private string _machineId;
        private bool _isLicenseValid = false;

        public static LicenseManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                            _instance = new LicenseManager();
                    }
                }
                return _instance;
            }
        }

        private LicenseManager()
        {
            _machineId = GetMachineId();
            InitializeHeartbeatTimer();
        }

        public bool IsLicenseValid => _isLicenseValid;
        public string SessionToken => _sessionToken;
        public string LicenseKey => _licenseKey;

        /// <summary>
        /// 라이선스 로그인
        /// </summary>
        public async Task<LicenseResult> LoginAsync(string licenseKey)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var parameters = new NameValueCollection
                    {
                        ["action"] = "login",
                        ["license_key"] = licenseKey,
                        ["machine_id"] = _machineId,
                        ["user_name"] = Environment.UserName
                    };

                    var response = await PostAsync(client, LICENSE_SERVER_URL, parameters);
                    var result = JsonConvert.DeserializeObject<LicenseApiResponse>(response);

                    if (result.Success)
                    {
                        _licenseKey = licenseKey;
                        _sessionToken = result.SessionToken;
                        _isLicenseValid = true;

                        StartHeartbeat();

                        return new LicenseResult
                        {
                            IsSuccess = true,
                            CompanyName = result.CompanyName,
                            ExpiryDate = result.ExpiryDate,
                            Message = "라이선스 인증 성공"
                        };
                    }
                    else
                    {
                        return new LicenseResult { IsSuccess = false, Message = result.Message };
                    }
                }
            }
            catch (Exception ex)
            {
                return new LicenseResult { IsSuccess = false, Message = $"네트워크 오류: {ex.Message}" };
            }
        }

        /// <summary>
        /// 라이선스 로그아웃
        /// </summary>
        public async Task<bool> LogoutAsync()
        {
            try
            {
                StopHeartbeat();

                if (string.IsNullOrEmpty(_sessionToken)) return true;

                using (var client = new HttpClient())
                {
                    var parameters = new NameValueCollection
                    {
                        ["action"] = "logout",
                        ["session_token"] = _sessionToken,
                        ["machine_id"] = _machineId
                    };

                    await PostAsync(client, LICENSE_SERVER_URL, parameters);
                }

                _sessionToken = null;
                _licenseKey = null;
                _isLicenseValid = false;

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 세션 유효성 검증
        /// </summary>
        public async Task<bool> ValidateSessionAsync()
        {
            try
            {
                if (string.IsNullOrEmpty(_sessionToken)) return false;

                using (var client = new HttpClient())
                {
                    var parameters = new NameValueCollection
                    {
                        ["action"] = "validate",
                        ["session_token"] = _sessionToken,
                        ["machine_id"] = _machineId
                    };

                    var response = await PostAsync(client, LICENSE_SERVER_URL, parameters);
                    var result = JsonConvert.DeserializeObject<LicenseValidationResponse>(response);

                    _isLicenseValid = result.Success && result.Valid;
                    return _isLicenseValid;
                }
            }
            catch
            {
                _isLicenseValid = false;
                return false;
            }
        }

        private void InitializeHeartbeatTimer()
        {
            _heartbeatTimer = new Timer(HEARTBEAT_INTERVAL);
            _heartbeatTimer.Elapsed += async (sender, e) => await SendHeartbeatAsync();
            _heartbeatTimer.AutoReset = true;
        }

        private void StartHeartbeat()
        {
            _heartbeatTimer?.Start();
        }

        private void StopHeartbeat()
        {
            _heartbeatTimer?.Stop();
        }

        private async Task SendHeartbeatAsync()
        {
            try
            {
                if (string.IsNullOrEmpty(_sessionToken)) return;

                using (var client = new HttpClient())
                {
                    var parameters = new NameValueCollection
                    {
                        ["action"] = "heartbeat",
                        ["session_token"] = _sessionToken,
                        ["machine_id"] = _machineId
                    };

                    var response = await PostAsync(client, LICENSE_SERVER_URL, parameters);
                    var result = JsonConvert.DeserializeObject<LicenseApiResponse>(response);

                    if (!result.Success)
                    {
                        _isLicenseValid = false;
                        StopHeartbeat();

                        // UI에 라이선스 만료 알림
                        System.Windows.Application.Current?.Dispatcher.Invoke(() =>
                        {
                            System.Windows.MessageBox.Show("라이선스 세션이 만료되었습니다. 프로그램을 다시 시작해주세요.",
                                "라이선스 오류", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                        });
                    }
                }
            }
            catch
            {
                // 네트워크 오류는 일시적일 수 있으므로 로그만 남김
                System.Diagnostics.Debug.WriteLine("Heartbeat failed - network error");
            }
        }

        private string GetMachineId()
        {
            try
            {
                // CPU ID + 마더보드 시리얼 + MAC 주소를 조합하여 고유 ID 생성
                var cpuId = GetCpuId();
                var motherboardSerial = GetMotherboardSerial();
                var macAddress = GetMacAddress();

                var combined = $"{cpuId}-{motherboardSerial}-{macAddress}";

                using (var sha256 = SHA256.Create())
                {
                    var hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(combined));
                    return BitConverter.ToString(hash).Replace("-", "").Substring(0, 32);
                }
            }
            catch
            {
                // 하드웨어 정보 수집 실패 시 대안
                return Environment.MachineName + "_" + Environment.UserName;
            }
        }

        private string GetCpuId()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT ProcessorId FROM Win32_Processor"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        return obj["ProcessorId"]?.ToString() ?? "UNKNOWN_CPU";
                    }
                }
            }
            catch { }
            return "UNKNOWN_CPU";
        }

        private string GetMotherboardSerial()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BaseBoard"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        return obj["SerialNumber"]?.ToString() ?? "UNKNOWN_MB";
                    }
                }
            }
            catch { }
            return "UNKNOWN_MB";
        }

        private string GetMacAddress()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT MACAddress FROM Win32_NetworkAdapter WHERE MACAddress IS NOT NULL"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        var mac = obj["MACAddress"]?.ToString();
                        if (!string.IsNullOrEmpty(mac))
                            return mac.Replace(":", "");
                    }
                }
            }
            catch { }
            return "UNKNOWN_MAC";
        }

        private async Task<string> PostAsync(HttpClient client, string url, NameValueCollection parameters)
        {
            var postData = new StringBuilder();
            foreach (string key in parameters.Keys)
            {
                if (postData.Length > 0) postData.Append("&");
                postData.Append($"{Uri.EscapeDataString(key)}={Uri.EscapeDataString(parameters[key])}");
            }

            var content = new StringContent(postData.ToString(), Encoding.UTF8, "application/x-www-form-urlencoded");
            var response = await client.PostAsync(url, content);
            return await response.Content.ReadAsStringAsync();
        }

        public void Dispose()
        {
            StopHeartbeat();
            _heartbeatTimer?.Dispose();
        }
    }

    // API 응답 모델들
    public class LicenseApiResponse
    {
        [JsonProperty("success")]
        public bool Success { get; set; }

        [JsonProperty("message")]
        public string Message { get; set; }

        [JsonProperty("session_token")]
        public string SessionToken { get; set; }

        [JsonProperty("company_name")]
        public string CompanyName { get; set; }

        [JsonProperty("expiry_date")]
        public string ExpiryDate { get; set; }
    }

    public class LicenseValidationResponse
    {
        [JsonProperty("success")]
        public bool Success { get; set; }

        [JsonProperty("valid")]
        public bool Valid { get; set; }

        [JsonProperty("message")]
        public string Message { get; set; }
    }

    public class LicenseResult
    {
        public bool IsSuccess { get; set; }
        public string Message { get; set; }
        public string CompanyName { get; set; }
        public string ExpiryDate { get; set; }
    }
}