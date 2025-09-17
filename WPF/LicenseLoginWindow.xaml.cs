using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;
using PureGIS_Geo_QC.Licensing;
using LicenseManager = PureGIS_Geo_QC.Licensing.LicenseManager;

namespace PureGIS_Geo_QC.WPF
{
    public partial class LicenseLoginWindow : Window
    {
        public bool IsAuthenticated { get; private set; } = false;
        public bool IsTrialMode { get; private set; } = false;
        public string CompanyName { get; private set; } = "";
        public string ExpiryDate { get; private set; } = "";

        public LicenseLoginWindow()
        {
            InitializeComponent();

            // Enter 키로 인증
            LicenseKeyTextBox.KeyDown += (s, e) =>
            {
                if (e.Key == Key.Enter)
                    AuthenticateButton_Click(s, null);
            };
        }

        private async void AuthenticateButton_Click(object sender, RoutedEventArgs e)
        {
            string licenseKey = LicenseKeyTextBox.Text.Trim();

            if (string.IsNullOrEmpty(licenseKey))
            {
                ShowStatus("라이선스 키를 입력하세요.", true);
                return;
            }

            // UI 비활성화
            SetUIEnabled(false);
            ShowStatus("인증 중...", false);

            try
            {
                var result = await LicenseManager.Instance.LoginAsync(licenseKey);

                if (result.IsSuccess)
                {
                    IsAuthenticated = true;
                    CompanyName = result.CompanyName;
                    ExpiryDate = result.ExpiryDate;

                    ShowStatus($"인증 성공! ({result.CompanyName})", false);

                    // 잠시 대기 후 창 닫기
                    await System.Threading.Tasks.Task.Delay(1000);
                    this.DialogResult = true;
                    this.Close();
                }
                else
                {
                    ShowStatus($"인증 실패: {result.Message}", true);
                }
            }
            catch (Exception ex)
            {
                ShowStatus($"오류: {ex.Message}", true);
            }
            finally
            {
                SetUIEnabled(true);
            }
        }

        private void TrialButton_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "체험판으로 시작하시겠습니까?\n\n" +
                "체험판 제한사항:\n" +
                "- 최대 5개 파일까지만 검사 가능\n" +
                "- 보고서에 체험판 워터마크 표시\n" +
                "- 일부 고급 기능 제한\n\n" +
                "정식 라이선스는 jindigo.kr에서 구매하실 수 있습니다.",
                "체험판 시작",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                IsTrialMode = true;
                this.DialogResult = true;
                this.Close();
            }
        }

        private void ExitButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void ShowStatus(string message, bool isError)
        {
            StatusMessage.Text = message;
            StatusMessage.Foreground = isError ?
                new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 107, 107)) :
                new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(72, 187, 120));
        }

        private void SetUIEnabled(bool enabled)
        {
            AuthenticateButton.IsEnabled = enabled;
            TrialButton.IsEnabled = enabled;
            LicenseKeyTextBox.IsEnabled = enabled;

            AuthenticateButton.Content = enabled ? "인증하기" : "인증 중...";
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            // 인증되지 않고 체험판도 아니면 프로그램 종료
            if (!IsAuthenticated && !IsTrialMode)
            {
                Environment.Exit(0);
            }

            base.OnClosing(e);
        }
    }
}