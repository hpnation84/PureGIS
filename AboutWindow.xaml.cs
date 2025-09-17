using System;
using System.Reflection;
using System.Windows;
using System.Windows.Input;

namespace PureGIS_Geo_QC.WPF
{
    /// <summary>
    /// AboutWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class AboutWindow : Window
    {
        public AboutWindow()
        {
            InitializeComponent();
            LoadVersionInfo();
        }

        /// <summary>
        /// 버전 정보를 동적으로 로드
        /// </summary>
        private void LoadVersionInfo()
        {
            try
            {
                // 어셈블리 정보 가져오기
                Assembly assembly = Assembly.GetExecutingAssembly();
                Version version = assembly.GetName().Version;

                // 버전 정보를 찾아서 업데이트 (XAML에서 TextBlock의 Name 속성이 필요한 경우)
                // 현재는 하드코딩되어 있지만 필요시 동적으로 변경 가능

                // 빌드 날짜 계산 (어셈블리 생성 시간 기반)
                DateTime buildDate = GetBuildDate(assembly);

                // UI 업데이트는 필요시 여기서 수행
                // 예: VersionTextBlock.Text = version.ToString();
                // 예: BuildDateTextBlock.Text = buildDate.ToString("yyyy.MM.dd");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"버전 정보 로드 오류: {ex.Message}");
            }
        }

        /// <summary>
        /// 어셈블리 빌드 날짜 계산
        /// </summary>
        private DateTime GetBuildDate(Assembly assembly)
        {
            try
            {
                // 어셈블리 파일의 생성 시간 반환
                return System.IO.File.GetCreationTime(assembly.Location);
            }
            catch
            {
                return DateTime.Now;
            }
        }

        /// <summary>
        /// 타이틀바 드래그로 창 이동
        /// </summary>
        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                try
                {
                    this.DragMove();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"창 이동 오류: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 닫기 버튼 클릭
        /// </summary>
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// ESC 키로 창 닫기
        /// </summary>
        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.Close();
            }
            base.OnKeyDown(e);
        }
    }
}