using System.Windows;
using System.Windows.Input;

namespace PureGIS_Geo_QC.WPF
{
    public partial class CustomMessageBox : Window
    {
        // 생성자는 private으로 만들어 외부에서 new로 생성하는 것을 막습니다.
        private CustomMessageBox()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 외부에 노출될 정적 Show 메서드
        /// </summary>
        public static bool? Show(Window owner, string title, string message, bool showCancelButton = false)
        {
            // 우리 커스텀 메시지 박스 창을 생성
            var msgBox = new CustomMessageBox();

            // 전달받은 값으로 UI 세팅
            msgBox.Owner = owner;
            msgBox.TitleText.Text = title;
            msgBox.MessageText.Text = message;

            // 취소 버튼 표시 여부 결정
            if (showCancelButton)
            {
                msgBox.CancelButton.Visibility = Visibility.Visible;
            }

            // 부모 창 위에 모달(Modal) 형태로 띄움
            // .ShowDialog()는 창이 닫힐 때까지 부모 창을 비활성화하고 코드를 멈춥니다.
            return msgBox.ShowDialog();
        }

        // '확인' 버튼 클릭 이벤트
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            // DialogResult를 true로 설정하고 창을 닫습니다.
            this.DialogResult = true;
            this.Close();
        }

        // '취소' 버튼 클릭 이벤트
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            // DialogResult를 false로 설정하고 창을 닫습니다.
            this.DialogResult = false;
            this.Close();
        }

        // 타이틀 바를 드래그하여 창을 이동시키는 이벤트
        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }
        }
    }
}