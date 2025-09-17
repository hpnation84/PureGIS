using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DotSpatial.Data;
using Microsoft.Win32;
using PureGIS_Geo_QC.Models;
using PureGIS_Geo_QC;
using PureGIS_Geo_QC.Managers;
using PureGIS_Geo_QC.Exports;
using PureGIS_Geo_QC.Exports.Models;

// 이름 충돌을 피하기 위한 using 별칭(alias) 사용
using ColumnDefinition = PureGIS_Geo_QC.Models.ColumnDefinition;
using TableDefinition = PureGIS_Geo_QC.Models.TableDefinition;
using PdfSharpCore.Fonts;

namespace PureGIS_Geo_QC.WPF
{    
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<TableDefinition> standardTables = new List<TableDefinition>();
        // DataTable 대신 Shapefile 객체를 직접 저장합니다.
        private Shapefile loadedShapefile = null;
        private ProjectDefinition currentProject = null;
        private TableDefinition currentSelectedTable = null;

        public MainWindow()
        {
            // =======================================================
            // ✨ PdfSharpCore 폰트 리졸버를 전역으로 설정
            // =======================================================
            GlobalFontSettings.FontResolver = new FontResolver();
            InitializeComponent();

        }
        // =======================================================
        // ✨ PdfSharpCore 폰트 리졸버 구현을 위한 내부 클래스 추가
        // =======================================================
        // MainWindow.xaml.cs 파일 내부에 있는 클래스입니다.

        public class FontResolver : IFontResolver
        {
            // =======================================================
            // ✨ 이 속성을 추가하여 오류를 해결합니다.
            // IFontResolver 인터페이스는 기본 폰트 이름을 지정하는 속성을 요구합니다.
            // =======================================================
            public string DefaultFontName => "Malgun Gothic";

            public byte[] GetFont(string faceName)
            {
                // 'faceName'에 따라 다른 폰트 파일을 반환할 수 있습니다.
                // 여기서는 'Malgun Gothic' 폰트 파일 경로를 사용합니다.
                // 대부분의 Windows 시스템에 해당 경로에 폰트가 있습니다.

                // 폰트 파일 경로는 대소문자를 구분하지 않도록 수정합니다.
                string fontPath = "C:/Windows/Fonts/malgun.ttf";
                if (faceName.Contains("Bold")) // 굵은 글꼴 요청 시
                {
                    fontPath = "C:/Windows/Fonts/malgunbd.ttf";
                }

                return File.ReadAllBytes(fontPath);
            }

            public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
            {
                // 폰트 패밀리 이름으로 폰트 파일을 매핑합니다.
                if (familyName.Equals("Malgun Gothic", StringComparison.OrdinalIgnoreCase))
                {
                    // PdfSharpCore에게 이 폰트 패밀리의 이름을 알려줍니다.
                    // GetFont 메서드에서 이 이름을 사용할 수 있습니다.
                    if (isBold)
                    {
                        // 굵은 글꼴일 경우 "Malgun Gothic Bold"로 구분
                        return new FontResolverInfo("Malgun Gothic Bold");
                    }

                    return new FontResolverInfo("Malgun Gothic");
                }

                // 지정된 폰트가 없으면 기본값 반환
                return null;
            }
        }

        public ProjectDefinition CurrentProject
        {
            get => currentProject;
            private set
            {
                currentProject = value;
                UpdateProjectUI();
            }
        }

        /// <summary>
        /// 프로젝트 변경 시 UI 업데이트
        /// </summary>
        private void UpdateProjectUI()
        {
            try
            {
                if (CurrentProject == null)
                {
                    // 프로젝트가 없을 때
                    ProjectTreeView.ItemsSource = null;
                    this.Title = "PureGIS Geo-QC";
                    return;
                }

                // 프로젝트가 로드되었을 때
                this.Title = $"PureGIS Geo-QC - {CurrentProject.ProjectName}";

                // TreeView에 카테고리 구조로 바인딩
                ProjectTreeView.ItemsSource = CurrentProject.Categories;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UpdateProjectUI 오류: {ex.Message}");
            }
        }

        // 2. TreeView 선택 변경 이벤트 핸들러 추가 (XAML에서 참조하고 있지만 구현되지 않음)
        private void ProjectTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            try
            {
                if (e.NewValue is TableDefinition selectedTable)
                {
                    currentSelectedTable = selectedTable;
                    StandardGrid.ItemsSource = selectedTable.Columns;
                    ShowTableInfo(selectedTable);
                }
                else
                {
                    currentSelectedTable = null;
                    StandardGrid.ItemsSource = null;
                    HideTableInfo();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ProjectTreeView_SelectedItemChanged 오류: {ex.Message}");
                CustomMessageBox.Show(this, "오류", $"테이블 선택 중 오류가 발생했습니다: {ex.Message}");
            }
        }

        // =================== 탭 1 로직 ===================
        /// <summary>
        /// 새 테이블 만들기
        /// </summary>
        private void NewTableButton_Click(object sender, RoutedEventArgs e)
        {
            if (CurrentProject == null)
            {
                CustomMessageBox.Show(this, "오류", "프로젝트를 먼저 생성하거나 불러오세요.");
                return;
            }

            // 카테고리 선택을 위한 간단한 방법 (첫 번째 카테고리에 추가)
            var firstCategory = CurrentProject.Categories.FirstOrDefault();
            if (firstCategory == null)
            {
                CustomMessageBox.Show(this, "오류", "사용 가능한 카테고리가 없습니다.");
                return;
            }

            var newTable = new TableDefinition
            {
                TableId = "NEW_TABLE_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"),
                TableName = "새 테이블"
            };

            firstCategory.Tables.Add(newTable);
            UpdateTableList();

            // TreeView에서 새로 만든 테이블을 선택하는 로직 (복잡함 - 일단 생략)
            CustomMessageBox.Show(this, "완료", "새 테이블이 생성되었습니다. 테이블 정보를 편집하고 Ctrl+V로 컬럼을 추가하세요.");
        }

        /// <summary>
        /// 선택 테이블 삭제
        /// </summary>
        private void DeleteTableButton_Click(object sender, RoutedEventArgs e)
        {
            if (currentSelectedTable == null)
            {
                CustomMessageBox.Show(this, "알림", "삭제할 테이블을 먼저 선택하세요.");
                return;
            }

            var result = CustomMessageBox.Show(this, "테이블 삭제",
                $"'{currentSelectedTable.TableName}' 테이블을 삭제하시겠습니까?", true);

            if (result == true)
            {
                // 해당 테이블이 속한 카테고리에서 제거
                foreach (var category in CurrentProject.Categories)
                {
                    if (category.Tables.Contains(currentSelectedTable))
                    {
                        category.Tables.Remove(currentSelectedTable);
                        break;
                    }
                }

                currentSelectedTable = null;
                UpdateTableList();
                HideTableInfo();
                CustomMessageBox.Show(this, "완료", "테이블이 삭제되었습니다.");
            }
        }               

        /// <summary>
        /// 테이블 목록 업데이트 (Null 안전 버전)
        /// </summary>
        private void UpdateTableList()
        {
            try
            {
                if (ProjectTreeView != null && CurrentProject != null)
                {
                    ProjectTreeView.ItemsSource = null;
                    ProjectTreeView.ItemsSource = CurrentProject.Categories;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UpdateTableList 오류: {ex.Message}");
            }
        }


        /// <summary>
        /// 기존 UpdateTreeView 호환성을 위한 래퍼
        /// </summary>
        private void UpdateTreeView()
        {
            UpdateTableList();
        }

        /// <summary>
        /// Ctrl+V로 컬럼 데이터 붙여넣기
        /// </summary>
        private void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                if (MainTabControl.SelectedIndex == 0) // 기준 정의 탭
                {
                    PasteColumnsToCurrentTable();
                    e.Handled = true;
                }
            }
        }

        /// <summary>
        /// 현재 선택된 테이블에 컬럼 붙여넣기
        /// </summary>
        private void PasteColumnsToCurrentTable()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("=== PasteColumnsToCurrentTable 시작 ===");

                // 1. 선택된 테이블 확인
                System.Diagnostics.Debug.WriteLine($"currentSelectedTable null 체크: {currentSelectedTable == null}");
                if (currentSelectedTable == null)
                {
                    CustomMessageBox.Show(this, "알림", "컬럼을 추가할 테이블을 먼저 선택하세요.");
                    return;
                }

                System.Diagnostics.Debug.WriteLine($"선택된 테이블: {currentSelectedTable.TableName}");
                System.Diagnostics.Debug.WriteLine($"기존 컬럼 수: {currentSelectedTable.Columns?.Count ?? 0}");

                // 2. 클립보드 텍스트 가져오기
                string clipboardText = null;
                try
                {
                    System.Diagnostics.Debug.WriteLine("클립보드 텍스트 가져오는 중...");
                    clipboardText = Clipboard.GetText();
                    System.Diagnostics.Debug.WriteLine($"클립보드 텍스트 길이: {clipboardText?.Length ?? 0}");
                    if (!string.IsNullOrEmpty(clipboardText))
                    {
                        System.Diagnostics.Debug.WriteLine($"클립보드 내용 일부: {clipboardText.Substring(0, Math.Min(100, clipboardText.Length))}");
                    }
                }
                catch (Exception clipEx)
                {
                    System.Diagnostics.Debug.WriteLine($"클립보드 오류: {clipEx.Message}");
                    CustomMessageBox.Show(this, "오류", $"클립보드에서 텍스트를 가져올 수 없습니다: {clipEx.Message}");
                    return;
                }

                if (string.IsNullOrWhiteSpace(clipboardText))
                {
                    CustomMessageBox.Show(this, "알림", "클립보드가 비어있습니다.");
                    return;
                }

                // 3. 컬럼 데이터 파싱
                System.Diagnostics.Debug.WriteLine("컬럼 데이터 파싱 중...");
                var newColumns = ParseColumnsFromClipboard(clipboardText);
                if (newColumns == null)
                {
                    System.Diagnostics.Debug.WriteLine("파싱 결과가 null입니다.");
                    CustomMessageBox.Show(this, "오류", "컬럼 데이터 파싱에 실패했습니다.");
                    return;
                }

                System.Diagnostics.Debug.WriteLine($"파싱된 컬럼 수: {newColumns.Count}");

                if (newColumns.Count > 0)
                {
                    // 4. 컬럼 추가
                    System.Diagnostics.Debug.WriteLine("컬럼 추가 중...");
                    if (currentSelectedTable.Columns == null)
                    {
                        System.Diagnostics.Debug.WriteLine("Columns 리스트가 null이므로 새로 생성합니다.");
                        currentSelectedTable.Columns = new List<ColumnDefinition>();
                    }

                    currentSelectedTable.Columns.AddRange(newColumns);
                    System.Diagnostics.Debug.WriteLine($"컬럼 추가 완료. 총 컬럼 수: {currentSelectedTable.Columns.Count}");

                    // 5. UI 업데이트
                    System.Diagnostics.Debug.WriteLine("UI 업데이트 중...");
                    try
                    {
                        RefreshSelectedTableGrid();
                        System.Diagnostics.Debug.WriteLine("DataGrid 새로고침 완료");
                    }
                    catch (Exception gridEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"DataGrid 새로고침 오류: {gridEx.Message}");
                    }

                    try
                    {
                        UpdateTableList(); // ListBox 업데이트
                        System.Diagnostics.Debug.WriteLine("ListBox 업데이트 완료");
                    }
                    catch (Exception listEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"ListBox 업데이트 오류: {listEx.Message}");
                    }

                    // 6. 헤더 업데이트 (이 부분에서 오류 발생 가능성 높음)
                    System.Diagnostics.Debug.WriteLine("헤더 업데이트 중...");
                    try
                    {
                        ShowTableInfo(currentSelectedTable);
                        System.Diagnostics.Debug.WriteLine("헤더 업데이트 완료");
                    }
                    catch (Exception headerEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"헤더 업데이트 오류: {headerEx.Message}");
                        System.Diagnostics.Debug.WriteLine($"헤더 업데이트 스택트레이스: {headerEx.StackTrace}");
                        // 헤더 업데이트 실패해도 계속 진행
                    }

                    // 7. 성공 메시지
                    System.Diagnostics.Debug.WriteLine("성공 메시지 표시 중...");
                    string tableName = currentSelectedTable?.TableName ?? "테이블";
                    CustomMessageBox.Show(this, "완료", $"{newColumns.Count}개의 컬럼이 '{tableName}' 에 추가되었습니다.");
                }
                else
                {
                    CustomMessageBox.Show(this, "오류", "올바른 컬럼 데이터를 찾을 수 없습니다.\n\n형식: 컬럼ID [Tab] 컬럼명 [Tab] 타입 [Tab] 길이");
                }

                System.Diagnostics.Debug.WriteLine("=== PasteColumnsToCurrentTable 완료 ===");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"=== PasteColumnsToCurrentTable 전체 오류 ===");
                System.Diagnostics.Debug.WriteLine($"오류 메시지: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"오류 위치: {ex.TargetSite}");
                System.Diagnostics.Debug.WriteLine($"스택 트레이스: {ex.StackTrace}");
                System.Diagnostics.Debug.WriteLine($"=== 오류 정보 끝 ===");

                CustomMessageBox.Show(this, "파싱 오류", $"컬럼 데이터 붙여넣기 중 오류가 발생했습니다:\n\n{ex.Message}\n\n자세한 정보는 디버그 출력을 확인하세요.");
            }
        }

        /// <summary>
        /// 클립보드에서 컬럼 데이터만 파싱 (Null 안전 버전)
        /// </summary>
        private List<ColumnDefinition> ParseColumnsFromClipboard(string clipboardText)
        {
            var columns = new List<ColumnDefinition>();

            try
            {
                if (string.IsNullOrWhiteSpace(clipboardText))
                {
                    return columns;
                }

                var lines = clipboardText.Trim().Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    var cols = line.Split('\t');
                    if (cols.Length < 2) continue; // 최소 2개 컬럼은 있어야 함

                    columns.Add(new ColumnDefinition
                    {
                        ColumnId = GetSafeArrayValue(cols, 0, "COL_" + DateTime.Now.Ticks.ToString().Substring(10)),
                        ColumnName = GetSafeArrayValue(cols, 1, "컬럼_" + columns.Count),
                        Type = GetSafeArrayValue(cols, 2, "VARCHAR2"),
                        Length = GetSafeArrayValue(cols, 3, "50"),
                        IsNotNull = false, // 기본값
                        KeyType = "" // 기본값
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ParseColumnsFromClipboard 오류: {ex.Message}");
                // 빈 리스트 반환
            }

            return columns;
        }

        /// <summary>
        /// 배열에서 안전하게 값을 가져오는 헬퍼 함수
        /// </summary>
        private string GetSafeArrayValue(string[] array, int index, string defaultValue)
        {
            if (array != null && index >= 0 && index < array.Length)
            {
                return string.IsNullOrWhiteSpace(array[index]) ? defaultValue : array[index].Trim();
            }
            return defaultValue;
        }

        /// <summary>
        /// 테이블 정보 패널 표시 (추가 디버깅 버전)
        /// </summary>
        private void ShowTableInfo(TableDefinition table)
        {
            System.Diagnostics.Debug.WriteLine("=== ShowTableInfo 시작 ===");

            // table이 null인지 먼저 체크
            if (table == null)
            {
                System.Diagnostics.Debug.WriteLine("table이 null입니다.");
                HideTableInfo();
                return;
            }

            System.Diagnostics.Debug.WriteLine($"테이블 정보: ID={table.TableId}, Name={table.TableName}");
            System.Diagnostics.Debug.WriteLine($"컬럼 수: {table.Columns?.Count ?? 0}");

            try
            {
                // UI 컨트롤 null 체크
                System.Diagnostics.Debug.WriteLine($"TableInfoPanel null 체크: {TableInfoPanel == null}");
                System.Diagnostics.Debug.WriteLine($"TableIdTextBox null 체크: {TableIdTextBox == null}");
                System.Diagnostics.Debug.WriteLine($"TableNameTextBox null 체크: {TableNameTextBox == null}");
                System.Diagnostics.Debug.WriteLine($"SelectedTableHeader null 체크: {SelectedTableHeader == null}");

                // TableInfoPanel 먼저 표시 (내부 컨트롤들이 초기화되도록)
                if (TableInfoPanel != null)
                {
                    TableInfoPanel.Visibility = Visibility.Visible;
                    System.Diagnostics.Debug.WriteLine("TableInfoPanel 표시 완료");
                }

                // 텍스트박스들 업데이트
                if (TableIdTextBox != null)
                {
                    TableIdTextBox.Text = table.TableId ?? "";
                    System.Diagnostics.Debug.WriteLine("TableIdTextBox 업데이트 완료");
                }

                if (TableNameTextBox != null)
                {
                    TableNameTextBox.Text = table.TableName ?? "";
                    System.Diagnostics.Debug.WriteLine("TableNameTextBox 업데이트 완료");
                }

                // SelectedTableHeader 업데이트 (패널이 표시된 후에)
                if (SelectedTableHeader != null)
                {
                    int columnCount = table.Columns?.Count ?? 0;
                    string headerText = $"📋 {table.TableName ?? "이름없음"} ({columnCount}개 컬럼)";
                    SelectedTableHeader.Text = headerText;
                    System.Diagnostics.Debug.WriteLine($"SelectedTableHeader 업데이트 완료: {headerText}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("SelectedTableHeader가 null이므로 건너뜁니다.");
                }

                System.Diagnostics.Debug.WriteLine("=== ShowTableInfo 완료 ===");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"=== ShowTableInfo 오류 ===");
                System.Diagnostics.Debug.WriteLine($"오류 메시지: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"스택 트레이스: {ex.StackTrace}");

                // 오류 발생 시 패널만 표시
                if (TableInfoPanel != null)
                {
                    TableInfoPanel.Visibility = Visibility.Visible;
                }
            }
        }

        /// <summary>
        /// 테이블 정보 패널 숨김 (Null 안전 버전)
        /// </summary>
        private void HideTableInfo()
        {
            try
            {
                if (TableInfoPanel != null)
                {
                    TableInfoPanel.Visibility = Visibility.Collapsed;
                }

                if (SelectedTableHeader != null)
                {
                    SelectedTableHeader.Text = "테이블을 선택하세요";
                }

                // 텍스트박스 클리어
                if (TableIdTextBox != null)
                {
                    TableIdTextBox.Text = "";
                }

                if (TableNameTextBox != null)
                {
                    TableNameTextBox.Text = "";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"HideTableInfo 오류: {ex.Message}");
            }
        }

        /// <summary>
        /// 테이블 정보 저장 (Null 안전 버전)
        /// </summary>
        private void SaveTableInfoButton_Click(object sender, RoutedEventArgs e)
        {
            if (currentSelectedTable == null)
            {
                CustomMessageBox.Show(this, "오류", "선택된 테이블이 없습니다.");
                return;
            }

            if (CurrentProject == null)
            {
                CustomMessageBox.Show(this, "오류", "프로젝트가 로드되지 않았습니다.");
                return;
            }

            try
            {
                string newTableId = TableIdTextBox?.Text?.Trim() ?? "";
                string newTableName = TableNameTextBox?.Text?.Trim() ?? "";

                if (string.IsNullOrEmpty(newTableId) || string.IsNullOrEmpty(newTableName))
                {
                    CustomMessageBox.Show(this, "오류", "테이블 ID와 테이블명을 모두 입력하세요.");
                    return;
                }

                // 🔥 수정: CurrentProject.Categories에서 중복 ID 체크
                bool isDuplicate = false;
                foreach (var category in CurrentProject.Categories)
                {
                    if (category.Tables.Any(t => t != currentSelectedTable && t.TableId == newTableId))
                    {
                        isDuplicate = true;
                        break;
                    }
                }

                if (isDuplicate)
                {
                    CustomMessageBox.Show(this, "오류", "동일한 테이블 ID가 이미 존재합니다.");
                    return;
                }

                // 테이블 정보 업데이트
                currentSelectedTable.TableId = newTableId;
                currentSelectedTable.TableName = newTableName;

                UpdateTableList();
                ShowTableInfo(currentSelectedTable); // 헤더 업데이트
                CustomMessageBox.Show(this, "완료", "테이블 정보가 저장되었습니다.");
            }
            catch (Exception ex)
            {
                CustomMessageBox.Show(this, "오류", $"테이블 정보 저장 중 오류가 발생했습니다: {ex.Message}");
            }
        }

        // <summary>
        /// 선택된 테이블의 DataGrid 새로고침 (Null 안전 버전)
        /// </summary>
        private void RefreshSelectedTableGrid()
        {
            try
            {
                if (currentSelectedTable != null && StandardGrid != null)
                {
                    StandardGrid.ItemsSource = null;
                    StandardGrid.ItemsSource = currentSelectedTable.Columns;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"RefreshSelectedTableGrid 오류: {ex.Message}");
            }
        }

        // =================== 탭 2 로직 ===================
        #region Tab2 Methods
        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog { Filter = "Shapefiles (*.shp)|*.shp" };
            if (openFileDialog.ShowDialog() != true) return;
            FileNameText.Text = openFileDialog.SafeFileName;
            try
            {
                // Shapefile.OpenFile() 메서드 사용
                if (Shapefile.OpenFile(openFileDialog.FileName) is Shapefile shapefile)
                {
                    this.loadedShapefile = shapefile;
                    var columnInfoList = new List<FileColumnInfo>();

                    // DataTable의 각 컬럼을 순회
                    foreach (DataColumn col in shapefile.DataTable.Columns)
                    {
                        // GetDbfFieldInfo 헬퍼 메서드를 사용하여 상세 정보 추출
                        var (typeName, precision, scale) = GetDbfFieldInfo(shapefile, col.ColumnName);
                        columnInfoList.Add(new FileColumnInfo
                        {
                            ColumnName = col.ColumnName,
                            DataType = new TypeInfo { Name = typeName },
                            MaxLength = scale > 0 ? $"{precision},{scale}" : precision.ToString()
                        });
                    }
                    LoadedFileGrid.ItemsSource = columnInfoList;
                }
                else
                {
                    CustomMessageBox.Show(this, "파일 형식 오류", "선택한 파일이 Shapefile이 아닙니다.");
                    this.loadedShapefile = null;
                    LoadedFileGrid.ItemsSource = null;
                }
            }
            catch (Exception ex)
            {
                CustomMessageBox.Show(this, "파일 열기 오류", $"파일을 여는 중 오류가 발생했습니다: {ex.Message}");
                this.loadedShapefile = null;
                LoadedFileGrid.ItemsSource = null;
            }
        }

        private void ValidateButton_Click(object sender, RoutedEventArgs e)
        {
            if (loadedShapefile == null)
            {
                CustomMessageBox.Show(this, "오류", "'파일 열기' 버튼으로 검사할 파일을 먼저 불러와주세요.");
                return;
            }

            if (CurrentProject == null)
            {
                CustomMessageBox.Show(this, "오류", "프로젝트를 먼저 생성하거나 불러오세요.");
                return;
            }

            string fileId = System.IO.Path.GetFileNameWithoutExtension(loadedShapefile.Filename);

            // 🔥 수정: standardTables 대신 CurrentProject.Categories에서 테이블 검색
            TableDefinition standardTableToValidate = null;

            foreach (var category in CurrentProject.Categories)
            {
                standardTableToValidate = category.Tables
                    .FirstOrDefault(t => t.TableId.Equals(fileId, StringComparison.OrdinalIgnoreCase));
                if (standardTableToValidate != null)
                    break;
            }

            if (standardTableToValidate == null)
            {
                CustomMessageBox.Show(this, "기준 없음",
                    $"'{fileId}' 파일에 해당하는 기준 테이블(TableId)이 '1. 기준 정의'에 없습니다.\n\n" +
                    $"현재 로드된 프로젝트: {CurrentProject?.ProjectName ?? "없음"}\n" +
                    $"등록된 카테고리 수: {CurrentProject?.Categories?.Count ?? 0}");
                return;
            }

            ValidateFile(loadedShapefile, standardTableToValidate);
            MainTabControl.SelectedIndex = 2;
        }
        #endregion

        // =================== 탭 3 로직 (검사 실행) ===================        
        #region Validation Methods
        private void ValidateFile(Shapefile shapefile, TableDefinition standardTable)
        {
            var results = new List<ColumnValidationResult>();
            try
            {
                foreach (var stdCol in standardTable.Columns)
                {
                    var resultRow = new ColumnValidationResult
                    {
                        // ✅ 추가: 컬럼ID 저장
                        Std_ColumnId = stdCol.ColumnId,

                        Std_ColumnName = stdCol.ColumnName,
                        Std_Type = stdCol.Type,
                        Std_Length = stdCol.Length,
                    };

                    // ❌ 기존: stdCol.ColumnName으로 찾았음
                    // ✅ 수정: stdCol.ColumnId로 변경
                    if (!shapefile.DataTable.Columns.Contains(stdCol.ColumnId))
                    {
                        resultRow.Status = "오류";
                        // ✅ 추가: 찾은 필드명과 존재 여부 설정
                        resultRow.Found_FieldName = "없음";
                        resultRow.IsFieldFound = false;

                        resultRow.Cur_Type = "없음";
                        resultRow.Cur_Length = "없음";
                        resultRow.IsTypeCorrect = false;
                        resultRow.IsLengthCorrect = false;
                        results.Add(resultRow);
                        continue;
                    }
                    // ❌ 기존: stdCol.ColumnName으로 필드 정보 가져왔음
                    // ✅ 수정: stdCol.ColumnId로 변경
                    var (curTypeName, curPrecision, curScale) = GetDbfFieldInfo(shapefile, stdCol.ColumnId);

                    // ✅ 추가: 찾은 필드명과 존재 여부 설정
                    resultRow.Found_FieldName = stdCol.ColumnId;
                    resultRow.IsFieldFound = true;

                    resultRow.Cur_Type = curTypeName;
                    resultRow.Cur_Length = curScale > 0 ? $"{curPrecision},{curScale}" : curPrecision.ToString();

                    if (stdCol.Type.Equals("VARCHAR2", StringComparison.OrdinalIgnoreCase))
                    {
                        resultRow.IsTypeCorrect = curTypeName.Equals("Character", StringComparison.OrdinalIgnoreCase);
                    }
                    else if (stdCol.Type.Equals("NUMBER", StringComparison.OrdinalIgnoreCase))
                    {
                        resultRow.IsTypeCorrect = curTypeName.Equals("Numeric", StringComparison.OrdinalIgnoreCase);
                    }
                    else
                    {
                        resultRow.IsTypeCorrect = stdCol.Type.Equals(curTypeName, StringComparison.OrdinalIgnoreCase);
                    }

                    if (resultRow.IsTypeCorrect)
                    {
                        var (stdPrecision, stdScale) = ParseStandardLength(stdCol.Length);
                        if (stdCol.Type.Equals("VARCHAR2", StringComparison.OrdinalIgnoreCase))
                        {
                            resultRow.IsLengthCorrect = (stdPrecision == curPrecision);
                        }
                        else if (stdCol.Type.Equals("NUMBER", StringComparison.OrdinalIgnoreCase))
                        {
                            resultRow.IsLengthCorrect = (stdPrecision == curPrecision && stdScale == curScale);
                        }
                        else
                        {
                            resultRow.IsLengthCorrect = true;
                        }
                    }
                    else
                    {
                        resultRow.IsLengthCorrect = false;
                    }

                    resultRow.Status = (resultRow.IsTypeCorrect && resultRow.IsLengthCorrect) ? "정상" : "오류";
                    results.Add(resultRow);
                }

                ResultGrid.ItemsSource = results;
                int errorCount = results.Count(r => r.Status == "오류");
                CustomMessageBox.Show(this, "검증 완료", $"검증 완료: 총 {results.Count}개 중 정상: {results.Count - errorCount}개, 오류: {errorCount}개");
            }
            catch (Exception ex)
            {
                CustomMessageBox.Show(this, "검사 오류", $"파일 검사 중 오류가 발생했습니다: {ex.Message}");
            }
        }

        /// <summary>
        /// *** 최종 수정된 헬퍼 메서드 (DotSpatial v2.0 호환) ***
        /// DataTable의 DataColumn을 DotSpatial.Data.Field로 캐스팅하여 상세 정보를 추출합니다.
        /// </summary>
        private (string TypeName, int Precision, int Scale) GetDbfFieldInfo(Shapefile shapefile, string fieldName)
        {
            try
            {
                var column = shapefile.DataTable.Columns[fieldName];

                if (column is DotSpatial.Data.Field field)
                {
                    string typeName = "Unknown";

                    // .NET 데이터 타입을 직접 비교하는 방식으로 변경
                    Type dotnetType = field.DataType;

                    if (dotnetType == typeof(string))
                    {
                        typeName = "Character";
                    }
                    else if (dotnetType == typeof(double) || dotnetType == typeof(float) || dotnetType == typeof(decimal) ||
                             dotnetType == typeof(int) || dotnetType == typeof(long) || dotnetType == typeof(short) || dotnetType == typeof(byte))
                    {
                        typeName = "Numeric";
                    }
                    else if (dotnetType == typeof(DateTime))
                    {
                        typeName = "Date";
                    }
                    else if (dotnetType == typeof(bool))
                    {
                        typeName = "Logical";
                    }

                    return (typeName, field.Length, field.DecimalCount);
                }

                return ("Not a DBF Field", 0, 0);
            }
            catch
            {
                return ("Error", 0, 0);
            }
        }
        /// <summary>
        /// *** 4. 새로운 헬퍼 메서드 2 ***
        /// 기준 정의의 길이 문자열(예: "50", "9,0", "7,2")을 분석하여 (전체 자릿수, 소수점 자릿수)로 변환합니다.
        /// </summary>
        private (int Precision, int Scale) ParseStandardLength(string lengthString)
        {
            if (string.IsNullOrWhiteSpace(lengthString)) return (0, 0);
            if (lengthString.Contains(","))
            {
                var parts = lengthString.Split(',');
                if (parts.Length == 2 && int.TryParse(parts[0], out int precision) && int.TryParse(parts[1], out int scale))
                {
                    return (precision, scale);
                }
            }
            else
            {
                if (int.TryParse(lengthString, out int precision))
                {
                    return (precision, 0);
                }
            }
            return (0, 0);
        }
        #endregion  
                
        /// <summary>
        /// DBF 파일에서 필드의 타입을 추출
        /// </summary>
        private string GetDbfFieldType(Shapefile shapefile, string fieldName)
        {
            try
            {
                // DotSpatial에서 DBF 파일의 필드 정보에 접근하는 방법
                // DataTable의 컬럼 타입을 확인
                var column = shapefile.DataTable.Columns[fieldName];
                if (column != null)
                {
                    Type colType = column.DataType;
                    if (colType == typeof(double) || colType == typeof(float) || colType == typeof(decimal))
                    {
                        return "NUMBER";
                    }
                    else if (colType == typeof(int) || colType == typeof(long))
                    {
                        return "NUMBER";
                    }
                    else if (colType == typeof(string))
                    {
                        return "VARCHAR2";
                    }
                    else if (colType == typeof(DateTime))
                    {
                        return "DATE";
                    }
                }
                return "VARCHAR2"; // 기본값
            }
            catch
            {
                return "UNKNOWN";
            }
        }

        /// <summary>
        /// DBF 파일에서 필드의 길이 정보를 추출
        /// </summary>
        private string GetDbfFieldLength(Shapefile shapefile, string fieldName)
        {
            try
            {
                // DataTable의 컬럼을 DotSpatial.Data.Field로 캐스팅
                var field = shapefile.DataTable.Columns[fieldName] as DotSpatial.Data.Field;
                if (field != null)
                {
                    if (field.DecimalCount > 0)
                        return $"{field.Length},{field.DecimalCount}"; // 예: "10,2"
                    else
                        return field.Length.ToString(); // 예: "50"
                }

                // Fallback: 일반 DataColumn 처리
                var column = shapefile.DataTable.Columns[fieldName];
                if (column != null && column.DataType == typeof(string))
                {
                    return column.MaxLength > 0 ? column.MaxLength.ToString() : "255";
                }

                return "UNKNOWN";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetDbfFieldLength 오류: {ex.Message}");
                return "ERROR";
            }
        }

        // MainWindow.xaml.cs 파일에 추가할 메서드들

        #region 타이틀바 이벤트 핸들러
        /// <summary>
        /// 타이틀바 드래그로 창 이동
        /// </summary>
        private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }
        }

        /// <summary>
        /// 최소화 버튼 클릭
        /// </summary>
        private void MinimizeButton_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        /// <summary>
        /// 최대화/복원 버튼 클릭
        /// </summary>
        private void MaximizeButton_Click(object sender, RoutedEventArgs e)
        {
            if (this.WindowState == WindowState.Maximized)
            {
                this.WindowState = WindowState.Normal;
            }
            else
            {
                this.WindowState = WindowState.Maximized;
            }
        }

        /// <summary>
        /// 닫기 버튼 클릭
        /// </summary>
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        #endregion

        #region 메뉴 이벤트 핸들러
        // 메뉴 이벤트 핸들러 구현
        private void NewProjectMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var result = CustomMessageBox.Show(this, "새 프로젝트", "새 프로젝트를 생성하시겠습니까?", true);
            if (result == true)
            {
                // TODO: 프로젝트명 입력 다이얼로그 표시
                var projectName = "새 프로젝트"; // 임시
                CurrentProject = ProjectManager.CreateNewProject(projectName);
            }
        }

        private void SaveProjectMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (CurrentProject == null)
            {
                CustomMessageBox.Show(this, "오류", "저장할 프로젝트가 없습니다.");
                return;
            }

            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "PureGIS 프로젝트 파일 (*.pgs)|*.pgs",
                DefaultExt = ".pgs"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    ProjectManager.SaveProject(CurrentProject, saveFileDialog.FileName);
                    CustomMessageBox.Show(this, "완료", "프로젝트가 저장되었습니다.");
                }
                catch (Exception ex)
                {
                    CustomMessageBox.Show(this, "오류", ex.Message);
                }
            }
        }

        private void OpenProjectMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "PureGIS 프로젝트 파일 (*.pgs)|*.pgs",
                DefaultExt = ".pgs"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    CurrentProject = ProjectManager.LoadProject(openFileDialog.FileName);
                    CustomMessageBox.Show(this, "완료", "프로젝트를 불러왔습니다.");
                }
                catch (Exception ex)
                {
                    CustomMessageBox.Show(this, "오류", ex.Message);
                }
            }
        }
        private void ExitMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // 종료 확인 메시지
            var result = CustomMessageBox.Show(this, "종료 확인", "프로그램을 종료하시겠습니까?", true);
            if (result == true)
            {
                this.Close();
            }
        }

        /// <summary>
        /// 프로그램 정보 메뉴 클릭
        /// </summary>
        private void AboutMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // TODO: AboutWindow 창 띄우기 기능 구현
            CustomMessageBox.Show(this, "정보", "프로그램 정보 창은 향후 구현 예정입니다.");
        }
        #endregion
        #region Export Methods

        /// <summary>
        /// QuestPDF로 내보내기
        /// </summary>
        private void ExportQuestPdfButton_Click(object sender, RoutedEventArgs e)
        {
            ExportReport(new QuestPdfExporter());
        }

        /// <summary>
        /// PdfSharp로 내보내기
        /// </summary>
        private void ExportPdfSharpButton_Click(object sender, RoutedEventArgs e)
        {
            ExportReport(new PdfSharpExporter());
        }

        /// <summary>
        /// Word로 내보내기
        /// </summary>
        private void ExportToWordButton_Click(object sender, RoutedEventArgs e)
        {
            ExportReport(new WordExporter());
        }

        /// <summary>
        /// 통합 내보내기 메서드
        /// </summary>
        /// <param name="exporter">사용할 내보내기 구현체</param>
        private void ExportReport(IReportExporter exporter)
        {
            try
            {
                // 1. 검사 결과 데이터 확인
                if (ResultGrid.ItemsSource == null)
                {
                    CustomMessageBox.Show(this, "알림", "내보낼 검사 결과가 없습니다.");
                    return;
                }

                // 2. 보고서 데이터 생성
                var reportData = CreateReportData();
                if (reportData == null)
                {
                    CustomMessageBox.Show(this, "오류", "보고서 데이터 생성에 실패했습니다.");
                    return;
                }

                // 3. 파일 저장 대화상자
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = exporter.FileFilter,
                    DefaultExt = exporter.FileExtension,
                    FileName = $"GeoQC_Report_{DateTime.Now:yyyyMMdd_HHmmss}{exporter.FileExtension}"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    // 4. 내보내기 실행
                    bool success = exporter.Export(reportData, saveFileDialog.FileName);

                    if (success)
                    {
                        CustomMessageBox.Show(this, "완료",
                            $"{exporter.ExporterName} 보고서를 생성했습니다.\n\n" +
                            $"파일: {saveFileDialog.FileName}");
                    }
                    else
                    {
                        CustomMessageBox.Show(this, "오류",
                            $"{exporter.ExporterName} 보고서 생성에 실패했습니다.");
                    }
                }
            }
            catch (Exception ex)
            {
                CustomMessageBox.Show(this, "오류",
                    $"보고서 내보내기 중 오류가 발생했습니다:\n\n" +
                    $"내보내기 방식: {exporter.ExporterName}\n" +
                    $"오류: {ex.Message}");
            }
        }

        /// <summary>
        /// 현재 상태에서 ReportData 객체 생성
        /// </summary>
        /// <returns>생성된 ReportData 객체</returns>
        private ReportData CreateReportData()
        {
            try
            {
                var results = ResultGrid.ItemsSource as List<ColumnValidationResult>;
                if (results == null || results.Count == 0)
                {
                    return null;
                }

                // 파일명 추출
                var fileName = "알 수 없음";
                if (loadedShapefile?.Filename != null)
                {
                    fileName = System.IO.Path.GetFileName(loadedShapefile.Filename);
                }

                // 프로젝트명 추출
                var projectName = CurrentProject?.ProjectName ?? "프로젝트 없음";

                // ReportData 생성
                var reportData = new ReportData
                {
                    ReportDate = DateTime.Now,
                    FileName = fileName,
                    ProjectName = projectName,
                    ValidationResults = new List<ColumnValidationResult>(results)
                };

                return reportData;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CreateReportData 오류: {ex.Message}");
                return null;
            }
        }

        #endregion
    }
}