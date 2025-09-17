using System;
using System.IO;
using System.Threading.Tasks;
using PureGIS_Geo_QC.Exports.Models;
using QuestPDF.Drawing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace PureGIS_Geo_QC.Exports
{
    /// <summary>
    /// QuestPDF를 사용한 PDF 내보내기 구현
    /// </summary>
    public class QuestPdfExporter : IReportExporter
    {
        public string FileExtension => ".pdf";
        public string FileFilter => "PDF 파일 (*.pdf)|*.pdf";
        public string ExporterName => "QuestPDF";

        static QuestPdfExporter()
        {
            // 라이센스 설정 (Community는 무료)
            QuestPDF.Settings.License = LicenseType.Community;
            // --- 이 부분을 추가하세요 ---
            try
            {
                // Windows Fonts 폴더에서 '맑은 고딕' 폰트 파일을 직접 읽어옵니다.
                string fontPath = "C:/Windows/Fonts/malgun.ttf";
                if (File.Exists(fontPath))
                {
                    // QuestPDF의 폰트 관리자에 폰트 파일을 등록합니다.
                    FontManager.RegisterFont(File.OpenRead(fontPath));
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("맑은 고딕 폰트 파일을 찾을 수 없습니다.");
                }
            }
            catch (Exception ex)
            {
                // 폰트 등록 중 오류가 발생해도 프로그램이 중단되지 않도록 처리합니다.
                System.Diagnostics.Debug.WriteLine($"폰트 등록 오류: {ex.Message}");
            }
        }

        public async Task<bool> ExportAsync(ReportData reportData, string filePath)
        {
            return await Task.Run(() => Export(reportData, filePath));
        }

        public bool Export(ReportData reportData, string filePath)
        {
            try
            {
                Document.Create(container =>
                {
                    container.Page(page =>
                    {
                        page.Size(PageSizes.A4.Landscape());
                        page.Margin(1.5f, Unit.Centimetre);
                        // 페이지 전체의 기본 텍스트 스타일을 한글 폰트로 지정합니다.
                        page.DefaultTextStyle(x => x.FontFamily("Malgun Gothic"));
                        // 헤더
                        page.Header().Element(Header);

                        // 내용
                        page.Content().Element(content => Content(content, reportData));

                        // 푸터
                        page.Footer().Element(Footer);
                    });
                })
                .GeneratePdf(filePath);

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"QuestPDF Export Error: {ex.Message}");
                return false;
            }
        }

        private void Header(IContainer container)
        {
            container.Column(column =>
            {
                column.Item().Text("PureGIS GEO-QC SHP데이터 형식 결과 보고서")
                    .FontSize(20).Bold().FontColor(Colors.Blue.Darken1)
                    .AlignCenter();

                column.Item().PaddingTop(10).LineHorizontal(1).LineColor(Colors.Grey.Medium);
            });
        }

        private void Content(IContainer container, ReportData reportData)
        {
            container.PaddingTop(20).Column(column =>
            {
                // 기본 정보
                column.Item().Element(content => BasicInfo(content, reportData));
                column.Item().PaddingTop(20);

                // 요약 통계
                column.Item().Element(content => SummaryStats(content, reportData));
                column.Item().PaddingTop(20);

                // 상세 결과
                column.Item().Element(content => DetailResults(content, reportData));
            });
        }

        private void BasicInfo(IContainer container, ReportData reportData)
        {
            container.Column(column =>
            {
                column.Item().Text("기본 정보").FontSize(16).Bold().FontColor(Colors.Blue.Medium);
                column.Item().PaddingTop(10);

                column.Item().Table(table =>
                {
                    table.ColumnsDefinition(columns =>
                    {
                        columns.RelativeColumn(2);
                        columns.RelativeColumn(5);
                    });

                    AddInfoRow(table, "검사 실행일", reportData.ReportDate.ToString("yyyy년 MM월 dd일 HH시 mm분"));
                    AddInfoRow(table, "검사 파일", reportData.FileName);
                    AddInfoRow(table, "프로젝트", reportData.ProjectName);
                });
            });
        }

        private void SummaryStats(IContainer container, ReportData reportData)
        {
            container.Column(column =>
            {
                column.Item().Text("검사 결과 요약").FontSize(16).Bold().FontColor(Colors.Blue.Medium);
                column.Item().PaddingTop(10);

                column.Item().Table(table =>
                {
                    table.ColumnsDefinition(columns =>
                    {
                        columns.RelativeColumn(1);
                        columns.RelativeColumn(1);
                        columns.RelativeColumn(1);
                        columns.RelativeColumn(1);
                    });

                    // 헤더
                    table.Header(header =>
                    {
                        header.Cell().Element(HeaderCellStyle).Text("전체 필드");
                        header.Cell().Element(HeaderCellStyle).Text("정상");
                        header.Cell().Element(HeaderCellStyle).Text("오류");
                        header.Cell().Element(HeaderCellStyle).Text("정상률");
                    });

                    // 데이터
                    table.Cell().Element(DataCellStyle).Text(reportData.TotalCount.ToString());
                    table.Cell().Element(DataCellStyle).Text(reportData.NormalCount.ToString()).FontColor(Colors.Green.Medium);
                    table.Cell().Element(DataCellStyle).Text(reportData.ErrorCount.ToString()).FontColor(Colors.Red.Medium);
                    table.Cell().Element(DataCellStyle).Text(reportData.SuccessRate);
                });
            });
        }

        private void DetailResults(IContainer container, ReportData reportData)
        {
            container.Column(column =>
            {
                column.Item().Text("상세 검사 결과").FontSize(16).Bold().FontColor(Colors.Blue.Medium);
                column.Item().PaddingTop(10);

                column.Item().Table(table =>
                {
                    table.ColumnsDefinition(columns =>
                    {
                        columns.RelativeColumn(0.8f); // 상태
                        columns.RelativeColumn(1.2f); // 기준컬럼ID
                        columns.RelativeColumn(1.5f); // 기준컬럼명
                        columns.RelativeColumn(1.0f); // 기준타입
                        columns.RelativeColumn(0.8f); // 기준길이
                        columns.RelativeColumn(1.2f); // 찾은필드명
                        columns.RelativeColumn(1.0f); // 파일타입
                        columns.RelativeColumn(0.8f); // 파일길이
                        columns.RelativeColumn(1.5f); // 비고
                    });

                    // 헤더
                    table.Header(header =>
                    {
                        header.Cell().Element(HeaderCellStyle).Text("상태");
                        header.Cell().Element(HeaderCellStyle).Text("기준컬럼ID");
                        header.Cell().Element(HeaderCellStyle).Text("기준컬럼명");
                        header.Cell().Element(HeaderCellStyle).Text("기준타입");
                        header.Cell().Element(HeaderCellStyle).Text("기준길이");
                        header.Cell().Element(HeaderCellStyle).Text("찾은필드명");
                        header.Cell().Element(HeaderCellStyle).Text("파일타입");
                        header.Cell().Element(HeaderCellStyle).Text("파일길이");
                        header.Cell().Element(HeaderCellStyle).Text("비고");
                    });

                    // 데이터 행
                    foreach (var result in reportData.ValidationResults)
                    {
                        var statusColor = result.Status == "정상" ? Colors.Green.Medium : Colors.Red.Medium;

                        table.Cell().Element(DataCellStyle).Text(result.Status ?? "").FontColor(statusColor);
                        table.Cell().Element(DataCellStyle).Text(result.Std_ColumnId ?? "");
                        table.Cell().Element(DataCellStyle).Text(result.Std_ColumnName ?? "");
                        table.Cell().Element(DataCellStyle).Text(result.Std_Type ?? "");
                        table.Cell().Element(DataCellStyle).Text(result.Std_Length ?? "");
                        table.Cell().Element(DataCellStyle).Text(result.Found_FieldName ?? "");
                        table.Cell().Element(DataCellStyle).Text(result.Cur_Type ?? "");
                        table.Cell().Element(DataCellStyle).Text(result.Cur_Length ?? "");
                        table.Cell().Element(DataCellStyle).Text(ReportData.GetRemarks(result));
                    }
                });
            });
        }

        private void Footer(IContainer container)
        {
            container.AlignRight().Text($"보고서 생성: {DateTime.Now:yyyy-MM-dd HH:mm:ss} | PureGIS GEO-QC v1.0 | {ExporterName}")
                .FontSize(10).FontColor(Colors.Grey.Medium);
        }

        private void AddInfoRow(TableDescriptor table, string label, string value)
        {
            table.Cell().Element(InfoLabelCellStyle).Text(label);
            table.Cell().Element(InfoValueCellStyle).Text(value);
        }

        // 스타일 정의
        private IContainer HeaderCellStyle(IContainer container) =>
            container.DefaultTextStyle(x => x.FontSize(10).Bold().FontColor(Colors.White))
                     .Background(Colors.Blue.Medium)
                     .Padding(6)
                     .AlignCenter();

        private IContainer DataCellStyle(IContainer container) =>
            container.DefaultTextStyle(x => x.FontSize(9))
                     .Padding(4)
                     .BorderBottom(0.5f)
                     .BorderColor(Colors.Grey.Lighten2)
                     .AlignCenter();

        private IContainer InfoLabelCellStyle(IContainer container) =>
            container.DefaultTextStyle(x => x.FontSize(11).Bold())
                     .Background(Colors.Grey.Lighten3)
                     .Padding(8)
                     .BorderBottom(0.5f)
                     .BorderColor(Colors.Grey.Medium);

        private IContainer InfoValueCellStyle(IContainer container) =>
            container.DefaultTextStyle(x => x.FontSize(11))
                     .Padding(8)
                     .BorderBottom(0.5f)
                     .BorderColor(Colors.Grey.Medium);
    }
}