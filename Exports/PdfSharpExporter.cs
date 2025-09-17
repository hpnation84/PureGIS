using System;
using System.Threading.Tasks;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using PureGIS_Geo_QC.Exports.Models;

namespace PureGIS_Geo_QC.Exports
{
    /// <summary>
    /// PdfSharpCore를 사용한 PDF 내보내기 구현
    /// </summary>
    public class PdfSharpExporter : IReportExporter
    {
        public string FileExtension => ".pdf";
        public string FileFilter => "PDF 파일 (*.pdf)|*.pdf";
        public string ExporterName => "PdfSharpCore";

        public async Task<bool> ExportAsync(ReportData reportData, string filePath)
        {
            return await Task.Run(() => Export(reportData, filePath));
        }

        public bool Export(ReportData reportData, string filePath)
        {
            try
            {
                var document = new PdfDocument();
                var page = document.AddPage();
                page.Orientation = PdfSharpCore.PageOrientation.Landscape; // 가로 방향
                var graphics = XGraphics.FromPdfPage(page);

                // 폰트 설정
                var titleFont = new XFont("Malgun Gothic", 18, XFontStyle.Bold);
                var headerFont = new XFont("Malgun Gothic", 12, XFontStyle.Bold);
                var normalFont = new XFont("Malgun Gothic", 10);
                var smallFont = new XFont("Malgun Gothic", 8);

                double yPos = 40;
                double leftMargin = 40;
                double pageWidth = page.Width - 80;

                // 제목
                DrawCenteredText(graphics, "PureGIS GEO-QC SHP데이터 형식 결과 보고서",
                    titleFont, XBrushes.DarkBlue, leftMargin, pageWidth, yPos);
                yPos += 35;

                // 구분선
                graphics.DrawLine(new XPen(XColors.Gray, 1), leftMargin, yPos, leftMargin + pageWidth, yPos);
                yPos += 20;

                // 기본 정보
                yPos = DrawBasicInfo(graphics, reportData, headerFont, normalFont, leftMargin, yPos);
                yPos += 20;

                // 요약 통계
                yPos = DrawSummaryStats(graphics, reportData, headerFont, normalFont, leftMargin, pageWidth, yPos);
                yPos += 25;

                // 상세 결과
                yPos = DrawDetailResults(graphics, reportData, headerFont, smallFont, leftMargin, pageWidth, yPos, page.Height);

                // 푸터
                DrawFooter(graphics, smallFont, leftMargin, pageWidth, page.Height);

                graphics.Dispose();
                document.Save(filePath);
                document.Close();
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"PdfSharp Export Error: {ex.Message}");
                return false;
            }
        }

        private double DrawBasicInfo(XGraphics graphics, ReportData reportData, XFont headerFont, XFont normalFont, double leftMargin, double yPos)
        {
            graphics.DrawString("기본 정보", headerFont, XBrushes.DarkBlue, leftMargin, yPos);
            yPos += 25;

            var infoItems = new[]
            {
                ("검사 실행일:", reportData.ReportDate.ToString("yyyy년 MM월 dd일 HH시 mm분")),
                ("검사 파일:", reportData.FileName),
                ("프로젝트:", reportData.ProjectName)
            };

            foreach (var (label, value) in infoItems)
            {
                graphics.DrawString(label, normalFont, XBrushes.Black, leftMargin, yPos);
                graphics.DrawString(value, normalFont, XBrushes.Black, leftMargin + 120, yPos);
                yPos += 18;
            }

            return yPos;
        }

        private double DrawSummaryStats(XGraphics graphics, ReportData reportData, XFont headerFont, XFont normalFont, double leftMargin, double pageWidth, double yPos)
        {
            graphics.DrawString("검사 결과 요약", headerFont, XBrushes.DarkBlue, leftMargin, yPos);
            yPos += 25;

            var headers = new[] { "전체 필드", "정상", "오류", "정상률" };
            var data = new[] { reportData.TotalCount.ToString(), reportData.NormalCount.ToString(), reportData.ErrorCount.ToString(), reportData.SuccessRate };
            var colors = new[] { XBrushes.Black, XBrushes.Green, XBrushes.Red, XBrushes.Black };

            double colWidth = pageWidth / 4;

            // 헤더 배경
            var headerRect = new XRect(leftMargin, yPos, pageWidth, 20);
            graphics.DrawRectangle(XBrushes.LightGray, headerRect);

            // 헤더 텍스트
            for (int i = 0; i < headers.Length; i++)
            {
                var rect = new XRect(leftMargin + i * colWidth, yPos, colWidth, 20);
                DrawCenteredText(graphics, headers[i], normalFont, XBrushes.Black, rect);
            }
            yPos += 20;

            // 데이터 배경
            var dataRect = new XRect(leftMargin, yPos, pageWidth, 20);
            graphics.DrawRectangle(XPens.Black, dataRect);

            // 데이터 텍스트
            for (int i = 0; i < data.Length; i++)
            {
                var rect = new XRect(leftMargin + i * colWidth, yPos, colWidth, 20);
                DrawCenteredText(graphics, data[i], normalFont, colors[i], rect);
            }

            return yPos + 20;
        }

        private double DrawDetailResults(XGraphics graphics, ReportData reportData, XFont headerFont, XFont smallFont, double leftMargin, double pageWidth, double yPos, double pageHeight)
        {
            graphics.DrawString("상세 검사 결과", headerFont, XBrushes.DarkBlue, leftMargin, yPos);
            yPos += 25;

            var headers = new[] { "상태", "기준컬럼ID", "기준컬럼명", "기준타입", "기준길이", "찾은필드명", "파일타입", "파일길이", "비고" };
            var colWidths = new double[] { 40, 70, 90, 60, 50, 70, 60, 50, 90 };

            // 테이블 헤더
            double xPos = leftMargin;
            var headerBgRect = new XRect(leftMargin, yPos, pageWidth, 18);
            graphics.DrawRectangle(XBrushes.LightGray, headerBgRect);

            for (int i = 0; i < headers.Length; i++)
            {
                var rect = new XRect(xPos, yPos, colWidths[i], 18);
                graphics.DrawRectangle(XPens.Black, rect);
                DrawCenteredText(graphics, headers[i], smallFont, XBrushes.Black, rect);
                xPos += colWidths[i];
            }
            yPos += 18;

            // 데이터 행 (페이지에 맞게 제한)
            int maxRows = (int)((pageHeight - yPos - 60) / 15); // 여유 공간 확보
            var displayResults = reportData.ValidationResults.Count > maxRows ?
                reportData.ValidationResults.GetRange(0, maxRows) :
                reportData.ValidationResults;

            foreach (var result in displayResults)
            {
                xPos = leftMargin;
                var rowData = new[]
                {
                    result.Status ?? "",
                    result.Std_ColumnId ?? "",
                    result.Std_ColumnName ?? "",
                    result.Std_Type ?? "",
                    result.Std_Length ?? "",
                    result.Found_FieldName ?? "",
                    result.Cur_Type ?? "",
                    result.Cur_Length ?? "",
                    ReportData.GetRemarks(result)
                };

                for (int i = 0; i < rowData.Length; i++)
                {
                    var rect = new XRect(xPos, yPos, colWidths[i], 15);
                    graphics.DrawRectangle(XPens.Black, rect);

                    var brush = (i == 0 && result.Status == "오류") ? XBrushes.Red : XBrushes.Black;
                    DrawCenteredText(graphics, TruncateText(rowData[i], colWidths[i]), smallFont, brush, rect);
                    xPos += colWidths[i];
                }
                yPos += 15;
            }

            // 페이지 제한으로 인한 생략 안내
            if (reportData.ValidationResults.Count > maxRows)
            {
                yPos += 10;
                graphics.DrawString($"※ 페이지 공간 제한으로 {maxRows}개 행만 표시됨 (전체: {reportData.ValidationResults.Count}개)",
                    smallFont, XBrushes.Red, leftMargin, yPos);
            }

            return yPos;
        }

        private void DrawFooter(XGraphics graphics, XFont smallFont, double leftMargin, double pageWidth, double pageHeight)
        {
            var footerText = $"보고서 생성: {DateTime.Now:yyyy-MM-dd HH:mm:ss} | PureGIS GEO-QC v1.0 | {ExporterName}";
            var footerY = pageHeight - 30;
            graphics.DrawString(footerText, smallFont, XBrushes.Gray, leftMargin + pageWidth - 300, footerY);
        }

        private void DrawCenteredText(XGraphics graphics, string text, XFont font, XBrush brush, double left, double width, double y)
        {
            var rect = new XRect(left, y, width, font.Height);
            graphics.DrawString(text, font, brush, rect, XStringFormats.Center);
        }

        private void DrawCenteredText(XGraphics graphics, string text, XFont font, XBrush brush, XRect rect)
        {
            graphics.DrawString(text, font, brush, rect, XStringFormats.Center);
        }

        private string TruncateText(string text, double maxWidth)
        {
            if (string.IsNullOrEmpty(text)) return "";

            // 대략적인 문자 수 제한 (컬럼 너비에 따라)
            int maxLength = (int)(maxWidth / 6); // 폰트 크기 8 기준 추정
            return text.Length > maxLength ? text.Substring(0, maxLength - 2) + ".." : text;
        }
    }
}