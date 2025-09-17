using System.Collections.Generic;
using System.Threading.Tasks;
using PureGIS_Geo_QC.Exports.Models;

namespace PureGIS_Geo_QC.Exports
{
    /// <summary>
    /// 보고서 내보내기 인터페이스
    /// </summary>
    public interface IReportExporter
    {
        /// <summary>
        /// 지원하는 파일 확장자
        /// </summary>
        string FileExtension { get; }

        /// <summary>
        /// 파일 필터 (SaveFileDialog용)
        /// </summary>
        string FileFilter { get; }

        /// <summary>
        /// 내보내기 방식 이름
        /// </summary>
        string ExporterName { get; }

        /// <summary>
        /// 보고서 내보내기 실행
        /// </summary>
        /// <param name="reportData">보고서 데이터</param>
        /// <param name="filePath">저장 경로</param>
        /// <returns>성공 여부</returns>
        Task<bool> ExportAsync(ReportData reportData, string filePath);

        /// <summary>
        /// 동기 내보내기 (호환성용)
        /// </summary>
        /// <param name="reportData">보고서 데이터</param>
        /// <param name="filePath">저장 경로</param>
        /// <returns>성공 여부</returns>
        bool Export(ReportData reportData, string filePath);
    }
}