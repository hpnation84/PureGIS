using System;
using System.Collections.Generic;
using PureGIS_Geo_QC.Models;

namespace PureGIS_Geo_QC.Exports.Models
{
    /// <summary>
    /// 보고서 생성에 필요한 데이터를 담는 클래스
    /// </summary>
    public class ReportData
    {
        /// <summary>
        /// 보고서 생성 일시
        /// </summary>
        public DateTime ReportDate { get; set; } = DateTime.Now;

        /// <summary>
        /// 검사 파일명
        /// </summary>
        public string FileName { get; set; } = "";

        /// <summary>
        /// 프로젝트명
        /// </summary>
        public string ProjectName { get; set; } = "";

        /// <summary>
        /// 검사 결과 리스트
        /// </summary>
        public List<ColumnValidationResult> ValidationResults { get; set; } = new List<ColumnValidationResult>();

        /// <summary>
        /// 총 필드 수
        /// </summary>
        public int TotalCount => ValidationResults?.Count ?? 0;

        /// <summary>
        /// 정상 필드 수
        /// </summary>
        public int NormalCount => ValidationResults?.FindAll(r => r.Status == "정상").Count ?? 0;

        /// <summary>
        /// 오류 필드 수
        /// </summary>
        public int ErrorCount => ValidationResults?.FindAll(r => r.Status == "오류").Count ?? 0;

        /// <summary>
        /// 정상률 (백분율)
        /// </summary>
        public string SuccessRate
        {
            get
            {
                if (TotalCount == 0) return "0%";
                return ((double)NormalCount / TotalCount * 100).ToString("F1") + "%";
            }
        }

        /// <summary>
        /// 비고 정보 생성
        /// </summary>
        /// <param name="result">검사 결과</param>
        /// <returns>비고 텍스트</returns>
        public static string GetRemarks(ColumnValidationResult result)
        {
            if (result == null) return "";

            var remarks = new List<string>();

            if (!result.IsFieldFound) remarks.Add("필드없음");
            if (!result.IsTypeCorrect) remarks.Add("타입불일치");
            if (!result.IsLengthCorrect) remarks.Add("길이불일치");

            return string.Join(" ", remarks);
        }
    }
}