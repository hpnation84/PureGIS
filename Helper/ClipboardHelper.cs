using System;
using System.Collections.Generic;
using System.Linq;
using PureGIS_Geo_QC.Models;

namespace PureGIS_Geo_QC.Helpers
{
    /// <summary>
    /// 클립보드 데이터 파싱 관련 헬퍼 클래스
    /// </summary>
    public static class ClipboardHelper
    {
        /// <summary>
        /// 클립보드 텍스트를 테이블 정의 리스트로 파싱
        /// </summary>
        /// <param name="clipboardText">클립보드에서 가져온 텍스트</param>
        /// <returns>파싱된 테이블 정의 리스트</returns>
        public static List<TableDefinition> ParseClipboardData(string clipboardText)
        {
            var standardTables = new List<TableDefinition>();

            if (string.IsNullOrWhiteSpace(clipboardText))
                return standardTables;

            // 디버깅을 위해 클립보드 내용 확인
            System.Diagnostics.Debug.WriteLine("클립보드 내용:");
            System.Diagnostics.Debug.WriteLine(clipboardText);

            // 1. "---"로 구분된 여러 테이블 처리
            string[] tablesData = clipboardText.Split(new[] { "---" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var tableText in tablesData)
            {
                if (string.IsNullOrWhiteSpace(tableText)) continue;

                var parsedTable = ParseSingleTable(tableText);
                if (parsedTable != null && parsedTable.Columns.Count > 0)
                {
                    standardTables.Add(parsedTable);
                }
            }

            // 2. 만약 "---" 구분자가 없는 단순한 엑셀 데이터라면
            if (standardTables.Count == 0)
            {
                var simpleTable = ParseSimpleExcelData(clipboardText);
                if (simpleTable != null && simpleTable.Columns.Count > 0)
                {
                    standardTables.Add(simpleTable);
                }
            }

            return standardTables;
        }

        /// <summary>
        /// 단일 테이블 텍스트를 TableDefinition으로 파싱
        /// </summary>
        /// <param name="tableText">테이블 텍스트</param>
        /// <returns>파싱된 TableDefinition</returns>
        private static TableDefinition ParseSingleTable(string tableText)
        {
            var lines = tableText.Trim().Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length < 3) return null;

            var newTable = new TableDefinition();

            // 첫 번째 줄에서 TableId 안전하게 추출
            if (lines.Length > 0 && lines[0].Contains(':'))
            {
                var tableParts = lines[0].Split(':');
                newTable.TableId = tableParts.Length > 1 ? tableParts[1].Trim() : "Unknown";
            }
            else
            {
                newTable.TableId = "Table_" + DateTime.Now.Ticks.ToString().Substring(0, 6);
            }

            // 두 번째 줄에서 TableName 안전하게 추출
            if (lines.Length > 1 && lines[1].Contains(':'))
            {
                var nameParts = lines[1].Split(':');
                newTable.TableName = nameParts.Length > 1 ? nameParts[1].Trim() : "Unknown Table";
            }
            else
            {
                newTable.TableName = "테이블_" + newTable.TableId;
            }

            // 나머지 줄들에서 컬럼 정보 추출
            for (int i = 2; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;

                var cols = lines[i].Split('\t');
                if (cols.Length < 2) continue; // 최소 2개 컬럼은 있어야 함

                newTable.Columns.Add(new ColumnDefinition
                {
                    ColumnId = GetSafeArrayValue(cols, 0, "COL_" + i),
                    ColumnName = GetSafeArrayValue(cols, 1, "컬럼_" + i),
                    Type = GetSafeArrayValue(cols, 2, "VARCHAR2"),
                    Length = GetSafeArrayValue(cols, 3, "50"),
                    IsNotNull = GetSafeArrayValue(cols, 4, "N").ToUpper() == "Y",
                    KeyType = GetSafeArrayValue(cols, 5, "")
                });
            }

            return newTable;
        }

        /// <summary>
        /// 단순한 엑셀 데이터 (헤더 + 데이터 행들) 파싱
        /// </summary>
        /// <param name="clipboardText">클립보드 텍스트</param>
        /// <returns>파싱된 TableDefinition</returns>
        private static TableDefinition ParseSimpleExcelData(string clipboardText)
        {
            var lines = clipboardText.Trim().Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length < 2) return null; // 최소 헤더 + 1개 행은 있어야 함

            // 첫 번째 줄이 헤더인지 확인 (컬럼ID, 컬럼명, 타입, 길이 등의 헤더가 있는지)
            var headerLine = lines[0];
            var isHeaderRow = headerLine.ToLower().Contains("컬럼") ||
                              headerLine.ToLower().Contains("column") ||
                              headerLine.ToLower().Contains("타입") ||
                              headerLine.ToLower().Contains("type");

            var newTable = new TableDefinition
            {
                TableId = "EXCEL_TABLE_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"),
                TableName = "엑셀에서 가져온 테이블"
            };

            int startRow = isHeaderRow ? 1 : 0; // 헤더가 있으면 1번째 줄부터, 없으면 0번째 줄부터

            for (int i = startRow; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;

                var cols = lines[i].Split('\t');
                if (cols.Length < 2) continue;

                newTable.Columns.Add(new ColumnDefinition
                {
                    ColumnId = GetSafeArrayValue(cols, 0, "COL_" + (i - startRow + 1)),
                    ColumnName = GetSafeArrayValue(cols, 1, "컬럼_" + (i - startRow + 1)),
                    Type = GetSafeArrayValue(cols, 2, "VARCHAR2"),
                    Length = GetSafeArrayValue(cols, 3, "50"),
                    IsNotNull = GetSafeArrayValue(cols, 4, "N").ToUpper() == "Y",
                    KeyType = GetSafeArrayValue(cols, 5, "")
                });
            }

            return newTable;
        }

        /// <summary>
        /// 배열에서 안전하게 값을 가져오는 헬퍼 함수
        /// </summary>
        /// <param name="array">대상 배열</param>
        /// <param name="index">인덱스</param>
        /// <param name="defaultValue">기본값</param>
        /// <returns>안전하게 추출된 값</returns>
        public static string GetSafeArrayValue(string[] array, int index, string defaultValue)
        {
            if (array != null && index >= 0 && index < array.Length)
            {
                return string.IsNullOrWhiteSpace(array[index]) ? defaultValue : array[index].Trim();
            }
            return defaultValue;
        }
    }
}