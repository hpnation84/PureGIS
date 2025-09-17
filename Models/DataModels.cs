using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PureGIS_Geo_QC.Models
{
    internal class DataModels
    {
    }
    // 프로젝트 최상위 클래스
    public class ProjectDefinition
    {
        public string ProjectName { get; set; }        // 지역명/기관명
        public string Description { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime LastModifiedDate { get; set; }
        public List<InfrastructureCategory> Categories { get; set; } = new List<InfrastructureCategory>();
    }

    // 8대 지하시설물 분류
    public class InfrastructureCategory
    {
        public string CategoryId { get; set; }     // ROAD, WATER, SEWER, ELECTRIC, TELECOM, HEAT, GAS, OIL
        public string CategoryName { get; set; }   // 도로, 상수, 하수, 전기, 통신, 열배관, 가스, 송유
        public List<TableDefinition> Tables { get; set; } = new List<TableDefinition>();
    }

    // 8대 시설물 기본 목록 제공용
    public static class InfrastructureTypes
    {
        public static readonly Dictionary<string, string> DefaultCategories = new Dictionary<string, string>
    {
        {"ROAD", "도로"},
        {"WATER", "상수"},
        {"SEWER", "하수"},
        {"ELECTRIC", "전기"},
        {"TELECOM", "통신"},
        {"HEAT", "열배관"},
        {"GAS", "가스"},
        {"OIL", "송유"}
    };
    }

    // 컬럼의 속성을 정의하는 클래스
    public class ColumnDefinition
    {
        public string ColumnId { get; set; }
        public string ColumnName { get; set; }
        public string Type { get; set; }
        public string Length { get; set; }
        public bool IsNotNull { get; set; }
        public string KeyType { get; set; } // PK/FK
        public string CodeName { get; set; }
        public string Remarks { get; set; }
    }

    // 테이블의 속성을 정의하는 클래스
    public class TableDefinition
    {
        public string TableId { get; set; }
        public string TableName { get; set; }
        public List<ColumnDefinition> Columns { get; set; } = new List<ColumnDefinition>();
    }        
    // 검증 결과를 담는 클래스
    public class ColumnValidationResult
    {
        // 전체 상태 (정상/오류)
        public string Status { get; set; }

        // 기준(Standard) 값 - ✅ Std_ColumnId는 이미 있음
        public string Std_ColumnId { get; set; }
        public string Std_ColumnName { get; set; }
        public string Std_Type { get; set; }
        public string Std_Length { get; set; }

        // ❌ 추가 필요한 필드들
        public string Found_FieldName { get; set; }     // 실제 찾은 필드명
        public bool IsFieldFound { get; set; }          // 필드 존재 여부

        // 현재(Current) 값
        public string Cur_Type { get; set; }
        public string Cur_Length { get; set; }

        // 각 항목의 일치 여부
        public bool IsTypeCorrect { get; set; }
        public bool IsLengthCorrect { get; set; }
    }

    /// <summary>
    /// 파일의 컬럼 정보를 표시하기 위한 클래스
    /// </summary>
    public class FileColumnInfo
    {
        public string ColumnName { get; set; }
        public TypeInfo DataType { get; set; }
        public string MaxLength { get; set; }
    }

    /// <summary>
    /// 타입 정보를 표시하기 위한 클래스
    /// </summary>
    public class TypeInfo
    {
        public string Name { get; set; }
    }
}