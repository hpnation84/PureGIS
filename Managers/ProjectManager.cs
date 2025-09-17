using System;
using System.IO;
using Newtonsoft.Json;
using PureGIS_Geo_QC.Models;

namespace PureGIS_Geo_QC.Managers
{
    public static class ProjectManager
    {
        private static readonly JsonSerializerSettings JsonSettings = new JsonSerializerSettings
        {
            Formatting = Formatting.Indented,
            DateFormatString = "yyyy-MM-dd HH:mm:ss"
        };

        /// <summary>
        /// 프로젝트를 JSON 파일로 저장
        /// </summary>
        public static void SaveProject(ProjectDefinition project, string filePath)
        {
            try
            {
                project.LastModifiedDate = DateTime.Now;
                var json = JsonConvert.SerializeObject(project, JsonSettings);
                File.WriteAllText(filePath, json, System.Text.Encoding.UTF8);
            }
            catch (Exception ex)
            {
                throw new Exception($"프로젝트 저장 실패: {ex.Message}");
            }
        }

        /// <summary>
        /// JSON 파일에서 프로젝트 불러오기
        /// </summary>
        public static ProjectDefinition LoadProject(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                    throw new FileNotFoundException("프로젝트 파일을 찾을 수 없습니다.");

                var json = File.ReadAllText(filePath, System.Text.Encoding.UTF8);
                return JsonConvert.DeserializeObject<ProjectDefinition>(json, JsonSettings);
            }
            catch (Exception ex)
            {
                throw new Exception($"프로젝트 불러오기 실패: {ex.Message}");
            }
        }

        /// <summary>
        /// 새 프로젝트 생성
        /// </summary>
        public static ProjectDefinition CreateNewProject(string projectName)
        {
            var project = new ProjectDefinition
            {
                ProjectName = projectName,
                Description = "",
                CreatedDate = DateTime.Now,
                LastModifiedDate = DateTime.Now
            };

            // 8대 시설물 기본 카테고리 생성
            foreach (var category in InfrastructureTypes.DefaultCategories)
            {
                project.Categories.Add(new InfrastructureCategory
                {
                    CategoryId = category.Key,
                    CategoryName = category.Value
                });
            }

            return project;
        }
    }
}