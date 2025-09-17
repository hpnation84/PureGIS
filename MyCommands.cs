using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

using Autodesk.Gis.Map;
using Autodesk.Gis.Map.ObjectData;
using Autodesk.Gis.Map.Topology;

using PureGIS_Geo_QC.WPF;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Exception = Autodesk.AutoCAD.Runtime.Exception;

namespace PureGIS_Geo_QC
{
    public class MyCommands
    {
        [CommandMethod("HELLO")]
        public void HelloCommand()
        {
            var ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("\nHello AutoCAD 2023 from VS2022!");
        }

        // MainWindow 인스턴스를 관리하기 위한 정적(static) 변수
        private static MainWindow qcWindow = null;

        [CommandMethod("GEOQC")]
        public void ShowQcWindowCommand()
        {
            var ed = Application.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                // 1. 창이 아직 열리지 않았거나, 닫혔는지 확인
                if (qcWindow == null)
                {
                    // 2. 새 창 인스턴스 생성
                    qcWindow = new MainWindow();

                    // 중요: 창이 닫힐 때 변수를 null로 리셋하는 이벤트 핸들러 추가
                    qcWindow.Closed += (sender, e) => {
                        qcWindow = null;
                        ed.WriteMessage("\n[GEOQC] 창이 닫혔습니다.");
                    };

                    // 3. AutoCAD 전용 메서드를 사용하여 모달리스 창으로 띄우기
                    Application.ShowModelessWindow(qcWindow);
                    ed.WriteMessage("\n[GEOQC] 새 창을 열었습니다.");
                }
                else
                {
                    // 4. 창이 이미 열려있다면, 활성화하여 맨 앞으로 가져오기
                    try
                    {
                        qcWindow.Activate();
                        qcWindow.WindowState = System.Windows.WindowState.Normal; // 최소화되어 있을 경우 복원
                        qcWindow.Topmost = true;  // 잠시 맨 앞으로
                        qcWindow.Topmost = false; // 다시 일반 상태로
                        ed.WriteMessage("\n[GEOQC] 기존 창을 활성화했습니다.");
                    }
                    catch (Exception activateEx)
                    {
                        // 창 활성화에 실패하면 (창이 이미 닫혔을 가능성) 새로 생성
                        ed.WriteMessage($"\n[GEOQC] 창 활성화 실패, 새로 생성: {activateEx.Message}");
                        qcWindow = null;
                        ShowQcWindowCommand(); // 재귀 호출로 새 창 생성
                    }
                }
            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\n[GEOQC] 오류 발생: {ex.Message}");
                ed.WriteMessage($"\n상세: {ex.StackTrace}");

                // 오류 발생 시 변수 리셋
                qcWindow = null;
            }
        }
    }
}