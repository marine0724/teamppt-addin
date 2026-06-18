using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    /// <summary>
    /// COM Add-in 진입점. PowerPoint가 시작될 때 이 클래스를 로드.
    ///
    /// IDTExtensibility2: Add-in 생명주기 관리
    /// - OnConnection: PowerPoint Application 참조를 Globals에 저장
    /// - OnDisconnection: 참조 해제
    ///
    /// ICustomTaskPaneConsumer: Task Pane 생성
    /// - CTPFactoryAvailable: CreateCTP("TeampptAddin.TaskPaneHost")로 오른쪽 패널 생성
    ///   → TaskPaneHost가 ActiveX 컨트롤로 호스팅됨
    ///
    /// 레지스트리 등록:
    /// HKCU\Software\Microsoft\Office\PowerPoint\Addins\TeampptAddin.Connect
    /// LoadBehavior=3 (시작 시 자동 로드)
    /// </summary>
    [ComVisible(true)]
    [Guid("7B3A4D1E-9F2C-4A85-B6D0-3E8F1C5A7B92")]
    [ProgId("TeampptAddin.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Connect : IDTExtensibility2, ICustomTaskPaneConsumer
    {
        private PowerPoint.Application _app;
        private _CustomTaskPane _taskPane;

        #region IDTExtensibility2

        public void OnConnection(object Application, ext_ConnectMode ConnectMode,
            object AddInInst, ref Array custom)
        {
            _app = (PowerPoint.Application)Application;
            Globals.Application = _app;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            _app = null;
            Globals.Application = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom) { }

        public void OnBeginShutdown(ref Array custom)
        {
            if (_taskPane != null)
                _taskPane.Visible = false;
        }

        #endregion

        #region ICustomTaskPaneConsumer

        public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst)
        {
            try
            {
                _taskPane = CTPFactoryInst.CreateCTP(
                    "TeampptAddin.TaskPaneHost",
                    "TEAMPPT");
                _taskPane.Width = 320;
                _taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _taskPane.Visible = true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"TEAMPPT Task Pane 생성 실패:\n{ex.Message}",
                    "TEAMPPT", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        #endregion
    }
}
