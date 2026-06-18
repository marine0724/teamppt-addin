using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
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
