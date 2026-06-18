using System;
using System.Runtime.InteropServices;

namespace TeampptAddin
{
    /// <summary>
    /// COM IObjectSafety 인터페이스.
    /// TaskPaneHost가 ActiveX로 호스팅될 때 PowerPoint가 보안 검증용으로 호출.
    /// GetInterfaceSafetyOptions/SetInterfaceSafetyOptions에서 항상 "안전"으로 응답하여
    /// Task Pane 로딩이 차단되지 않도록 함.
    /// </summary>
    [ComImport]
    [Guid("CB5BDC81-93C1-11CF-8F20-00805F2CD064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IObjectSafety
    {
        [PreserveSig]
        int GetInterfaceSafetyOptions(ref Guid riid, out int pdwSupportedOptions, out int pdwEnabledOptions);

        [PreserveSig]
        int SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions);
    }
}
