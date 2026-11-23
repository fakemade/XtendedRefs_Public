using System.Runtime.InteropServices;

namespace XtendedRefs.ComAddIn.Infrastructure.OfficeCom;

[ComVisible(true)]
[Guid("000C0396-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IRibbonExtensibility
{
    [DispId(1)]
    [return: MarshalAs(UnmanagedType.BStr)]
    string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string RibbonID);

    [DispId(2)]
    [return: MarshalAs(UnmanagedType.Interface)]
    public stdole.IPictureDisp GetImage([MarshalAs(UnmanagedType.BStr)] string imageName);
}

[Guid("000C03A7-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IRibbonUI
{
    [DispId(1)] void Invalidate();
    [DispId(2)] void InvalidateControl([MarshalAs(UnmanagedType.BStr)] string ControlID);
    [DispId(3)] void ActivateTab([MarshalAs(UnmanagedType.BStr)] string ControlID);
    [DispId(4)] void ActivateTabMso([MarshalAs(UnmanagedType.BStr)] string ControlID);
    [DispId(5)] void ActivateTabQ([MarshalAs(UnmanagedType.BStr)] string ControlID, [MarshalAs(UnmanagedType.BStr)] string Namespace);
}

[Guid("000C0395-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IRibbonControl
{
    string Id { [return: MarshalAs(UnmanagedType.BStr)] get; }
    object Tag { get; }
    object Context { get; }
}

[ComImport]
[Guid("000C033D-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface ICTPFactory
{
    [return: MarshalAs(UnmanagedType.Interface)]
    object CreateCTP(
        [MarshalAs(UnmanagedType.BStr)] string CTPAxID,
        [MarshalAs(UnmanagedType.BStr)] string CTPTitle,
        [MarshalAs(UnmanagedType.Struct)] object CTPParentWindow);
}

[ComImport]
[Guid("000C033B-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface _CustomTaskPane
{
    bool Visible { get; set; }

    int Width { get; set; }

    MsoCTPDockPosition DockPosition { get; set; }
}

public enum MsoCTPDockPosition
{
    msoCTPDockPositionLeft = 0,
    msoCTPDockPositionTop = 1,
    msoCTPDockPositionRight = 2,
    msoCTPDockPositionBottom = 3,
    msoCTPDockPositionFloating = 4
}