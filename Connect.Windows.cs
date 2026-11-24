using System;
using System.Windows;
using System.Windows.Interop;
using XtendedRefs.ComAddIn.Infrastructure.OfficeCom;
using XtendedRefs.ComAddIn.UI.Windows;
using Word = Microsoft.Office.Interop.Word;
using MessageBox = System.Windows.MessageBox;

namespace XtendedRefs.ComAddIn;

public sealed partial class Connect
{
    public void OnSettingsPressed(IRibbonControl control)
    {
        Word.Application application = AddInContext.Application;
        if (application == null)
        {
            MessageBox.Show("Word РЅРµРґРѕСЃС‚СѓРїРµРЅ.", AddInFriendlyName, MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        Word.Document activeDoc;
        try
        {
            activeDoc = application.ActiveDocument;
        }
        catch
        {
            MessageBox.Show("РќРµС‚ РѕС‚РєСЂС‹С‚РѕРіРѕ РґРѕРєСѓРјРµРЅС‚Р°.", AddInFriendlyName, MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (activeDoc == null)
        {
            MessageBox.Show("РќРµС‚ РѕС‚РєСЂС‹С‚РѕРіРѕ РґРѕРєСѓРјРµРЅС‚Р°.", AddInFriendlyName, MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        
        SettingsWindow window = new SettingsWindow(activeDoc);
        try
        {
            if (activeDoc.ActiveWindow != null)
            {
                new WindowInteropHelper(window).Owner = new IntPtr(activeDoc.ActiveWindow.Hwnd);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("Р§С‚Рѕ-С‚Рѕ РЅРµ С‚Р°Рє! РћС€РёР±РєР°:\n" + ex.Message, "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        try
        {
            window.ShowDialog();
        }
        catch (Exception e)
        {
            MessageBox.Show(e.Message, "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        
    }

    private static string SafeDocumentName(Word.Document document)
    {
        try
        {
            return document?.FullName ?? "<null>";
        }
        catch
        {
            return "<unavailable>";
        }
    }
}
