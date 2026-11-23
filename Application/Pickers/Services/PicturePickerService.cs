using System;
using System.Collections.Generic;
using System.Windows.Interop;
using Microsoft.Office.Interop.Word;
using XtendedRefs.ComAddIn.Application.Contracts;
using XtendedRefs.ComAddIn.Application.Utilities;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using XtendedRefs.ComAddIn.Domain.State;
using XtendedRefs.ComAddIn.UI.Windows;

namespace XtendedRefs.ComAddIn.Application.Pickers.Services
{
    public class PicturePickerService : IImagePickerService
    {
        public DocumentAnchor PickBookmark(Document document)
        {
            

            ImagePickerWindow window = new ImagePickerWindow(document, "Рисунки", LinkType.Picture);

            try
            {
                if (document?.ActiveWindow != null)
                {
                    new WindowInteropHelper(window).Owner = new IntPtr(document.ActiveWindow.Hwnd);
                }
            }
            catch
            {
            }

            bool? dialogResult = window.ShowDialog();
            return dialogResult == true
                ? window.SelectedBookmarkName
                : null;
        }
        public List<PictureInfo> BuildSourceList()
        {
            RuntimeStore rs = Singleton<RuntimeStore>.Instance;
            rs.PictureList.Pictures.Sort((a, b) =>
                DocumentUtilities.CompareParts(a.NumberParts, b.NumberParts)
            );
            return rs.PictureList.Pictures;
        }
    }
}
