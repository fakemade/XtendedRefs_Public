using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using XtendedRefs.ComAddIn.Application.Contracts;
using XtendedRefs.ComAddIn.Application.Utilities;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using XtendedRefs.ComAddIn.Domain.State;
using XtendedRefs.ComAddIn.UI.Windows;

namespace XtendedRefs.ComAddIn.Application.Pickers.Services
{
    public class TablePickerService : IImagePickerService
    {
        public DocumentAnchor PickBookmark(Document document)
        {
            ImagePickerWindow window = new ImagePickerWindow(document, "Таблицы", LinkType.Table);
            bool? dialogResult = window.ShowDialog();
            return dialogResult == true
                ? window.SelectedBookmarkName
                : null;
        }
        
        public List<TableInfo> BuildSourceList()
        {
            RuntimeStore rs = Singleton<RuntimeStore>.Instance;
            rs.TableList.Tables.Sort((a, b) =>
                DocumentUtilities.CompareParts(a.NumberParts, b.NumberParts)
            );
            return rs.TableList.Tables;
        }
    }
}
