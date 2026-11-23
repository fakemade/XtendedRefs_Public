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
    public class FormulaPickerService : IImagePickerService
    {
        public DocumentAnchor PickBookmark(Document document)
        {
            ImagePickerWindow window = new ImagePickerWindow(document, "Формулы", LinkType.Formula);
            
            bool? dialogResult = window.ShowDialog();
            return dialogResult == true
                ? window.SelectedBookmarkName
                : null;
        }
        
        public List<FormulaInfo> BuildSourceList()
        {
            RuntimeStore rs = Singleton<RuntimeStore>.Instance;
            rs.FormulaList.Formulas.Sort((a, b) =>
                DocumentUtilities.CompareParts(a.NumberParts, b.NumberParts)
            );
            return rs.FormulaList.Formulas;
        }
    }
}
