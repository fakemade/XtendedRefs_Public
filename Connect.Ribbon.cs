using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Forms;
using XtendedRefs.ComAddIn.Application.Commands;
using XtendedRefs.ComAddIn.Infrastructure.Persistence;
using XtendedRefs.ComAddIn.Application.Utilities;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Documents;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using XtendedRefs.ComAddIn.Domain.State;
using XtendedRefs.ComAddIn.Infrastructure.Events;
using XtendedRefs.ComAddIn.Infrastructure.OfficeCom;
using MessageBox = System.Windows.MessageBox;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn;

public sealed partial class Connect
{
    public string GetCustomUI(string ribbonId)
    {
        try
        {
            string xml = LoadRibbonXml("XtendedRefs.ComAddIn.Ribbon.MainRibbon.xml");
            return xml;
        }
        catch (Exception ex)
        {
            return string.Empty;
        }
    }

    public stdole.IPictureDisp GetImage(string imageName)
    {
        try
        {
            Bitmap bmp = (Bitmap)Properties.Resources.ResourceManager.GetObject(imageName);
            if (bmp == null)
            {
                return null;
            }

            return PictureConverter.ImageToPictureDisp(bmp);
        }
        catch (Exception ex)
        {
            return null;
        }
    }

    public void OnRibbonLoad(IRibbonUI ribbonUi)
    {
        _ribbon = ribbonUi;
    }

    public void OnPingButtonPressed(IRibbonControl control)
    {
        string message = AddInContext.IsWordAvailable
            ? "COM add-in loaded and connected to Word."
            : "COM add-in loaded, but Word application is unavailable.";
        MessageBox.Show(message, AddInFriendlyName, System.Windows.MessageBoxButton.OK,
            System.Windows.MessageBoxImage.Information);
    }

    public void OnAddLiteraturePressed(IRibbonControl control) =>
        RunCommand(_commands.LiteratureCommands.AddLiterature);

    public void OnInsertLiteratureReferencePressed(IRibbonControl control) =>
        RunCommand(_commands.LiteratureCommands.InsertLiteratureReference);

    public void OnDeleteLiteraturePressed(IRibbonControl control)=>
        RunCommand(_commands.LiteratureCommands.DeleteLiteratureTextReference);

    public void OnTestPressed(IRibbonControl control)
    {
        CommonCommands.ShowInfo(Singleton<RuntimeStore>.Instance.TableList.Tables.Count.ToString());
    }
    
    public void OnCreateFigurePressed(IRibbonControl control) => RunCommand(_commands.PictureCommands.CreatePicture);
    public void OnDeleteFigurePressed(IRibbonControl control) => RunCommand(_commands.PictureCommands.DeletePicture);
    public void OnInsertFigureReferencePressed(IRibbonControl control) => RunCommand(_commands.PictureCommands.InsertPictureReference);

    public void OnCreateFormulaPressed(IRibbonControl control) => RunCommand(_commands.FormulaCommands.AddFormula);
    public void OnDeleteFormulaPressed(IRibbonControl control) => RunCommand(_commands.FormulaCommands.DeleteFormula);
    public void OnInsertFormulaReferencePressed(IRibbonControl control) => RunCommand(_commands.FormulaCommands.InsertFormula);

    public void OnCreateTableCaptionPressed(IRibbonControl control) =>
        RunCommand(_commands.TableCommands.CreateTable);
    
    public void OnDeleteTableCaptionPressed(IRibbonControl control) =>
        RunCommand(_commands.TableCommands.DeleteTableCaption);

    public void OnInsertTableReferencePressed(IRibbonControl control) =>
        RunCommand(_commands.TableCommands.InsertTableReference);

    public void OnClearXmlPressed(IRibbonControl control)
    {
        Singleton<XmlProvider>.Instance.ClearAllCustomXmlParts();
        MessageBox.Show("Все xml части очищены!");
    }

    public void OnConvertLiteratureBookmarksPressed(IRibbonControl control)
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        Word.Document doc = rs.Document;
        Word.Style normal = doc.Styles[Word.WdBuiltinStyle.wdStyleHyperlink];
        Word.Style visited = doc.Styles[Word.WdBuiltinStyle.wdStyleHyperlinkFollowed];

        normal.Font.ColorIndex = Word.WdColorIndex.wdAuto;
        normal.Font.Underline = Word.WdUnderline.wdUnderlineNone;

        visited.Font.ColorIndex = Word.WdColorIndex.wdAuto;
        visited.Font.Underline = Word.WdUnderline.wdUnderlineNone;
    }

    public void OnAboutProgramPressed(IRibbonControl control)
    {
        Version ver = Assembly.GetExecutingAssembly().GetName().Version;
        MessageBox.Show($"Dev-сборка v{ver.ToString()}");
        
        
    }

    public void OnInsertChapterPressed(IRibbonControl control)
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        try
        {
            RunCommand(_commands.ChapterCommands.InsertChapterLink);
            EventBus<ChapterInfo>.Raise(rs.Document);
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            throw;
        }
    }
    
    public void OnCreateSubBookmarksPressed(IRibbonControl control)
    {
        
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        
        
        try
        {
            
            if (DocumentUtilities.UpdateChapterHyperLinks(rs.Document))
                MessageBox.Show("Номера в ссылках текста обновлены!", "Успех", MessageBoxButton.OK,
                    MessageBoxImage.Information);
            if (DocumentUtilities.UpdateLiteratureHyperLinks(rs.Document))
                MessageBox.Show("Номера в ссылках текста обновлены!", "Успех", MessageBoxButton.OK,
                    MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
        }

    }

    public void OnTestLogicalHeadingNumberingPressed(IRibbonControl control)
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        DocumentSettings documentSettings = rs.DocumentSettings;
        try
        {
            List<ChapterInfo> l = DocumentUtilities.FindNearChaptersByTemplate(rs.Document, SymbolType.Literature);
            foreach (ChapterInfo chapterInfo in l)
            {
                MessageBox.Show(chapterInfo.ToString());
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public void OnScanDocumentPressed(IRibbonControl control)
    {
        RuntimeStore rs = Singleton<RuntimeStore>.Instance;
        DocumentSettings documentSettings = rs.DocumentSettings;
        if (documentSettings.IsScanned)
        {
            MessageBox.Show("Документ уже просканирован, повторное сканирование не требуется", "Внимание",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            rs.ChaptersList.Chapters = DocumentUtilities.ScanChapters(rs.Document, documentSettings.LevelStyles);
            documentSettings.IsScanned = true;
        }
        catch (Exception e)
        {
            MessageBox.Show(e.Message);
        }
    }

    public void OnRescanDocumentPressed(IRibbonControl control)
    {
        try
        {
            RuntimeStore rs = Singleton<RuntimeStore>.Instance;
            DocumentSettings documentSettings = rs.DocumentSettings;
            rs.ChaptersList.Chapters.Clear();
            rs.ChaptersList.Chapters = DocumentUtilities.ScanChapters(rs.Document, documentSettings.LevelStyles);
            MessageBox.Show("Пересканирование прошло успешно!");
        }
        catch (Exception e)
        {
            MessageBox.Show(e.Message);
        }
    }

    
    
    public void DeleteBookmarksByPrefixPressed(IRibbonControl control)
    {
        try
        {
            RuntimeStore rs = Singleton<RuntimeStore>.Instance;
            DocumentSettings documentSettings = rs.DocumentSettings;
            rs.Document.Bookmarks
                .Cast<Word.Bookmark>()
                .Where(b => b.Name.StartsWith("_fm_"))
                .Select(b => b.Name)
                .ToList()
                .ForEach(n => rs.Document.Bookmarks[n].Delete());
            documentSettings.IsScanned = false;
            rs.FormulaList.Formulas.Clear();
        }
        catch (Exception e)
        {
            MessageBox.Show(e.Message);
        }
    }


    

    private void RunCommand(Action action)
    {
        string commandName = action?.Method.Name ?? "<null>";
        try
        {
            if (action != null) action();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, AddInFriendlyName, System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error);
        }
    }

    private sealed class PictureConverter : AxHost
    {
        private PictureConverter() : base("1A2D0B5D-3B55-4F1D-9C69-2D6A54FB2C44")
        {
        }

        public static stdole.IPictureDisp ImageToPictureDisp(Image image)
        {
            return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }
    }

    private static string LoadRibbonXml(string resourceName)
    {
        Assembly assembly = typeof(Connect).Assembly;
        using Stream stream = assembly.GetManifestResourceStream(resourceName)
                              ?? throw new InvalidOperationException($"Ribbon resource '{resourceName}' not found.");
        using StreamReader reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}