using System;
using System.Windows;
using XtendedRefs.ComAddIn.Domain.Documents;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using XtendedRefs.ComAddIn.Infrastructure.Persistence;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Domain.State;

public class RuntimeStore
{
    private readonly XmlProvider _provider;

    public Word.Document Document { get; }

    public DocumentSettings DocumentSettings { get; private set; }
    
    public ChaptersList ChaptersList { get; private set; }
    
    public LiteratureList LiteratureList { get; private set; }
    
    public FormulaList FormulaList { get; private set; }
    
    public PictureList PictureList { get; private set; }
    
    public TableList TableList { get;  private set; }
    
    
    /// <summary>
    /// Основной конструктор хранилища всех объектов в рантайме
    /// </summary>
    /// <param name="document">Документ, для которого создается хранилище</param>
    /// <param name="provider">Класс для загрузки CustomXMLParts из документа</param>
    public RuntimeStore(Word.Document document, XmlProvider provider)
    {
        Document = document;
        _provider = provider;
    }

    public void Init()
    {

        DocumentSettings = _provider.GetOrCreate<DocumentSettings>();
        DocumentSettings.CheckSettings();

        try
        {
            ChaptersList = _provider.GetOrCreate<ChaptersList>();
            LiteratureList = _provider.GetOrCreate<LiteratureList>();
            FormulaList = _provider.GetOrCreate<FormulaList>();
            PictureList = _provider.GetOrCreate<PictureList>();
            TableList = _provider.GetOrCreate<TableList>();
        }
        catch (Exception e)
        {
            MessageBox.Show(e.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public void SaveBeforeClose()
    {
        _provider.SaveAll();
    }
}
