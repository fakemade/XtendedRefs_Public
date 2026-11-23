using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using XtendedRefs.ComAddIn.Application.Composition;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using Word = Microsoft.Office.Interop.Word;
namespace XtendedRefs.ComAddIn.UI.Windows;

public partial class ChapterPickerWindow : Window
{
    private readonly Word.Document _document;
    public ChapterInfo selectedChapter;
    private readonly List<ChapterInfo> _items;

    
    public ChapterPickerWindow(
        Word.Document document
        )
    {
        InitializeComponent();
        _document = document;
        _items = Singleton<ServiceLocator>.Instance.ChapterPickerService.BuildSourceList();
        SourcesList.ItemsSource = _items;
    }

    
    private void SourcesList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (SourcesList.SelectedItem is ChapterInfo item)
        {
            OkButton.IsEnabled = true;
        }
        else
        {
            OkButton.IsEnabled = false;
        }
    }

    
    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void SourcesList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        Accept();
    }

    private void Accept()
    {
        if (SourcesList.SelectedItem is ChapterInfo ch)
            selectedChapter = ch;
        else
            return;
        DialogResult = true;
        Close();
    }


    private void OkButton_OnClick(object sender, RoutedEventArgs e)
    {
        Accept();
    }
}
