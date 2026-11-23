using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using XtendedRefs.ComAddIn.Application.Commands;
using XtendedRefs.ComAddIn.Application.Composition;
using XtendedRefs.ComAddIn.Application.Utilities;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.UI.Windows;

internal partial class ImagePickerWindow : Window
{   
    
    private IReadOnlyList<DocumentAnchor> _items;
    public ImagePickerWindow(Word.Document doc, string title, LinkType type)
    {
        InitializeComponent();
        Title = title;
        switch (type)
        {
            case LinkType.Formula:
                _items = Singleton<ServiceLocator>.Instance.FormulaPickerService.BuildSourceList();
                ItemsList.ItemsSource = _items;
                break;
            case LinkType.Picture:
                _items = Singleton<ServiceLocator>.Instance.PicturePickerService.BuildSourceList();
                ItemsList.ItemsSource = _items;
                break;
            case LinkType.Table:
                _items = Singleton<ServiceLocator>.Instance.TablePickerService.BuildSourceList();
                foreach (TableInfo documentAnchor in _items)
                {
                    try
                    {
                        CommonCommands.ShowInfo(documentAnchor.ToString());
                    }
                    catch (Exception ex)
                    {
                        CommonCommands.ShowError(ex.ToString());
                    }
                }
                ItemsList.ItemsSource = _items;
                break;
            default:
                _items = new List<DocumentAnchor>();
                ItemsList.ItemsSource = _items;
                break;
        }
        
        if (ItemsList.Items.Count > 0)
        {
            ItemsList.SelectedIndex = 0;
        }
    }

    public DocumentAnchor SelectedBookmarkName { get; private set; } = null;

    private void ItemsList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (ItemsList.SelectedItem is DocumentAnchor item)
        {
            //LoadPreview(item.ImagePath);
            OkButton.IsEnabled = true;
        }
        else
        {
            NoPreviewText.Visibility = Visibility.Visible;
            OkButton.IsEnabled = false;
        }
    }

    private void LoadPreview(string imagePath)
    {
        BitmapSource source = WpfImageLoader.TryLoad(imagePath);
        if (source != null)
        {
            PreviewImage.Source = source;
            NoPreviewText.Visibility = Visibility.Collapsed;
        }
        else
        {
            PreviewImage.Source = null;
            NoPreviewText.Visibility = Visibility.Visible;
        }
    }

    private void ItemsList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        Accept();
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        Accept();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void Accept()
    {
        if (ItemsList.SelectedItem is not DocumentAnchor item)
        {
            CommonCommands.ShowWarning("Выберите элемент!", "Предупреждение");
            return;
        }
    
        SelectedBookmarkName = item;
        DialogResult = true;
        Close();
    }

}
