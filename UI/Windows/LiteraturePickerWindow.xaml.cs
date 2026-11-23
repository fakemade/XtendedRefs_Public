using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using XtendedRefs.ComAddIn.Application.Composition;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Models.NewModels;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.UI.Windows;

public partial class LiteraturePickerWindow : Window
{
    private readonly Word.Document _document;
    private List<LiteratureInfo> _items;
    
    public LiteratureInfo SelectedBookmarkName { get; private set; }
    
    public LiteraturePickerWindow(Word.Document document)
    {
        InitializeComponent();
        _document = document;

        _items = Singleton<ServiceLocator>.Instance.LiteraturePickerService.BuildSourceList();
        SourcesList.ItemsSource = _items;

        if (_items.Count > 0)
        {
            SourcesList.SelectedIndex = 0;
        }
    }

    

    private void SourcesList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (SourcesList.SelectedItem is LiteratureInfo item)
        {
            SourceTextBox.Text = item.LiteratureText;
            SourceTextBox.IsEnabled = true;
            // SaveButton.IsEnabled = true;
            OkButton.IsEnabled = true;

        }
        else
        {
            SourceTextBox.Text = string.Empty;
            SourceTextBox.IsEnabled = false;
            SaveButton.IsEnabled = false;
            OkButton.IsEnabled = false;
            MoveUpButton.IsEnabled = false;
            MoveDownButton.IsEnabled = false;
            DeleteButton.IsEnabled = false;
        }
    }

    private void UpdateMoveButtons(int idx)
    {
        MoveUpButton.IsEnabled = idx > 0;
        MoveDownButton.IsEnabled = idx < _items.Count - 1;
    }

    private void MoveUpButton_Click(object sender, RoutedEventArgs e)
    {
        return;
        int idx = SourcesList.SelectedIndex;
        if (idx <= 0)
        {
            return;
        }

        try
        {

            _document.Fields.Update();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Ошибка при перемещении: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        RebuildItems(idx - 1);
    }

    private void MoveDownButton_Click(object sender, RoutedEventArgs e)
    {
        return;
        int idx = SourcesList.SelectedIndex;
        if (idx < 0 || idx >= _items.Count - 1)
        {
            return;
        }

        try
        {

            _document.Fields.Update();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Ошибка при перемещении: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        RebuildItems(idx + 1);
    }

    private void DeleteButton_Click(object sender, RoutedEventArgs e)
    {
        if (SourcesList.SelectedItem is not LiteratureInfo item)
        {
            return;
        }

        MessageBoxResult answer = MessageBox.Show(
            this,
            $"Удалить источник '{item.DisplayText}'? Ссылки в тексте будут заменены заглушкой.",
            "Подтверждение",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (answer != MessageBoxResult.Yes)
        {
            return;
        }

        return;
        int prevIdx = SourcesList.SelectedIndex;

        try
        {
            //_onDelete(item.LiteratureObject, true);
            _items.Remove(item);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Ошибка при удалении: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        int nextIdx = Math.Min(prevIdx, _items.Count - 2);
        RebuildItems(nextIdx);
    }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        if (SourcesList.SelectedItem is not LiteratureInfo item)
        {
            return;
        }

        int idx = SourcesList.SelectedIndex;

        try
        {
            SaveLiteratureText(item, SourceTextBox.Text);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Ошибка при сохранении: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        RebuildItems(idx);
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

    private void SourcesList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        Accept();
    }

    private void Accept()
    {
        if (SourcesList.SelectedItem is not LiteratureInfo item)
        {
            MessageBox.Show(this, "Выберите источник!", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        SelectedBookmarkName = item;
        DialogResult = true;
        Close();
    }

    

    private void SaveLiteratureText(LiteratureInfo litObj, string text)
    {
        litObj.LiteratureText = text;
    }

    private void RebuildItems(int selectIdx)
    {
        _items = Singleton<ServiceLocator>.Instance.LiteraturePickerService.BuildSourceList();
        SourcesList.ItemsSource = null;
        SourcesList.ItemsSource = _items;

        if (_items.Count > 0)
        {
            int clampedIdx = Math.Max(0, Math.Min(selectIdx, _items.Count - 1));
            SourcesList.SelectedIndex = clampedIdx;
            SourcesList.ScrollIntoView(SourcesList.SelectedItem);
        }
    }
}
