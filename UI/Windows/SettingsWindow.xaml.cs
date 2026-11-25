using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using XtendedRefs.ComAddIn.Application.Utilities;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Documents;
using XtendedRefs.ComAddIn.Domain.State;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.UI.Windows;

public partial class SettingsWindow : Window
{
    private readonly Word.Document _document;
    private readonly RuntimeStore mObj;

    public SettingsWindow(Word.Document document)
    {
        InitializeComponent();
        _document = document;
        try
        {
            mObj = Singleton<RuntimeStore>.Instance;
        }
        catch (Exception e)
        {
            MessageBox.Show("Error: \n" + e.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        LoadSettings();
    }

    private void LoadSettings()
    {
        
        PictureSignTemplate.Text = mObj.DocumentSettings.PictureSignTemplate;
        FormulaSignTemplate.Text = mObj.DocumentSettings.FormulaSignTemplate;
        LiteratureSignTemplate.Text = mObj.DocumentSettings.LiteratureSignTemplate;
        TableSignTemplate.Text = mObj.DocumentSettings.TableSignTemplate;
        
        PictureInsertTemplate.Text = mObj.DocumentSettings.PictureInsertTemplate;
        FormulaInsertTemplate.Text = mObj.DocumentSettings.FormulaInsertTemplate;
        LiteratureInsertTemplate.Text = mObj.DocumentSettings.LiteratureInsertTemplate;
        TableInsertTemplate.Text = mObj.DocumentSettings.TableInsertTemplate;
        

        List<string> paraStyles = DocumentUtilities.GetParagraphStyleNames(_document);
        paraStyles.Sort(StringComparer.CurrentCultureIgnoreCase);
        
        LiteratureSymbolCombo.Items.Add("Нет");
        LiteratureSymbolCombo.Items.Add("Пробел");
        LiteratureSymbolCombo.Items.Add("Табуляция");
        LiteratureSymbolCombo.SelectedIndex = 0;
        
        PictureStyleCombo.Items.Add("(не задано)");
        LiteratureStyleCombo.Items.Add("(не задано)");
        TableStyleCombo.Items.Add("(не задано)");

        foreach (string style in paraStyles)
        {
            PictureStyleCombo.Items.Add(style);
            LiteratureStyleCombo.Items.Add(style);
            TableStyleCombo.Items.Add(style);
            AvailableStylesList.Items.Add(style);
        }

        // LiteratureBracketCombo.Items.Add(Brackets.GetCircleText());
        // LiteratureBracketCombo.Items.Add(Brackets.GetSquareText());
        // LiteratureBracketCombo.Items.Add(Brackets.GetCurlyText());
        // LiteratureBracketCombo.Items.Add("(не задано)");
        
        SelectInCombo(PictureStyleCombo, mObj.DocumentSettings.PictureStyle);
        SelectInCombo(LiteratureStyleCombo, mObj.DocumentSettings.LiteratureStyle);
        SelectInCombo(TableStyleCombo, mObj.DocumentSettings.TableStyle);

        TableCaptionNextLineCheck.IsChecked = string.Equals(
            mObj.DocumentSettings.TableCaptionNextLine.ToString(),
            "True",
            StringComparison.OrdinalIgnoreCase);

        string litSymbol = mObj.DocumentSettings.LiteratureSymbol;
        LiteratureSymbolCombo.SelectedIndex = litSymbol switch
        {
            "SPACE" => 1,
            "TAB" => 2,
            _ => 0
        };
        for (int level = 0; level < mObj.DocumentSettings.LevelStyles.Count; level++)
        {
            try
            {
                string levelStyle = mObj.DocumentSettings.LevelStyles[level].StyleName;
                SelectedStylesList.Items.Add(levelStyle);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private static void SelectInCombo(System.Windows.Controls.ComboBox combo, string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            combo.SelectedIndex = 0;
            return;
        }

        for (int i = 0; i < combo.Items.Count; i++)
        {
            if (string.Equals(combo.Items[i].ToString(), value, StringComparison.OrdinalIgnoreCase))
            {
                combo.SelectedIndex = i;
                return;
            }
        }

        combo.SelectedIndex = 0;
    }

    private void AddLevelStyle_Click(object sender, RoutedEventArgs e)
    {
        if (AvailableStylesList.SelectedItem is not string style)
        {
            return;
        }

        if (SelectedStylesList.Items.Count >= 3)
        {
            MessageBox.Show(this, "Максимум 3 уровня!", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (SelectedStylesList.Items.Contains(style))
        {
            MessageBox.Show(this, "Этот стиль уже добавлен!", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        SelectedStylesList.Items.Add(style);
    }

    private void RemoveLevelStyle_Click(object sender, RoutedEventArgs e)
    {
        int idx = SelectedStylesList.SelectedIndex;
        if (idx < 0)
        {
            return;
        }

        SelectedStylesList.Items.RemoveAt(idx);
    }

    private void LevelStyleUp_Click(object sender, RoutedEventArgs e)
    {
        int idx = SelectedStylesList.SelectedIndex;
        if (idx <= 0)
        {
            return;
        }

        object item = SelectedStylesList.Items[idx];
        SelectedStylesList.Items.RemoveAt(idx);
        SelectedStylesList.Items.Insert(idx - 1, item);
        SelectedStylesList.SelectedIndex = idx - 1;
    }

    private void LevelStyleDown_Click(object sender, RoutedEventArgs e)
    {
        int idx = SelectedStylesList.SelectedIndex;
        if (idx < 0 || idx >= SelectedStylesList.Items.Count - 1)
        {
            return;
        }

        object item = SelectedStylesList.Items[idx];
        SelectedStylesList.Items.RemoveAt(idx);
        SelectedStylesList.Items.Insert(idx + 1, item);
        SelectedStylesList.SelectedIndex = idx + 1;
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        if (!PictureSignTemplate.Text.Contains("$"))
        {
            MessageBox.Show(this, "Шаблон подписи рисунка должен содержать $!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (!PictureInsertTemplate.Text.Contains("$"))
        {
            MessageBox.Show(this, "Шаблон вставки рисунка должен содержать $!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!FormulaSignTemplate.Text.Contains("$"))
        {
            MessageBox.Show(this, "Шаблон подписи формулы должен содержать $!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (!FormulaInsertTemplate.Text.Contains("$"))
        {
            MessageBox.Show(this, "Шаблон вставки формулы должен содержать $!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!LiteratureSignTemplate.Text.Contains("$"))
        {
            MessageBox.Show(this, "Шаблон подписи литературы должен содержать $!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (!LiteratureInsertTemplate.Text.Contains("$"))
        {
            MessageBox.Show(this, "Шаблон вставки литературы должен содержать $!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!TableSignTemplate.Text.Contains("$"))
        {
            MessageBox.Show(this, "Шаблон подписи таблицы должен содержать $!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (!TableInsertTemplate.Text.Contains("$"))
        {
            MessageBox.Show(this, "Шаблон вставки таблицы должен содержать $!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        
        //Check equality $ symbols
        // if (DocumentUtilities.CountSymbols('$', LiteratureSignTemplate.Text) !=
        //     DocumentUtilities.CountSymbols('$', LiteratureInsertTemplate.Text))
        // {
        //     MessageBox.Show(this, "Количество $ для подписи и вставки литературы должно совпадать!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
        //     return;
        // }
        if (DocumentUtilities.CountSymbols('$', FormulaSignTemplate.Text) !=
            DocumentUtilities.CountSymbols('$', FormulaInsertTemplate.Text))
        {
            MessageBox.Show(this, "Количество $ для подписи и вставки формулы должно совпадать!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (DocumentUtilities.CountSymbols('$', TableSignTemplate.Text) !=
            DocumentUtilities.CountSymbols('$', TableInsertTemplate.Text))
        {
            MessageBox.Show(this, "Количество $ для подписи и вставки таблицы должно совпадать!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (DocumentUtilities.CountSymbols('$', PictureSignTemplate.Text) !=
            DocumentUtilities.CountSymbols('$', PictureInsertTemplate.Text))
        {
            MessageBox.Show(this, "Количество $ для подписи и вставки рисунка должно совпадать!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        
        try
        {
            SaveSettings();
            MessageBox.Show(this, "Настройки сохранены.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Ошибка при сохранении: " + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void SelectedStylesList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        try
        {
            ListBox listBox = sender as ListBox;

            if (listBox != null && listBox.SelectedItem != null)
            {
                LevelStyleItem item =
                    mObj.DocumentSettings.LevelStyles.FirstOrDefault(x => x.StyleLevel == listBox.SelectedIndex + 1);
                MessageBox.Show($"{item.StyleLevel},{item.StyleName},{item.StyleLevel.ToString()}");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.ToString());
        }
    }
    
   

    private void SaveSettings()
    {
        
        mObj.DocumentSettings.PictureSignTemplate = PictureSignTemplate.Text;
        mObj.DocumentSettings.FormulaSignTemplate = FormulaSignTemplate.Text;
        mObj.DocumentSettings.LiteratureSignTemplate = LiteratureSignTemplate.Text;
        mObj.DocumentSettings.TableSignTemplate = TableSignTemplate.Text;
        
        mObj.DocumentSettings.PictureInsertTemplate = PictureInsertTemplate.Text;
        mObj.DocumentSettings.FormulaInsertTemplate = FormulaInsertTemplate.Text;
        mObj.DocumentSettings.LiteratureInsertTemplate = LiteratureInsertTemplate.Text;
        mObj.DocumentSettings.TableInsertTemplate = TableInsertTemplate.Text;

        string figStyle = PictureStyleCombo.SelectedIndex > 0 ? PictureStyleCombo.SelectedItem.ToString() : string.Empty;
        mObj.DocumentSettings.PictureStyle = figStyle;

        string litStyle = LiteratureStyleCombo.SelectedIndex > 0 ? LiteratureStyleCombo.SelectedItem.ToString() : string.Empty;
        mObj.DocumentSettings.LiteratureStyle = litStyle;

        string tableStyle = TableStyleCombo.SelectedIndex > 0 ? TableStyleCombo.SelectedItem.ToString() : string.Empty;
        mObj.DocumentSettings.TableStyle = tableStyle;

        mObj.DocumentSettings.TableCaptionNextLine = TableCaptionNextLineCheck.IsChecked == true;
        

        string litSymbol = LiteratureSymbolCombo.SelectedIndex switch
        {
            1 => "SPACE",
            2 => "TAB",
            _ => string.Empty
        };
        mObj.DocumentSettings.LiteratureSymbol = litSymbol;

        // Write level styles atomically: build the new values first, then write
        // each slot in a single pass so a crash mid-save doesn't leave an empty list.
        mObj.DocumentSettings.LevelStyles.Clear();
        for (int i = 0; i < SelectedStylesList.Items.Count; i++)
        {
            mObj.DocumentSettings.LevelStyles.Add(new LevelStyleItem { StyleLevel = i + 1, StyleName = SelectedStylesList.Items[i].ToString() });
        }
    }
}
