using System;
using System.Drawing;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Xml.Serialization;
using XtendedRefs.ComAddIn.Application.Contracts;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using XtendedRefs.ComAddIn.Domain.Documents;
using XtendedRefs.ComAddIn.Domain.State;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Domain.Models.NewModels;

[XmlRoot(ElementName = "FormulaList", Namespace = "XtendedRefs.ComAddIn.Formula")]
public class FormulaList : IRestorable
{
    public FormulaList()
    {
    }

    [XmlElement("Formula", Namespace = "XtendedRefs.ComAddIn.Formula")]
    public List<FormulaInfo> Formulas { get; set; } = new List<FormulaInfo>();

    public void Restore(Word.Document document)
    {
        foreach (var formulaInfo in Formulas)
        {
            formulaInfo.Restore();
        }
    }
}

public class FormulaInfo : DocumentAnchor
{
    
    private string _preview;

    public string Preview
    {
        get => _preview;
        set => _preview = value;
    }
    
    
    private Bitmap _picture;
    [XmlIgnore]
    public Bitmap Picture
    {
        get => _picture;
        set => _picture = value;
    }
    
    public FormulaInfo()
    {
    }

    public FormulaInfo(List<string> numberParts)
    {
        NumberParts.AddRange(numberParts);
        /*
        _picture = GetFormulaPreviewViaHtml(shape);
        _preview = GetPreviewAsBase64(_picture);
        */
    }
    
    public override string GetNumber(Word.Range range)
    {
        return string.Join(".", Number(range));
    }
    
    //--------------
    [XmlArray("Part")]
    public List<string> NumberParts { get; set; } = new List<string>();
    
    public List<string> Number(Word.Range range)
    {
        List<string> nums = range.Fields.Cast<Word.Field>()
            .Select(f => 
            {
                // Пишем полноценную логику с ветвлением прямо внутри LINQ
                if (f.Type == Word.WdFieldType.wdFieldSequence)
                {
                    return f.Result.Text;
                }

                if (f.Type == Word.WdFieldType.wdFieldHyperlink)
                {
                    if (f.Result.Hyperlinks.Count > 0)
                    {
                        Word.Hyperlink link = f.Result.Hyperlinks[1];
                        string visibleText = link.TextToDisplay;

                        // Теперь вы можете передать объект link в свой метод обработки
                        return visibleText;
                    }
                }

                return "Пропущено"; // Возвращаем строковый статус для остальных
            })
            .ToList(); // Запускает итерацию по коллекции
        return nums;
    }
    
    [XmlIgnore] 
    public sealed override string DisplayText
    {
        get => NumberParts?.Count > 0 ? "Формула " + string.Join(".", NumberParts) : "Нет данных";
        set => throw new NotImplementedException();
    }


    [XmlIgnore]
    public string InsertText
    {
        get
        {
            RuntimeStore rs = Singleton<RuntimeStore>.Instance;
            DocumentSettings documentSettings = rs.DocumentSettings;
            
            int index = 0;
            string result = Regex.Replace(documentSettings.FormulaInsertTemplate, @"\$", m => 
            {
                if (index < NumberParts.Count)
                {
                    return NumberParts[index++];
                }
                return m.Value;
            });

            return result;
        }
    }
    
    
    /// <summary>
    /// Метод получению превью для отображения
    /// </summary>
    /// <param name="originalShape">Объект, который содержит формулу (должна быть выделена формула MathType)</param>
    /// <returns>Изображение типа <see cref="Bitmap"/></returns>
    public static Bitmap GetFormulaPreviewViaHtml(Word.InlineShape originalShape)
    {
        Word.Application wordApp = originalShape.Application;
        Word.Document tempDoc = null;
        string tempFolder = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        string tempHtmlPath = Path.Combine(tempFolder, "temp.html");

        try
        {
            // 1. Создаем директорию для временных файлов
            Directory.CreateDirectory(tempFolder);

            // 2. Создаем временный скрытый документ
            object visible = false;
            tempDoc = wordApp.Documents.Add(Visible: ref visible);

            // 3. Копируем формулу во временный документ
            originalShape.Range.Copy();
            tempDoc.Content.Paste();

            // 4. Сохраняем документ как HTML (Word сам отрендерит формулу в картинку)
            object fileFormat = Word.WdSaveFormat.wdFormatHTML;
            tempDoc.SaveAs2(FileName: tempHtmlPath, FileFormat: ref fileFormat);
            tempDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            tempDoc = null;

            // 5. Word создает папку "temp.files" рядом с temp.html и складывает туда картинки
            string filesFolder = Path.Combine(tempFolder, "temp.files");
            if (Directory.Exists(filesFolder))
            {
                // Ищем файлы изображений (.wmf, .png, .gif) в этой папке
                var imageFile = Directory.GetFiles(filesFolder)
                    .FirstOrDefault(f => f.EndsWith(".wmf", StringComparison.OrdinalIgnoreCase) ||
                                         f.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ||
                                         f.EndsWith(".gif", StringComparison.OrdinalIgnoreCase));

                if (imageFile != null)
                {
                    // Читаем файл в массив байт, чтобы не блокировать его на диске
                    byte[] fileBytes = File.ReadAllBytes(imageFile);
                    using (MemoryStream ms = new MemoryStream(fileBytes))
                    {
                        // Возвращаем чистый растровый Bitmap формулы
                        return new Bitmap(ms);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine("Ошибка HTML-экспорта: " + ex.Message);
        }
        finally
        {
            // Гарантированно закрываем временный документ, если он завис
            if (tempDoc != null)
            {
                try
                {
                    tempDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                catch
                {
                }
            }

            // Чистим за собой временную папку
            try
            {
                if (Directory.Exists(tempFolder))
                    Directory.Delete(tempFolder, true);
            }
            catch
            {
            }
        }

        return null;
    }

    /// <summary>
    /// Преобразование превью в строку формата Base64
    /// </summary>
    /// <returns>Строка формата <see cref="string"/> <c>Base64</c></returns>
    public string GetPreviewAsBase64(Bitmap p)
    {
        if (Preview == null) return string.Empty;

        using (MemoryStream ms = new MemoryStream())
        {
            // Сохраняем в PNG для сохранения прозрачности и качества
            p.Save(ms, ImageFormat.Png);
            byte[] imageBytes = ms.ToArray();
            return Convert.ToBase64String(imageBytes);
        }
    }
    
    
    /// <summary>
    /// Метод, который преобразует существующую строку в формате <c>Base64</c> в изображение типа <see cref="Bitmap"/>
    /// </summary>
    /// <param name="base64String">Строка с данными изображения</param>
    /// <returns>Возвращает изображение</returns>
    public Bitmap LoadPreviewFromBase64(string base64String)
    {
        if (string.IsNullOrEmpty(base64String)) return null;

        byte[] imageBytes = Convert.FromBase64String(base64String);
        using (MemoryStream ms = new MemoryStream(imageBytes))
        {
            Bitmap b = new Bitmap(ms);
            return b;
        }
    }

    public void Restore()
    {
        //_picture = LoadPreviewFromBase64(_preview);
    }

    
}