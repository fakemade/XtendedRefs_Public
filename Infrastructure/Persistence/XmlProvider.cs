using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml.Serialization;
using XtendedRefs.ComAddIn.Application.Contracts;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace XtendedRefs.ComAddIn.Infrastructure.Persistence
{
    // =========================================
    // ЭЛЕМЕНТ ХРАНИЛИЩА (ОБЕРТКА)
    // =========================================

    /// <summary>
    /// Представляет собой связку: Тип данных + Десериализованный объект в памяти + Ссылка на XML-часть в документе Word.
    /// </summary>
    public class XmlProviderItem
    {
        /// <summary>
        /// Тип C# объекта (например, typeof(MySettings)).
        /// </summary>
        public Type Type { get; set; }

        /// <summary>
        /// Сам объект, восстановленный из XML (или созданный с нуля).
        /// </summary>
        public object Value { get; set; }

        /// <summary>
        /// Ссылка на физическую XML-часть (CustomXMLPart) внутри документа Word.
        /// </summary>
        public Office.CustomXMLPart Part { get; set; }
    }

    // =========================================
    // ГЛАВНОЕ ХРАНИЛИЩЕ XML
    // =========================================

    /// <summary>
    /// Класс для сохранения и загрузки данных надстройки непосредственно в тело документа Word 
    /// с использованием механизма CustomXMLParts.
    /// </summary>
    public class XmlProvider
    {
        // Словарь "известных типов". Ключ - это namespace (или ElementName), Значение - сам Type.
        private Dictionary<string, Type> _knownTypes;

        // Текущий активный документ Word, с которым мы работаем.
        private readonly Word.Document _document;

        // Кэш загруженных объектов в оперативной памяти. Ключ - Type, Значение - обертка XmlProviderItem.
        private readonly Dictionary<Type, XmlProviderItem> _store = new();

        // =========================================
        // КОНСТРУКТОР
        // =========================================

        /// <summary>
        /// Инициализирует хранилище для указанного документа.
        /// </summary>
        /// <param name="document">Документ Word, в котором будут храниться данные.</param>
        /// <param name="knownTypes">Массив типов, которые разрешено сериализовать/десериализовать.</param>
        public XmlProvider(Word.Document document, params Type[] knownTypes)
        {
            _document = document;

            // Формируем словарь известных типов. 
            // В качестве ключа используется ElementName из атрибута [XmlRoot].
            _knownTypes = knownTypes.ToDictionary(
                t => t.GetCustomAttribute<XmlRootAttribute>()?.Namespace ?? "",
                t => t
            );

            // 1. Загружаем уже существующие XML-части из документа в память.
            LoadExistingParts();
            
            // 2. Вызываем метод восстановления состояния для объектов, реализующих IRestorable.
            Restore();
        }

        /// <summary>
        /// Удаляет все XML-части
        /// </summary>
        public void Clear()
        {
            foreach (Office.CustomXMLPart part in _document.CustomXMLParts)
            {
                part.Delete();
                _document.Save();
            }
        }


        // =========================================
        // ЗАГРУЗКА СУЩЕСТВУЮЩИХ XML ЧАСТЕЙ
        // =========================================

        /// <summary>
        /// Сканирует документ Word, находит все CustomXMLPart, десериализует их в объекты C# 
        /// и помещает в локальный кэш (_store).
        /// </summary>
        private void LoadExistingParts()
        {
            foreach (Office.CustomXMLPart part in _document.CustomXMLParts)
            {
                // Проверяем, есть ли namespace этой XML-части в списке наших известных типов
                if (!_knownTypes.TryGetValue(part.NamespaceURI, out var targetType))
                {
                    continue; // Если нет - пропускаем (возможно, это XML от другой надстройки)
                }

                var serializer = new XmlSerializer(targetType);
                object obj = null;
                
                using (var reader = new StringReader(part.XML))
                {
                    obj = serializer.Deserialize(reader);

                    if (obj == null)
                    {
                        continue;
                    }

                    // Сохраняем в кэш
                    _store[targetType] = new XmlProviderItem
                    {
                        Type = targetType,
                        Value = obj,
                        Part = part
                    };
                }
            }
        }

        // =========================================
        // ПОЛУЧИТЬ ИЛИ СОЗДАТЬ
        // =========================================

        /// <summary>
        /// Возвращает объект типа T из кэша. Если его там нет, создает новый, 
        /// добавляет для него XML-часть в документ и сохраняет в кэш.
        /// </summary>
        public T GetOrCreate<T>() where T : class, new()
        {
            if (_store.TryGetValue(typeof(T), out var existing))
            {
                return (T)existing.Value;
            }

            // Создаем новый экземпляр
            var created = new T();

            // Сериализуем в XML и добавляем в документ как новую CustomXMLPart
            string xml = Serialize(created);
            var xmlPart = _document.CustomXMLParts.Add(xml);

            // Кэшируем
            _store[typeof(T)] = new XmlProviderItem
            {
                Type = typeof(T),
                Value = created,
                Part = xmlPart
            };

            return created;
        }

        // =========================================
        // ПОЛУЧИТЬ ТОЛЬКО СУЩЕСТВУЮЩИЙ
        // =========================================

        /// <summary>
        /// Возвращает объект типа T из кэша, если он был загружен. Иначе возвращает null.
        /// </summary>
        public T Get<T>() where T : class
        {
            if (_store.TryGetValue(typeof(T), out var existing))
            {
                return (T)existing.Value;
            }
            
            return null;
        }

        /// <summary>
        /// Вызывает метод Restore() у всех объектов в кэше, которые реализуют интерфейс IRestorable.
        /// Обычно используется для перерисовки UI или восстановления внутренних связей после загрузки.
        /// </summary>
        public void Restore()
        {
            foreach (XmlProviderItem item in _store.Values)
            {
                IRestorable obj = item.Value as IRestorable;
                if(obj != null)
                {
                    obj.Restore(_document);
                }
            }
        }
        
        
        // =========================================
        // СОХРАНЕНИЕ ОДНОГО ОБЪЕКТА
        // =========================================

        /// <summary>
        /// Сериализует и сохраняет изменения объекта типа T обратно в документ Word.
        /// </summary>
        public void Save<T>() where T : class
        {
            if (!_store.TryGetValue(typeof(T), out var item))
            {
                return;
            }
            SaveItem(item);
        }

        // =========================================
        // СОХРАНЕНИЕ ВСЕХ ОБЪЕКТОВ
        // =========================================

        /// <summary>
        /// Сериализует и сохраняет все объекты, находящиеся в кэше, обратно в документ Word.
        /// </summary>
        public void SaveAll()
        {
            foreach (XmlProviderItem item in _store.Values)
            {
                try
                {
                    SaveItem(item);
                }
                catch (Exception e)
                {

                    MessageBox.Show(e.Message);
                }
                
            }
        }

        // =========================================
        // ВНУТРЕННЯЯ ЛОГИКА СОХРАНЕНИЯ
        // =========================================

        /// <summary>
        /// Физически обновляет XML-часть в документе Word.
        /// </summary>
        private void SaveItem(XmlProviderItem item)
        {
            string xml = Serialize(item.Value);

            // ВАЖНЫЙ НЮАНС: 
            // API CustomXMLPart не всегда позволяет безопасно перезаписать XML напрямую.
            // Поэтому используется паттерн "удалить старое -> создать новое".
            item.Part.Delete();
            item.Part = _document.CustomXMLParts.Add(xml);
            
        }
        
        // =========================================
        // СЕРИАЛИЗАЦИЯ
        // =========================================

        /// <summary>
        /// Превращает C# объект в строку XML.
        /// </summary>
        private string Serialize(object obj)
        {
            var serializer = new XmlSerializer(obj.GetType());

            using var writer = new StringWriter();
            serializer.Serialize(writer, obj);

            return writer.ToString();
        }

        // =========================================
        // ПОЛНАЯ ОЧИСТКА (ВСЕХ ЧАСТЕЙ)
        // =========================================

        /// <summary>
        /// Удаляет АБСОЛЮТНО ВСЕ CustomXMLParts из документа.
        /// Осторожно: это может удалить XML-данные других надстроек, если они есть в документе.
        /// </summary>
        public void ClearAllCustomXmlParts()
        {
            // Снова идем с конца в начало из-за 1-базовой индексации COM
            for (int i = _document.CustomXMLParts.Count; i >= 1; i--)
            {
                try
                {
                    _document.CustomXMLParts[i].Delete();
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
                
            }
            _store.Clear();
        }

        // =========================================
        // УДАЛЕНИЕ ОДНОГО ТИПА
        // =========================================

        /// <summary>
        /// Удаляет объект типа T из кэша и удаляет соответствующую XML-часть из документа Word.
        /// </summary>
        public void Remove<T>()
        {
            if (!_store.TryGetValue(typeof(T), out var item))
            {
                return;
            }

            item.Part.Delete();
            _store.Remove(typeof(T));
        }

        // =========================================
        // ПРОВЕРКА НАЛИЧИЯ
        // =========================================

        /// <summary>
        /// Проверяет, загружен ли объект типа T в кэш.
        /// </summary>
        public bool Contains<T>()
        {
            return _store.ContainsKey(typeof(T));
        }

        // =========================================
        // ПЕРЕЧИСЛЕНИЕ
        // =========================================

        /// <summary>
        /// Возвращает коллекцию всех десериализованных объектов из кэша.
        /// </summary>
        public IEnumerable<object> AllValues =>
            _store.Values.Select(x => x.Value);
    }
}