using System.Collections.Generic;
using System.Xml.Serialization;

namespace XtendedRefs.ComAddIn.Domain.Models.NewModels;

public abstract class ListXml<T>
{
    // Храним данные здесь, но не сериализуем в базовом классе
    [XmlIgnore]
    protected List<T> _internalItems = new List<T>();

    public List<T> GetItems() => _internalItems;
    public void SetItems(List<T> value) => _internalItems = value;
}