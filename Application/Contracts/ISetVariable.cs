
using XtendedRefs.ComAddIn.Infrastructure.Persistence;

namespace XtendedRefs.ComAddIn.Application.Contracts;

public interface ISetVariable
{
    public void SetVar(XmlProvider provider);
}