using XtendedRefs.ComAddIn.Application.Commands.NewCommands;
using LiteratureCommands = XtendedRefs.ComAddIn.Application.Commands.NewCommands.LiteratureCommands;
using FormulaCommands = XtendedRefs.ComAddIn.Application.Commands.NewCommands.FormulaCommands;
using PictureCommands = XtendedRefs.ComAddIn.Application.Commands.NewCommands.PictureCommands;
using TableCommands = XtendedRefs.ComAddIn.Application.Commands.NewCommands.TableCommands;

namespace XtendedRefs.ComAddIn.Application.Composition;

public sealed class CommandsLocator
{
    public LiteratureCommands LiteratureCommands { get; }

    public FormulaCommands FormulaCommands { get; }

    public PictureCommands PictureCommands { get; }

    public ChapterCommands ChapterCommands { get; }
    
    public TableCommands TableCommands { get; }

    public CommandsLocator()
    {
        ChapterCommands = new ChapterCommands();
        LiteratureCommands = new LiteratureCommands();
        FormulaCommands = new FormulaCommands();
        PictureCommands =  new PictureCommands();
        TableCommands = new TableCommands();
    }

    public void SetVariables()
    {
    }
}
