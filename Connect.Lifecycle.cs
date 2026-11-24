using System;
using System.IO;
using System.Reflection;
using System.Windows;
using Extensibility;
using XtendedRefs.ComAddIn.Application.Composition;
using XtendedRefs.ComAddIn.Domain.Abstractions;
using Word = Microsoft.Office.Interop.Word;
using MessageBox = System.Windows.MessageBox;
using WordApplicationEventHandler = XtendedRefs.ComAddIn.Application.Events.WordApplicationEventHandler;

namespace XtendedRefs.ComAddIn;

public sealed partial class Connect
{
    static Connect()
    {

        Assembly self = Assembly.GetExecutingAssembly();
        string location = self.Location;
        if (string.IsNullOrEmpty(location))
        {
            location = AppContext.BaseDirectory;
        }

        string addinDir = string.IsNullOrEmpty(location) ? null : Path.GetDirectoryName(location);

        AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
        {
            if (addinDir == null) return null;
            try
            {
                string name = new AssemblyName(args.Name).Name;
                string path = Path.Combine(addinDir, name + ".dll");
                if (File.Exists(path))
                {
                    return Assembly.LoadFrom(path);
                }
            }
            catch (Exception ex)
            {
            }
            
            return null;
        };
        
    }

    public Connect()
    {
        try
        {
            Singleton<ServiceLocator>.Initialize(() => new ServiceLocator());
            Singleton<CommandsLocator>.Initialize(() => new CommandsLocator());
            _commands = Singleton<CommandsLocator>.Instance;
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            throw;
        }
    }

    public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
    {
        try
        {

            AddInContext.Application = application as Word.Application;
            AddInContext.ComAddInInstance = addInInst;
            AddInContext.ConnectMode = connectMode;


            if (AddInContext.Application != null)
            {
                _wordEventHandler = new WordApplicationEventHandler(_ribbon);
                _wordEventHandler.Subscribe();
            }

            _instance = this;

        }
        catch (Exception ex)
        {

            MessageBox.Show(ex.ToString());
            throw;
        }
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
    {

        try
        {
            _wordEventHandler?.Unsubscribe();
            _wordEventHandler = null;
            _ribbon = null;
            AddInContext.Application = null;
            AddInContext.ComAddInInstance = null;

        }
        catch (Exception ex)
        {

            throw;
        }
    }

    public void OnAddInsUpdate(ref Array custom)
    {

    }

    public void OnStartupComplete(ref Array custom)
    {
    }

    public void OnBeginShutdown(ref Array custom)
    {
    }
}
