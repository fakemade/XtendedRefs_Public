#nullable enable
using System;
using System.Diagnostics;
using System.Linq;

namespace XtendedRefs.ComAddIn.Domain.Abstractions;

public static class Singleton<T> where T : class
{
    private static Lazy<T>? _instance;

    public static void Initialize(Func<T> factory)
    {
        if (_instance != null) throw new Exception("Singleton already initialized");
        _instance = new Lazy<T>(factory);
    }

    public static void SetInstance(T instance)
    {
        _instance = new Lazy<T>(() => instance);
    }
    
    public static T Instance
    {
        get
        {
            if (_instance == null)
            {
 
                var trace = new StackTrace();
                // 0 = текущий метод (Instance getter)
                // 1 = место, откуда вызвали Instance
                var frame = trace.GetFrame(1) ?? trace.GetFrame(0);
                
                var frames = trace.GetFrames();

                var caller = frames?.FirstOrDefault(f =>
                    f.GetMethod()?.DeclaringType != typeof(Singleton<T>));

                var method = caller?.GetMethod();
                var file = frame?.GetFileName();
                var line = frame?.GetFileLineNumber();

                throw new Exception(
                    $"Singleton<{typeof(T).Name}> not initialized\n" +
                    $"Called from: {method?.DeclaringType?.FullName}.{method?.Name}\n" +
                    $"At: {file}:{line}");
            }
            
            return _instance.Value;
        }
    }
}
