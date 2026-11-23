using System;
using System.Collections.Generic;
using System.Linq;

namespace XtendedRefs.ComAddIn.Domain.Models.NewModels;

public static class EnumerableExtensions
{
    public static List<TResult> ToList<T, TResult>(
        this IEnumerable<T> source,
        Func<T, TResult> selector)
    {
        return source.Select(selector).ToList();
    }
}