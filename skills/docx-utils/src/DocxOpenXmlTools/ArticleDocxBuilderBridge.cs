using System.Reflection;
using System.Runtime.ExceptionServices;

internal static class ArticleDocxBuilderBridge
{
    public static int Run(string[] args)
    {
        var assembly = LoadAssembly();
        var entryPoint = assembly.EntryPoint ?? throw new InvalidOperationException("ArticleDocxBuilder entry point not found.");

        try
        {
            var parameters = entryPoint.GetParameters();
            object? result = parameters.Length == 1
                ? entryPoint.Invoke(null, [args])
                : entryPoint.Invoke(null, []);

            return result switch
            {
                int value => value,
                null => 0,
                _ => 0
            };
        }
        catch (TargetInvocationException ex) when (ex.InnerException is not null)
        {
            ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
            throw;
        }
    }

    private static Assembly LoadAssembly()
    {
        try
        {
            return Assembly.Load("ArticleDocxBuilder");
        }
        catch
        {
            var candidate = Path.Combine(AppContext.BaseDirectory, "ArticleDocxBuilder.dll");
            if (File.Exists(candidate))
            {
                return Assembly.LoadFrom(candidate);
            }

            throw;
        }
    }
}
