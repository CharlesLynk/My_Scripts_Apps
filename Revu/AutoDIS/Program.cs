using System;
using System.Threading;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using WinRT;

namespace MbrohDrawingIndexCreator;

public static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        // The program starts itself as a render helper (draws pages on request) or runs the helper self-test before any window exists.
        if (args.Length > 0 && args[0] == RenderHelper.Switch)
        {
            Environment.ExitCode = RenderHelper.RunHelper();
            return;
        }
        if (args.Length > 0 && args[0] == RenderHelper.SelfTestSwitch)
        {
            Environment.ExitCode = RenderHelper.RunSelfTest(args);
            return;
        }
        // Tesseract is built with OpenMP and would start one thread per processor in every engine. The app already runs one engine per worker, so each engine stays single-threaded (set before the native library loads).
        Environment.SetEnvironmentVariable("OMP_THREAD_LIMIT", "1");
        ComWrappersSupport.InitializeComWrappers();
        Application.Start(p =>
        {
            SynchronizationContext.SetSynchronizationContext(new DispatcherQueueSynchronizationContext(DispatcherQueue.GetForCurrentThread()));
            _ = new App();
        });
    }
}
