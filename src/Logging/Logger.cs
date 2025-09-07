using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop; // IVsOutputWindow, IVsOutputWindowPane
using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Shapes;

namespace NF.VSTools
{
    internal static class Log
    {
        // Lazy-initialized VS Output pane (nullable until InitAsync runs)
        private static IVsOutputWindowPane? _pane;

        // Stable GUID for our pane (any GUID you keep constant)
        private static readonly Guid PaneGuid = new Guid("97C3DF24-2B6E-4B8A-B37A-6C8B5AD8C3A0");

        /// Initialize once on the UI thread during package init.
        public static async Task InitAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var outWin = await package.GetServiceAsync(typeof(SVsOutputWindow)) as IVsOutputWindow;
            if (outWin == null)
                return;

            // Can't pass a static readonly field by ref – copy to a local first.
            Guid g = PaneGuid;

            // Create or reuse our pane
            outWin.CreatePane(ref g, "NF.VSTools", fInitVisible: 1, fClearWithSolution: 1);
            outWin.GetPane(ref g, out _pane);
        }

        /// Thread-safe write from any thread. We intentionally suppress VSTHRD010
        /// because OutputStringThreadSafe is explicitly designed for non-UI usage.
        [System.Diagnostics.CodeAnalysis.SuppressMessage(
            "Microsoft.VisualStudio.Threading.Analyzers",
            "VSTHRD010",
            Justification = "IVsOutputWindowPane.OutputStringThreadSafe is safe from any thread.")]
        public static void Post(string message)
        {
            string line = $"NF.VSTools: [{DateTime.Now:HH:mm:ss}] {message}";
            try { _pane?.OutputStringThreadSafe(line + "\r\n"); } catch { }
            Debug.WriteLine(line); // shows up in the OUTER VS (Output → Debug)
        }

        /// Synchronous ActivityLog write (fast; fine for infrequent events).
        public static void Sync(string message)
        {
            try { ActivityLog.LogInformation("NF.VSTools", message); } catch { }
            Debug.WriteLine($"NF.VSTools [SYNC] {message}");
        }
    }
}