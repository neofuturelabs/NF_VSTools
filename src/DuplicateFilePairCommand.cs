using Community.VisualStudio.Toolkit;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text;  // for Encoding

namespace NF.VSTools
{
    // Note: CommandAttribute signature is (string commandSetGuid, int commandId) — keep your order.
    [Command(PackageGuids.guidNF_VSToolsPackageCmdSetString, PackageIds.cmdDuplicateFilePairCommand)]
    internal sealed class DuplicateFilePairCommand : BaseCommand<DuplicateFilePairCommand>
    {
        private static string? s_lastBQSString;

        private static readonly string[] HeaderExts = { ".h", ".hpp" };
        private static readonly string[] SourceExts = { ".cpp", ".cc", ".cxx" };

        private static bool IsHeader(string p) => HeaderExts.Any(e => p.EndsWith(e, StringComparison.OrdinalIgnoreCase));
        private static bool IsSource(string p) => SourceExts.Any(e => p.EndsWith(e, StringComparison.OrdinalIgnoreCase));
        private static string GetBaseName(string p) => Path.GetFileNameWithoutExtension(p) ?? "";

        // On UI thread
        protected override void BeforeQueryStatus(EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var (_, pivotPath) = GetFirstSelectedProjectItem();
            bool isCxx = pivotPath != null && (IsHeader(pivotPath) || IsSource(pivotPath));

            Command.Visible = isCxx;
            Command.Enabled = isCxx;

            string sig = $"vis={Command.Visible} enabled={Command.Enabled}; pivot={pivotPath ?? "<none>"}";
            if (!string.Equals(sig, s_lastBQSString, StringComparison.Ordinal))
            {
                s_lastBQSString = sig;
                Log.Post($"BeforeQueryStatus: " + sig);
            }
        }

        // -------------- Execute: full logic (single- or multi-select → first item) --------------
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var (pivotItem, pivotPath) = GetFirstSelectedProjectItem();
            if (pivotPath == null || !(IsHeader(pivotPath) || IsSource(pivotPath)))
            {
                await VS.StatusBar.ShowMessageAsync("Select a C++ header/source file.");
                return;
            }

            // Resolve the complementary file: same dir first, then search the containing project
            if (!TryResolvePair(pivotPath, out string headerPath, out string sourcePath))
            {
                await VS.StatusBar.ShowMessageAsync("Could not find matching header/source pair.");
                return;
            }

            string dir = Path.GetDirectoryName(headerPath)!;
            string oldBase = GetBaseName(headerPath);

            // Respect original extensions
            var headerExt = Path.GetExtension(headerPath);
            var sourceExt = Path.GetExtension(sourcePath);

            // UI dialog: note we pass dir/exts so it can preview full paths and conflicts
            using var dlg = new ReplaceDialog(headerPath, sourcePath, dir, headerExt, sourceExt, oldBase);
            if (dlg.ShowDialog() != DialogResult.OK)
                return;

            string find = dlg.FindText;
            string replace = dlg.ReplaceText;
            if (string.IsNullOrEmpty(find))
            {
                await VS.StatusBar.ShowMessageAsync("Find text is empty; cancelled.");
                return;
            }

            string newBase = dlg.PreviewBase;
            string newHeader = dlg.NewHeaderPath;
            string newSource = dlg.NewSourcePath;

            // double-check conflicts (defense-in-depth)
            if (File.Exists(newHeader) || File.Exists(newSource))
            {
                await VS.StatusBar.ShowMessageAsync("Target files already exist; aborting.");
                return;
            }

            // Copy
            File.Copy(headerPath, newHeader, overwrite: false);
            File.Copy(sourcePath, newSource, overwrite: false);

            // ---- preserve original encoding (incl. BOM) asynchronously ----
            static async Task<(string text, Encoding enc)> ReadWithEncodingAsync(string path)
            {
                using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true);
                using var sr = new StreamReader(fs, detectEncodingFromByteOrderMarks: true);
                string txt = await sr.ReadToEndAsync().ConfigureAwait(false);
                return (txt, sr.CurrentEncoding);
            }

            static async Task WriteWithEncodingAsync(string path, string text, Encoding enc)
            {
                using var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true);
                using var sw = new StreamWriter(fs, enc);
                await sw.WriteAsync(text).ConfigureAwait(false);
            }

            // Rewrite NEW header (preserve encoding + dominant EOL)
            var (hOrig, hEnc) = await ReadWithEncodingAsync(newHeader);
            var headerEol = DetectDominantEol(hOrig);
            var h = hOrig;
            h = h.Replace($"{oldBase}.generated.h", $"{newBase}.generated.h");
            h = Regex.Replace(h, Regex.Escape(find), replace);
            h = NormalizeEol(h, headerEol);
            await WriteWithEncodingAsync(newHeader, h, hEnc);

            // Rewrite NEW source (preserve encoding + dominant EOL)
            var (cppOrig, cEnc) = await ReadWithEncodingAsync(newSource);
            var sourceEol = DetectDominantEol(cppOrig);
            var cpp = cppOrig;
            // robust include update (quotes only; UE style)
            cpp = Regex.Replace(
                cpp,
                @"(?m)^\s*#\s*include\s*""\s*" + Regex.Escape(oldBase) + @"\s*\.h\s*""\s*$",
                $"#include \"{newBase}.h\""
            );
            cpp = Regex.Replace(cpp, Regex.Escape(find), replace);
            cpp = NormalizeEol(cpp, sourceEol);
            await WriteWithEncodingAsync(newSource, cpp, cEnc);

            // Add to Solution Explorer under the same filter/folder as the pivot
            bool added = TryAddBesidePivot(pivotItem, newHeader, newSource);
            if (!added)
            {
                var project = await VS.Solutions.GetActiveProjectAsync();
                if (project != null)
                {
                    try { await project.AddExistingFilesAsync(newHeader, newSource); }
                    catch { /* vcxproj may fail here; at least files exist on disk */ }
                }
            }

            await VS.StatusBar.ShowMessageAsync($"Duplicated {oldBase} → {newBase}");
            Log.Post($"Done: {headerPath} + {sourcePath}  =>  {newHeader} + {newSource}");
        }

        // Detect dominant EOL in a file: "\r\n", "\n", or "\r" (fallback to Environment.NewLine if none found)
        private static string DetectDominantEol(string text)
        {
            int crlf = 0, lf = 0, cr = 0;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '\r')
                {
                    if (i + 1 < text.Length && text[i + 1] == '\n') { crlf++; i++; }
                    else { cr++; }
                }
                else if (c == '\n') { lf++; }
            }
            if (crlf >= lf && crlf >= cr) return "\r\n";
            if (lf >= crlf && lf >= cr) return "\n";
            if (cr >= crlf && cr >= lf) return "\r";
            return Environment.NewLine;
        }

        // Normalize all EOLs in 'text' to the specified 'eol'
        private static string NormalizeEol(string text, string eol)
        {
            // unify to LF first, then expand
            text = text.Replace("\r\n", "\n").Replace("\r", "\n");
            return eol == "\n" ? text : text.Replace("\n", eol);
        }

        // ---------- selection helpers (first item only) ----------
        private static (ProjectItem? item, string? path) GetFirstSelectedProjectItem()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var dte = (DTE2)ServiceProvider.GlobalProvider.GetService(typeof(DTE));
            if (dte == null) return (null, null);

            var se = (UIHierarchy)dte.Windows.Item(EnvDTE.Constants.vsWindowKindSolutionExplorer).Object;
            if (se?.SelectedItems is not Array selected || selected.Length == 0) return (null, null);

            var ui = selected.GetValue(0) as UIHierarchyItem;
            var pi = ui?.Object as ProjectItem;
            if (pi == null) return (null, null);

            try
            {
                string p = pi.FileCount > 0 ? pi.FileNames[1] : "";
                return (!string.IsNullOrEmpty(p) && File.Exists(p)) ? (pi, p) : (pi, null);
            }
            catch { return (pi, null); }
        }

        // ---------- pair resolution (same dir → project search) ----------
        private static bool TryResolvePair(string pivotPath, out string header, out string source)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            header = source = "";

            var dir = Path.GetDirectoryName(pivotPath)!;
            var bn = GetBaseName(pivotPath);

            if (IsHeader(pivotPath))
            {
                foreach (var ext in SourceExts)
                {
                    var cand = Path.Combine(dir, bn + ext);
                    if (File.Exists(cand)) { header = pivotPath; source = cand; return true; }
                }
                return TryFindComplementInProject(pivotPath, isHeader: true, out header, out source);
            }
            else if (IsSource(pivotPath))
            {
                foreach (var ext in HeaderExts)
                {
                    var cand = Path.Combine(dir, bn + ext);
                    if (File.Exists(cand)) { header = cand; source = pivotPath; return true; }
                }
                return TryFindComplementInProject(pivotPath, isHeader: false, out header, out source);
            }
            return false;
        }

        // project search via DTE (no Toolkit GetFilesAsync; avoid Task.Wait in BQS)
        private static bool TryFindComplementInProject(string pivotPath, bool isHeader, out string header, out string source)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            header = source = "";
            var dte = (DTE2)ServiceProvider.GlobalProvider.GetService(typeof(DTE));
            if (dte == null) return false;

            string bn = GetBaseName(pivotPath);
            bool match(string f, string[] exts) =>
                exts.Any(e => f.EndsWith(e, StringComparison.OrdinalIgnoreCase)) &&
                string.Equals(GetBaseName(f), bn, StringComparison.OrdinalIgnoreCase);

            foreach (EnvDTE.Project p in dte.Solution.Projects)
            {
                foreach (var f in EnumerateProjectFiles(p))
                {
                    if (isHeader)
                    {
                        if (match(f, SourceExts)) { header = pivotPath; source = f; return true; }
                    }
                    else
                    {
                        if (match(f, HeaderExts)) { header = f; source = pivotPath; return true; }
                    }
                }
            }
            return false;
        }

        private static IEnumerable<string> EnumerateProjectFiles(EnvDTE.Project project)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            IEnumerable<ProjectItem> Walk(ProjectItems items)
            {
                if (items == null) yield break;
                foreach (ProjectItem it in items)
                {
                    yield return it;
                    foreach (var child in Walk(it.ProjectItems)) yield return child;
                }
            }

            foreach (var item in Walk(project.ProjectItems))
            {
                string path = "";
                try { if (item.FileCount > 0) path = item.FileNames[1]; } catch { }
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                    yield return path;
            }
        }

        private static List<string> GetSelectedFilePathsViaDte()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var list = new List<string>();
            var dte = (DTE2)ServiceProvider.GlobalProvider.GetService(typeof(DTE));
            if (dte == null) return list;

            var se = (UIHierarchy)dte.Windows.Item(EnvDTE.Constants.vsWindowKindSolutionExplorer).Object;
            var selected = se?.SelectedItems as Array;
            if (selected == null) return list;

            foreach (UIHierarchyItem ui in selected)
            {
                var pi = ui.Object as ProjectItem;
                if (pi != null)
                {
                    try
                    {
                        string p = pi.FileNames[1];
                        if (!string.IsNullOrEmpty(p) && File.Exists(p))
                            list.Add(p);
                    }
                    catch { /* ignore non-file nodes */ }
                }
            }
            return list;
        }

        // ---------- add under the same filter/folder as the pivot ----------
        private static bool TryAddBesidePivot(ProjectItem? pivotItem, params string[] newFiles)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                if (pivotItem?.Collection != null)
                {
                    foreach (var f in newFiles)
                        pivotItem.Collection.AddFromFile(f);
                    return true;
                }
            }
            catch { /* VC++ filters can be finicky; fall back handled by caller */ }
            return false;
        }

        // ---------- improved two-field dialog with live preview + conflict gating ----------
        private sealed class ReplaceDialog : Form
        {
            public string FindText => _find.Text;
            public string ReplaceText => _repl.Text;

            public string PreviewBase { get; private set; } = "";
            public string NewHeaderPath { get; private set; } = "";
            public string NewSourcePath { get; private set; } = "";
            public bool HasConflicts { get; private set; }

            private string _dir;
            private readonly string _headerExt, _sourceExt, _oldBase;

            private readonly TextBox _find = new() { Dock = DockStyle.Fill };
            private readonly TextBox _repl = new() { Dock = DockStyle.Fill };

            private readonly TextBox _hdrPath = new() { ReadOnly = true, Dock = DockStyle.Fill };
            private readonly TextBox _srcPath = new() { ReadOnly = true, Dock = DockStyle.Fill };

            private readonly TextBox _outDir = new() { Dock = DockStyle.Fill, ReadOnly = true };
            private readonly Button _browse = new() { Text = "Browse…", AutoSize = true, Anchor = AnchorStyles.Right };

            private readonly Label _conflict = new() { AutoSize = true, ForeColor = System.Drawing.Color.Firebrick };

            private readonly Button _ok = new() { Text = "OK", DialogResult = DialogResult.OK, Anchor = AnchorStyles.Right };
            private readonly Button _cancel = new() { Text = "Cancel", DialogResult = DialogResult.Cancel, Anchor = AnchorStyles.Right };

            public ReplaceDialog(string headerPath, string sourcePath, string dir, string headerExt, string sourceExt, string oldBase)
            {
                Text = "Duplicate C++ Pair";
                AutoScaleMode = AutoScaleMode.Font;
                FormBorderStyle = FormBorderStyle.Sizable;             // resizable
                MaximizeBox = true;
                MinimizeBox = false;
                StartPosition = FormStartPosition.CenterParent;
                ClientSize = new System.Drawing.Size(1500, 360);       // wider default
                MinimumSize = new System.Drawing.Size(720, 320);
                Padding = new Padding(10);

                this.ShowIcon = true;
                this.ShowInTaskbar = false; // typical for modal tool dialogs

                // Directly embedded .ico file
                var asm = typeof(NF_VSToolsPackage).Assembly;
                // confirm the exact resource name if needed with asm.GetManifestResourceNames()
                using var resourceStream = asm.GetManifestResourceStream("NF.VSTools.Resources.NF_Icon_32.ico");
                if (resourceStream != null) this.Icon = new System.Drawing.Icon(resourceStream);

                _dir = dir; _headerExt = headerExt; _sourceExt = sourceExt; _oldBase = oldBase;

                var table = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 3,
                    RowCount = 9
                };
                table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));       // labels
                table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));   // inputs
                table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));       // buttons

                // Set row heights: autosize for content rows, fixed for button row
                const int ButtonRowHeight = 56; // ↑ bump from 48 to give room for taller buttons
                for (int r = 0; r < 8; r++) table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                table.RowStyles.Add(new RowStyle(SizeType.Absolute, ButtonRowHeight));  // bottom row

                // Make bottom buttons slightly taller + consistent
                _ok.MinimumSize = new System.Drawing.Size(96, 32);
                _cancel.MinimumSize = new System.Drawing.Size(96, 32);

                // optional: a touch more padding for hit target
                _ok.Padding = new Padding(4, 2, 4, 2);
                _cancel.Padding = new Padding(4, 2, 4, 2);

                // Row 0: Header path (read-only)
                table.Controls.Add(new Label { Text = "Header:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
                var txtHeader = new TextBox { Text = headerPath, ReadOnly = true, Dock = DockStyle.Fill };
                table.Controls.Add(txtHeader, 1, 0);
                table.SetColumnSpan(txtHeader, 2);

                // Row 1: Source path (read-only)
                table.Controls.Add(new Label { Text = "Source:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
                var txtSource = new TextBox { Text = sourcePath, ReadOnly = true, Dock = DockStyle.Fill };
                table.Controls.Add(txtSource, 1, 1);
                table.SetColumnSpan(txtSource, 2);

                // Row 2: Basename (read-only)
                table.Controls.Add(new Label { Text = "Basename:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
                var txtBase = new TextBox { Text = oldBase ?? "", ReadOnly = true, Anchor = AnchorStyles.Left, Width = 520 };
                table.Controls.Add(txtBase, 1, 2);

                // Row 3: Output folder + browse
                table.Controls.Add(new Label { Text = "Output folder:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 3);
                _outDir.Text = _dir;
                table.Controls.Add(_outDir, 1, 3);
                table.Controls.Add(_browse, 2, 3);

                // Row 4: Find
                table.Controls.Add(new Label { Text = "Find:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 4);
                _find.Text = oldBase ?? ""; // default
                table.Controls.Add(_find, 1, 4);
                table.SetColumnSpan(_find, 2);

                // Row 5: Replace with
                table.Controls.Add(new Label { Text = "Replace with:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 5);
                _repl.Text = oldBase ?? ""; // default to basename
                table.Controls.Add(_repl, 1, 5);
                table.SetColumnSpan(_repl, 2);

                // Row 6–7: Preview new file paths
                table.Controls.Add(new Label { Text = "New header:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 6);
                table.Controls.Add(_hdrPath, 1, 6);
                table.SetColumnSpan(_hdrPath, 2);

                table.Controls.Add(new Label { Text = "New source:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 7);
                table.Controls.Add(_srcPath, 1, 7);
                table.SetColumnSpan(_srcPath, 2);

                // Row 8: conflict label + OK/Cancel
                var btnPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill, Padding = new Padding(0, 6, 0, 0) };
                btnPanel.Controls.Add(_cancel);
                btnPanel.Controls.Add(_ok);
                table.Controls.Add(_conflict, 0, 8);
                table.Controls.Add(btnPanel, 1, 8);
                table.SetColumnSpan(btnPanel, 2);

                Controls.Add(table);
                AcceptButton = _ok; CancelButton = _cancel;

                // Wire events
                _find.TextChanged += (_, __) => UpdatePreview();
                _repl.TextChanged += (_, __) => UpdatePreview();
                _browse.Click += (_, __) => BrowseFolder();

                // Initial preview
                UpdatePreview();
            }

            private void BrowseFolder()
            {
                using var fb = new FolderBrowserDialog
                {
                    Description = "Select output folder",
                    SelectedPath = Directory.Exists(_dir) ? _dir : Environment.CurrentDirectory, // default to current output
                    ShowNewFolderButton = true
                };
                if (fb.ShowDialog(this) == DialogResult.OK && Directory.Exists(fb.SelectedPath))
                {
                    _dir = fb.SelectedPath;
                    _outDir.Text = _dir;
                    UpdatePreview();
                }
            }

            private void UpdatePreview()
            {
                var find = _find.Text ?? "";
                var repl = _repl.Text ?? "";

                PreviewBase = (_oldBase ?? "").Replace(find, repl);
                NewHeaderPath = Path.Combine(_dir, PreviewBase + _headerExt);
                NewSourcePath = Path.Combine(_dir, PreviewBase + _sourceExt);

                _hdrPath.Text = NewHeaderPath;
                _srcPath.Text = NewSourcePath;

                bool invalid = PreviewBase.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0;
                HasConflicts = File.Exists(NewHeaderPath) || File.Exists(NewSourcePath);

                _conflict.Text = invalid
                    ? "Resulting base name contains invalid characters."
                    : (HasConflicts ? "One or both target files already exist." : "");

                _ok.Enabled = !invalid && !HasConflicts && !string.IsNullOrWhiteSpace(PreviewBase) && Directory.Exists(_dir);
            }
        }
    }
}
