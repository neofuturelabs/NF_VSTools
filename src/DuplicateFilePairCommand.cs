using Community.VisualStudio.Toolkit;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Package;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.VCProjectEngine;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;  // for Encoding
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

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

        // Logging shortcut for this command
        private static void LOG(string msg) => Log.Post("[DuplicatePair] " + msg);

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

            LOG("----- START -----");

            var (pivotItem, pivotPath) = GetFirstSelectedProjectItem();
            if (pivotPath == null || !(IsHeader(pivotPath) || IsSource(pivotPath)))
            {
                await VS.StatusBar.ShowMessageAsync("Select a C++ header/source file.");
                return;
            }

            LOG($"Pivot selected: {pivotPath}");

            // Resolve the complementary file: same dir first, then search the containing project
            if (!TryResolvePair(pivotPath, out string headerPath, out string sourcePath))
            {
                LOG("Could not find matching header/source pair.");
                await VS.StatusBar.ShowMessageAsync("Could not find matching header/source pair.");
                return;
            }

            LOG($"Pair resolved: header='{headerPath}' source='{sourcePath}'");

            string dir = Path.GetDirectoryName(headerPath)!;
            string oldBase = GetBaseName(headerPath);

            // Respect original extensions
            var headerExt = Path.GetExtension(headerPath);
            var sourceExt = Path.GetExtension(sourcePath);

            // UI dialog: note we pass dir/exts so it can preview full paths and conflicts
            using var dlg = new ReplaceDialog(headerPath, sourcePath, dir, headerExt, sourceExt, oldBase);
            LOG($"Dialog init: dir='{dir}' headerExt='{headerExt}' sourceExt='{sourceExt}' oldBase='{oldBase}'");

            if (dlg.ShowDialog() != DialogResult.OK)
            {
                LOG("Dialog cancelled by user.");
                return;
            }

            string find = dlg.FindText;
            string replace = dlg.ReplaceText;
            if (string.IsNullOrEmpty(find))
            {
                LOG("Find text is empty; cancelled.");
                await VS.StatusBar.ShowMessageAsync("Find text is empty; cancelled.");
                return;
            }

            string newBase = dlg.PreviewBase;
            string newHeaderPath = dlg.NewHeaderPath;
            string newSourcePath = dlg.NewSourcePath;

            LOG($"Execution plan: find='{find}' → replace='{replace}', newBase='{newBase}'");
            LOG($"Output paths: header='{newHeaderPath}', source='{newSourcePath}'");

            // double-check conflicts (defense-in-depth)
            if (File.Exists(newHeaderPath) || File.Exists(newSourcePath))
            {
                LOG("Target files already exist; aborting.");
                await VS.StatusBar.ShowMessageAsync("Target files already exist; aborting.");
                return;
            }

            // Copy
            File.Copy(headerPath, newHeaderPath, overwrite: false);
            File.Copy(sourcePath, newSourcePath, overwrite: false);
            LOG($"File created: {newHeaderPath}");
            LOG($"File created: {newSourcePath}");

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
            var (hOrig, hEnc) = await ReadWithEncodingAsync(newHeaderPath);
            var headerEol = DetectDominantEol(hOrig);
            var h = hOrig;
            h = h.Replace($"{oldBase}.generated.h", $"{newBase}.generated.h");
            h = Regex.Replace(h, Regex.Escape(find), replace);
            h = NormalizeEol(h, headerEol);
            await WriteWithEncodingAsync(newHeaderPath, h, hEnc);
            LOG($"Header rewrite: EOL='{EolName(headerEol)}' Encoding='{hEnc?.WebName ?? "<unknown>"}' Changed={(h != hOrig)}");

            // Rewrite NEW source (preserve encoding + dominant EOL)
            var (cppOrig, cEnc) = await ReadWithEncodingAsync(newSourcePath);
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
            await WriteWithEncodingAsync(newSourcePath, cpp, cEnc);
            LOG($"Source rewrite: EOL='{EolName(sourceEol)}' Encoding='{cEnc?.WebName ?? "<unknown>"}' Changed={(cpp != cppOrig)}");


            // Add under destination filter(s) that mirror the chosen output folder(s)
            if (!TryAddUnderDestinationFilters(pivotItem, newHeaderPath, newSourcePath))
            {
                // very last fallback: add at project root if VCProject engine isn’t available
                var project = await VS.Solutions.GetActiveProjectAsync();
                if (project != null)
                {
                    try
                    {
                        await project.AddExistingFilesAsync(newHeaderPath, newSourcePath);
                        LOG("Added files via fallback: AddExistingFilesAsync at project root.");
                    }
                    catch (Exception ex)
                    {
                        LOG($"Failed to add files via fallback: {ex.Message}");
                        // files still exist on disk
                    }
                }
            }

            await VS.StatusBar.ShowMessageAsync($"Duplicated {oldBase} → {newBase}");
            LOG($"Done: {headerPath} + {sourcePath}  =>  {newHeaderPath} + {newSourcePath}");
        }

        // helper for readable EOL names
        private static string EolName(string eol) => eol switch
        {
            "\r\n" => "CRLF",
            "\n" => "LF",
            "\r" => "CR",
            _ => "Unknown"
        };

        // ---------- add to filters that mirror the destination folder(s) ----------
        private static bool TryAddUnderDestinationFilters(ProjectItem? pivotItem, params string[] absolutePaths)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var dteProj = pivotItem?.ContainingProject;
            if (dteProj == null) { LOG("AddUnderDestination: no containing project."); return false; }

            // Only works for VC++ projects
            if (dteProj.Object is not VCProject vcproj) { LOG("AddUnderDestination: not a VC++ project."); return false; }

            string vcProjDir = GetProjectDir(dteProj); // may be ...\Intermediate\ProjectFiles
            LOG($"AddUnderDestination: project='{dteProj.Name}', vcProjDir='{vcProjDir}'");

            foreach (var file in absolutePaths.Where(f => !string.IsNullOrWhiteSpace(f)))
            {
                string? absDir = Path.GetDirectoryName(file);
                if (string.IsNullOrWhiteSpace(absDir)) continue;

                bool underProject = TryMakeRelativeDirPathMultiBase(dteProj, absDir!, out string relDir, out string baseUsed);
                string projName = GetProjectName(dteProj); // UI thread asserted earlier in this method
                relDir = SanitizeRelDirForProject(projName, relDir, baseUsed);
                LOG($"Planning add: file='{file}', absDir='{absDir}', baseUsed='{baseUsed}', underProject={underProject}, relDir='{relDir}', projName='{projName}'");

                // Build (or find) a filter hierarchy that mirrors the relative directory.
                // If relDir == "" we add to project root (no filter).
                VCFilter? targetFilter = EnsureFilterPath(vcproj, relDir);
                if (targetFilter != null)
                    LOG($"Target filter path ensured: '{FilterFullPath(targetFilter)}'");
                else
                    LOG("Target filter: <project root>");

                // If the file is already in the project (possibly under the wrong filter), move it.
                VCFile? existing = FindVCFile(vcproj, file);
                if (existing != null)
                {
                    try { existing.Remove(); LOG($"(Shouldn't ever happen?) Existing project item removed before re-add: '{file}'"); }
                    catch (Exception ex) { LOG($"(Really shouldn't ever happen??) Existing item remove failed (continuing): '{file}' — {ex.Message}"); }
                }

                try
                {
                    if (targetFilter != null) { targetFilter.AddFile(file); LOG($"Added to filter: '{file}'"); }
                    else { vcproj.AddFile(file); LOG($"Warning: Added to project root: '{file}'"); }
                }
                catch (Exception ex)
                {
                    // If AddFile throws (rare), try project root as an emergency fallback.
                    LOG($"Error: Add to target filter failed (will try root): '{file}' — {ex.Message}");
                    try { vcproj.AddFile(file); LOG($"Warning: Added to project root (fallback): '{file}'"); }
                    catch (Exception ex2) { LOG($"Error: Add to project failed: '{file}' — {ex2.Message}"); }
                }
            }

            return true;
        }        

        private static string SanitizeRelDirForProject(string projectName, string relDir, string baseUsed)
        {
            if (string.IsNullOrWhiteSpace(relDir)) return relDir;

            var segs = relDir.Replace('/', '\\')
                             .Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            if (segs.Length == 0) return relDir;

            // If relDir came from the solution root, it often starts with "<ProjectName>\(Source|Plugins)\..."
            // Inside a VC++ project, filters should be relative to the project root, so drop that first segment.
            if (segs.Length >= 2 &&
                !string.IsNullOrEmpty(projectName) &&
                segs[0].Equals(projectName, StringComparison.OrdinalIgnoreCase) &&
                (segs[1].Equals("Source", StringComparison.OrdinalIgnoreCase) ||
                 segs[1].Equals("Plugins", StringComparison.OrdinalIgnoreCase)))
            {
                var stripped = string.Join("\\", segs.Skip(1));
                LOG($"relDir sanitized: stripped project head '{projectName}\\' -> '{stripped}' (baseUsed='{baseUsed}')");
                relDir = stripped;
            }

            // Collapse accidental repeated head like "Plugins\\Foo\\Plugins\\Foo\\..."
            string dedup = CollapseRepeatedHeadSegments(relDir);
            if (!string.Equals(dedup, relDir, StringComparison.OrdinalIgnoreCase))
                LOG($"relDir normalized (dedup head): '{relDir}' -> '{dedup}'");

            return dedup;
        }

        private static string CollapseRepeatedHeadSegments(string relDir)
        {
            if (string.IsNullOrWhiteSpace(relDir)) return relDir;

            var segs = relDir.Replace('/', '\\')
                             .Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            if (segs.Length < 2) return relDir;

            int maxK = Math.Min(5, segs.Length / 2);
            for (int k = 1; k <= maxK; k++)
            {
                bool same = true;
                for (int i = 0; i < k; i++)
                {
                    if (!segs[i].Equals(segs[i + k], StringComparison.OrdinalIgnoreCase))
                    {
                        same = false; break;
                    }
                }
                if (same)
                {
                    var dedup = new List<string>(segs.Take(k));
                    dedup.AddRange(segs.Skip(2 * k));
                    return string.Join("\\", dedup);
                }
            }
            return relDir;
        }

        private static string GetProjectName(EnvDTE.Project dteProj)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try { return dteProj?.Name ?? string.Empty; }
            catch { return string.Empty; }
        }

        private static string GetProjectDir(EnvDTE.Project dteProj)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                if (dteProj.Object is VCProject vcp)
                    return NormalizePath(vcp.ProjectDirectory); // VC++ authoritative project dir
            }
            catch { /* ignore */ }

            try
            {
                var d = Path.GetDirectoryName(dteProj.FullName) ?? "";
                return NormalizePath(d);
            }
            catch { return ""; }
        }

        private static VCFilter? EnsureFilterPath(VCProject vcproj, string relDir)
        {
            if (string.IsNullOrWhiteSpace(relDir)) return null;

            string[] segments = relDir
                .Replace('/', '\\')
                .Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);

            VCFilter? parent = null;
            int i = 0;

            while (i < segments.Length)
            {
                // 1) Try to match an existing composite filter at this level (e.g., "Plugins\<Proj>")
                if (TryMatchCompositeFilterAtLevel(vcproj, parent, segments, i, out VCFilter? composite, out int consumed))
                {
                    LOG($"Filter ensure (composite match): '{string.Join("\\", segments.Take(i + consumed))}' via '{composite!.Name}' (consumed {consumed} segs)");
                    parent = composite;
                    i += consumed;
                    continue;
                }

                // 2) No composite match → proceed with single-segment ensure
                string seg = segments[i];
                parent = parent == null
                    ? FindOrAddTopLevelFilter(vcproj, seg)
                    : FindOrAddChildFilter(parent, seg);

                LOG($"Filter ensure: '{string.Join("\\", segments.Take(i + 1))}' (createdOrFound='{seg}')");
                i++;
            }

            return parent;
        }

        // Try to find a filter at the current level whose Name equals a prefix of the remaining path segments joined by '\'
        // e.g., remaining "Plugins","MyPlugin","Source","Module" might match a single filter named "Plugins\\MyPlugin"
        private static bool TryMatchCompositeFilterAtLevel(VCProject vcproj, VCFilter? parent, string[] segs, int startIndex, out VCFilter? match, out int consumed)
        {
            match = null; consumed = 0;

            // Search scope: top-level vs child filters
            IEnumerable<VCFilter> scope = parent == null
                ? ((IVCCollection)vcproj.Filters).Cast<VCFilter>()
                : ((IVCCollection)parent.Filters).Cast<VCFilter>();

            // Limit how deep we try (2..6 segments is plenty and avoids silly joins)
            int maxJoin = Math.Min(6, segs.Length - startIndex);

            // Try the **longest** possible prefix first
            for (int k = maxJoin; k >= 2; k--)
            {
                string candidateName = string.Join("\\", segs.Skip(startIndex).Take(k));
                foreach (var f in scope)
                {
                    if (f?.Name != null && f.Name.Equals(candidateName, StringComparison.OrdinalIgnoreCase))
                    {
                        match = f;
                        consumed = k;
                        return true;
                    }
                }
            }
            return false;
        }
        private static VCFilter FindOrAddTopLevelFilter(VCProject vcproj, string name)
        {
            // Search existing top-level filters
            foreach (VCFilter f in (IVCCollection)vcproj.Filters)
            {
                if (string.Equals(f.Name, name, StringComparison.OrdinalIgnoreCase))
                    return f;
            }
            // Create it
            var created = (VCFilter)vcproj.AddFilter(name);
            LOG($"Filter created (top): '{name}'");
            return created;
        }

        private static VCFilter FindOrAddChildFilter(VCFilter parent, string name)
        {
            foreach (VCFilter f in (IVCCollection)parent.Filters)
            {
                if (string.Equals(f.Name, name, StringComparison.OrdinalIgnoreCase))
                    return f;
            }
            var created = (VCFilter)parent.AddFilter(name);
            LOG($"Filter created (child of '{FilterFullPath(parent)}'): '{name}'");
            return created;
        }

        // Build a readable full path for logs (best-effort)
        private static string FilterFullPath(VCFilter f)
        {
            try
            {
                var names = new List<string>();
                VCFilter? cur = f;
                while (cur != null)
                {
                    names.Add(cur.Name);
                    cur = cur.Parent as VCFilter;
                }
                names.Reverse();
                return string.Join("\\", names);
            }
            catch { return f?.Name ?? "<unknown>"; }
        }

        private static VCFile? FindVCFile(VCProject vcproj, string fullPath)
        {
            string norm = NormalizePath(fullPath);
            foreach (VCFile f in (IVCCollection)vcproj.Files)
            {
                try
                {
                    if (string.Equals(NormalizePath(f.FullPath), norm, StringComparison.OrdinalIgnoreCase))
                        return f;
                }
                catch { /* some items may not have FullPath */ }
            }
            return null;
        }

        private static string NormalizePath(string p)
        {
            try { return Path.GetFullPath(p).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar); }
            catch { return p; }
        }

        // Choose the best ancestor base dir (VC dir, project dir, solution dir, parents of VC dir)
        // that actually contains targetDir. Returns relDir and which base was used (for logs).
        private static bool TryMakeRelativeDirPathMultiBase(EnvDTE.Project dteProj, string targetDir, out string relDir, out string baseUsed)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            relDir = ""; baseUsed = "";

            var candidates = new List<string>();

            // 1) VC++ project directory (often ...\Intermediate\ProjectFiles for UE)
            try { if (dteProj.Object is VCProject vcp && !string.IsNullOrWhiteSpace(vcp.ProjectDirectory)) candidates.Add(vcp.ProjectDirectory); } catch { }

            // 2) Directory of the .vcxproj
            try { if (!string.IsNullOrWhiteSpace(dteProj.FullName)) candidates.Add(Path.GetDirectoryName(dteProj.FullName)!); } catch { }

            // 3) Solution directory (usually the UE project root)
            try { var sln = dteProj.DTE?.Solution?.FullName; if (!string.IsNullOrWhiteSpace(sln)) candidates.Add(Path.GetDirectoryName(sln)!); } catch { }

            // 4) Parents of the VC project dir (hop out of Intermediate\ProjectFiles)
            try
            {
                if (dteProj.Object is VCProject vcp2 && !string.IsNullOrWhiteSpace(vcp2.ProjectDirectory))
                {
                    string d = vcp2.ProjectDirectory;
                    for (int i = 0; i < 3; i++)
                    {
                        d = Directory.GetParent(d)?.FullName ?? "";
                        if (string.IsNullOrWhiteSpace(d)) break;
                        candidates.Add(d);
                    }
                }
            }
            catch { }

            string targFull = NormalizePath(targetDir);

            string bestRel = "";
            string bestBase = "";
            int bestScore = int.MinValue;

            foreach (var cand in candidates.Select(NormalizePath).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (TryMakeRelativeDirPath(cand, targFull, out var rel))
                {
                    // Prefer the deepest ancestor (longest base path) that contains target
                    int score = cand.Length;
                    if (score > bestScore)
                    {
                        bestScore = score;
                        bestBase = cand;
                        bestRel = rel;
                    }
                }
            }

            bool ok = bestScore != int.MinValue;
            baseUsed = ok ? bestBase : "";
            relDir = ok ? bestRel : "";
            LOG($"TryMakeRelativeDirPathMultiBase: target='{targFull}', baseUsed='{baseUsed}', ok={ok}, relDir='{relDir}'");
            return ok;
        }

        private static bool TryMakeRelativeDirPath(string baseDir, string targetDir, out string relDir)
        {
            relDir = "";
            if (string.IsNullOrWhiteSpace(baseDir) || string.IsNullOrWhiteSpace(targetDir)) return false;

            try
            {
                // Normalize to absolute, unify separators, and ensure trailing slash for directories
                string baseFull = AppendSlash(Path.GetFullPath(baseDir));
                string targFull = AppendSlash(Path.GetFullPath(targetDir));

                var baseUri = new Uri(baseFull, UriKind.Absolute);
                var targUri = new Uri(targFull, UriKind.Absolute);

                bool isBase = baseUri.IsBaseOf(targUri);
                LOG($"TryMakeRelativeDirPath: baseFull='{baseFull}', targFull='{targFull}', isBase={isBase}");
                if (!isBase) return false;

                var rel = Uri.UnescapeDataString(baseUri.MakeRelativeUri(targUri).ToString());
                relDir = rel.Replace('/', '\\').TrimEnd('\\');
                return true;
            }
            catch (Exception ex)
            {
                LOG($"TryMakeRelativeDirPath Error: {ex.Message}");
                return false;
            }
        }

        private static string AppendSlash(string d)
        {
            if (string.IsNullOrEmpty(d)) return d;

            if (d.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) ||
                d.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal))
            {
                return d;
            }

            return d + Path.DirectorySeparatorChar;
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
                    if (File.Exists(cand))
                    {
                        header = pivotPath; source = cand;
                        LOG($"Pair found in same directory: '{cand}'");
                        return true;
                    }
                }
                LOG("Pair not in same directory; scanning project for source...");
                return TryFindComplementInProject(pivotPath, isHeader: true, out header, out source);
            }
            else if (IsSource(pivotPath))
            {
                foreach (var ext in HeaderExts)
                {
                    var cand = Path.Combine(dir, bn + ext);
                    if (File.Exists(cand))
                    {
                        header = cand; source = pivotPath;
                        LOG($"Pair found in same directory: '{cand}'");
                        return true;
                    }
                }
                LOG("Pair not in same directory; scanning project for header...");
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

            int scanned = 0;
            foreach (EnvDTE.Project p in dte.Solution.Projects)
            {
                foreach (var f in EnumerateProjectFiles(p))
                {
                    scanned++;
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
            LOG($"Project scan finished, complement not found (scanned {scanned} files).");
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
