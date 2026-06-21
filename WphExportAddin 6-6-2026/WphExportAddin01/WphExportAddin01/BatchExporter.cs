using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace WphExportAddin
{
    /// <summary>
    /// Drives the multi-model batch. For every selected .rvt it:
    ///   1. opens the file in the BACKGROUND (no UI view), detaching if the
    ///      file is workshared so no central / workset prompt blocks the run;
    ///   2. runs the SAME <see cref="WphExporter"/> you already use, so the
    ///      NWC / IFC settings are identical to the single-model command
    ///      (and the per-view/format delay also applies inside each model);
    ///   3. closes the document (without saving);
    ///   4. waits ExportSettings.DelaySeconds before the next file (skipped
    ///      after the last one) so Revit can release handles and settle.
    ///
    /// Every model is wrapped in its own try / catch: a single bad file is
    /// logged and skipped, the rest of the batch keeps going. The per-format
    /// (NWC / IFC) isolation already lives inside WphExporter, so the two
    /// layers together mean the batch never aborts midway.
    /// </summary>
    public class BatchExporter
    {
        private readonly Application _app;          // the Revit Application (opens docs)
        private readonly ExportSettings _settings;  // shared export options (incl. DelaySeconds)
        private readonly List<string> _files;       // .rvt paths to process
        private readonly StringBuilder _log = new StringBuilder();

        public BatchExporter(UIApplication uiApp, ExportSettings settings, List<string> files)
        {
            if (uiApp == null) throw new ArgumentNullException(nameof(uiApp));
            _app = uiApp.Application;
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _files = files ?? throw new ArgumentNullException(nameof(files));
        }

        private void Log(string line) => _log.AppendLine(line);

        public string Run()
        {
            int delay = _settings.DelaySeconds < 0 ? 0 : _settings.DelaySeconds;

            Log(string.Format("=== RUKN BIM API BATCH EXPORT @ {0} ===",
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            Log(string.Format("models={0}  mode={1}  delay={2}s  folder={3}",
                _files.Count, _settings.Mode, delay, _settings.OutputFolder));
            Log("");

            int index = 0;
            int ok = 0, failed = 0;

            foreach (string path in _files)
            {
                index++;
                string name = SafeFileName(path);
                Log("=========================================");
                Log(string.Format("[{0}/{1}] {2}", index, _files.Count, name));
                Log("=========================================");

                Document doc = null;
                try
                {
                    if (!File.Exists(path))
                    {
                        Log("MODEL : SKIPPED -> file not found on disk: " + path);
                        failed++;
                    }
                    else
                    {
                        doc = OpenBackground(path);
                        if (doc == null)
                        {
                            Log("MODEL : SKIPPED -> could not open (see message above).");
                            failed++;
                        }
                        else
                        {
                            // Reuse the exact same exporter as the single-model run.
                            var exporter = new WphExporter(doc, _settings);
                            Log(exporter.Run());
                            ok++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log("MODEL : ERROR   -> " + Flatten(ex));
                    failed++;
                }
                finally
                {
                    // Always close a background-opened doc so memory is freed
                    // before the next file. Never closes the user's active doc
                    // (we only ever close docs we opened here).
                    if (doc != null)
                    {
                        try { doc.Close(false); }
                        catch (Exception ex) { Log("MODEL : WARN   -> close failed: " + Flatten(ex)); }
                    }
                }

                // Pause between models (not after the last one). This blocks
                // Revit for the duration, which is expected for a batch.
                if (index < _files.Count && delay > 0)
                {
                    Log(string.Format("... waiting {0}s before next model ...", delay));
                    Thread.Sleep(delay * 1000);
                }

                Log("");
            }

            Log("=========================================");
            Log(string.Format("BATCH COMPLETE -> {0} ok, {1} failed, {2} total.",
                ok, failed, _files.Count));
            Log("-----------------------------------------");
            Log(AddinInfo.CreditBlock());
            return _log.ToString();
        }

        /// <summary>
        /// Open a model with no UI view. Detaches workshared files and opens all
        /// worksets so geometry is complete for export. Returns null on failure
        /// (already logged).
        /// </summary>
        private Document OpenBackground(string path)
        {
            try
            {
                ModelPath mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(path);
                var openOpts = new OpenOptions();

                // Detect workshared so we open detached (no central interaction).
                bool workshared = false;
                try
                {
                    BasicFileInfo info = BasicFileInfo.Extract(path);
                    workshared = info != null && info.IsWorkshared;
                }
                catch { /* if we can't read it, treat as non-workshared */ }

                if (workshared)
                    openOpts.DetachFromCentralOption =
                        DetachFromCentralOption.DetachAndPreserveWorksets;

                // Open every workset so nothing is missing from the export.
                openOpts.SetOpenWorksetsConfiguration(
                    new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets));

                Document doc = _app.OpenDocumentFile(mp, openOpts);
                if (doc == null)
                    Log("MODEL : OPEN FAILED -> Revit returned no document.");
                return doc;
            }
            catch (Exception ex)
            {
                // Common causes: file already open in this session, version
                // mismatch (newer .rvt than this Revit), or corrupt file.
                Log("MODEL : OPEN ERROR -> " + Flatten(ex));
                return null;
            }
        }

        private static string SafeFileName(string path)
        {
            try { return Path.GetFileName(path); } catch { return path; }
        }

        private static string Flatten(Exception ex)
        {
            while (ex.InnerException != null) ex = ex.InnerException;
            return string.Format("{0}: {1}", ex.GetType().Name, ex.Message);
        }
    }
}
