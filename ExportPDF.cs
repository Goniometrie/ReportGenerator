using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Grasshopper.Kernel;

namespace BalkroosterExport
{
    public class ExportPDF : GH_Component
    {
        public ExportPDF()
          : base("ExportPDF", "ExpPDF",
            "Convert a .docx file to a .pdf file using Microsoft Word (Late Binding).",
            "Document", "Export")
        {
        }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("DocPath", "Doc", "Path to the Word .docx file.", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Create", "Create", "When true, convert the file to PDF.", GH_ParamAccess.item, false);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("PdfPath", "Pdf", "Path to the generated PDF file.", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Success", "Ok", "True when conversion succeeded.", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string docPath = string.Empty;
            bool create = false;

            if (!DA.GetData(0, ref docPath)) return;
            DA.GetData(1, ref create);

            bool success = false;
            string pdfPath = string.Empty;

            if (string.IsNullOrWhiteSpace(docPath))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "DocPath is empty.");
                DA.SetData(0, pdfPath);
                DA.SetData(1, success);
                return;
            }

            if (!File.Exists(docPath))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Document not found: {docPath}");
                DA.SetData(0, pdfPath);
                DA.SetData(1, success);
                return;
            }

            if (!create)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Create is false — no conversion performed.");
                DA.SetData(0, pdfPath);
                DA.SetData(1, success);
                return;
            }

            // Target PDF path
            try
            {
                pdfPath = Path.ChangeExtension(docPath, ".pdf");
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Invalid output path: {ex.Message}");
                DA.SetData(0, string.Empty);
                DA.SetData(1, false);
                return;
            }

            // Late Binding variables
            object wordApp = null;
            object doc = null;

            try
            {
                // Get Type from ProgID for Word.Application
                Type wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Microsoft Word is not installed (could not find ProgID 'Word.Application').");
                    DA.SetData(0, string.Empty);
                    DA.SetData(1, false);
                    return;
                }

                // Create Word Instance
                wordApp = Activator.CreateInstance(wordType);
                
                // Set Visible = false
                wordType.InvokeMember("Visible", BindingFlags.SetProperty, null, wordApp, new object[] { false });

                // Set DisplayAlerts = 0 (wdAlertsNone)
                // wordType.InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, wordApp, new object[] { 0 });

                // Open Document
                // Documents.Open(FileName, ...)
                object documents = wordType.InvokeMember("Documents", BindingFlags.GetProperty, null, wordApp, null);
                
                object[] openArgs = new object[] 
                { 
                    docPath, // FileName
                    false,   // ConfirmConversions
                    true,    // ReadOnly
                    false,   // AddToRecentFiles
                    // ... other optional args are omitted in default dispatch or passed as Missing.Value if strictly positional, 
                    // but for simplified late binding we can often just pass the first few if the underlying COM supports it.
                    // However, C# dynamic is easier than pure reflection for optional params.
                };
                
                // Using 'dynamic' makes late binding MUCH easier than raw Reflection
                dynamic dWordApp = wordApp;
                dWordApp.DisplayAlerts = 0; // wdAlertsNone

                doc = dWordApp.Documents.Open(
                    FileName: docPath, 
                    ReadOnly: true, 
                    AddToRecentFiles: false
                );

                // ExportAsFixedFormat contants:
                // wdExportFormatPDF = 17
                // wdExportOptimizeForPrint = 0
                // wdExportAllDocument = 0
                // wdExportDocumentContent = 0
                // wdExportCreateHeadingBookmarks = 1

                ((dynamic)doc).ExportAsFixedFormat(
                    OutputFileName: pdfPath,
                    ExportFormat: 17, // PDF
                    OpenAfterExport: false,
                    OptimizeFor: 0,   // Print
                    Range: 0,         // All
                    Item: 0,          // Content
                    IncludeDocProps: true,
                    KeepIRM: true,
                    CreateBookmarks: 1, // Heading bookmarks
                    DocStructureTags: true,
                    BitmapMissingFonts: true,
                    UseISO19005_1: false
                );

                success = true;
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"Successfully exported PDF: {pdfPath}");
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Error during export: {ex.Message}");
                // Unwrap InnerException if available (common in reflection)
                if (ex.InnerException != null)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Inner Error: {ex.InnerException.Message}");
                }
            }
            finally
            {
                // Close Doc
                if (doc != null)
                {
                    try
                    {
                        // wdDoNotSaveChanges = 0
                        ((dynamic)doc).Close(SaveChanges: 0);
                    }
                    catch { /* ignore close errors */ }
                    
                    Marshal.ReleaseComObject(doc);
                }

                // Quit App
                if (wordApp != null)
                {
                    try
                    {
                        ((dynamic)wordApp).Quit();
                    }
                    catch { /* ignore quit errors */ }

                    Marshal.ReleaseComObject(wordApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            DA.SetData(0, pdfPath);
            DA.SetData(1, success);
        }

        public override GH_Exposure Exposure => GH_Exposure.primary;

        protected override System.Drawing.Bitmap Icon => null;

        public override Guid ComponentGuid => new Guid("4e7c3a72-1b2c-4a3b-9f8e-2d4c6d8e0a12");
    }
}
