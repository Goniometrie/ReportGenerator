using System;
using System.IO;
using System.Linq;
using System.Globalization;

using Grasshopper;
using Grasshopper.Kernel;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace BalkroosterExport
{
    public class ProjectInfo : GH_Component
    {
        public ProjectInfo()
          : base("ProjectInfo", "ProjectInfo",
            "Fill project information (right column) in a project-info table inside a .docx template.",
            "Document", "Export")
        { }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("DocPath", "Doc", "Path to working .docx file", GH_ParamAccess.item);
            pManager.AddTextParameter("ProjectName", "Name", "Project name", GH_ParamAccess.item, string.Empty);
            pManager.AddTextParameter("Client", "Client", "Client name", GH_ParamAccess.item, string.Empty);
            pManager.AddTextParameter("Date", "Date", "Project date (will be written as yyyy-MM-dd). If empty uses now.", GH_ParamAccess.item, string.Empty);
            pManager.AddTextParameter("Version", "Ver", "Project version", GH_ParamAccess.item, string.Empty);
            pManager.AddTextParameter("Author", "Author", "Author name", GH_ParamAccess.item, string.Empty);
            pManager.AddTextParameter("CheckedBy", "Checked", "Checked by", GH_ParamAccess.item, string.Empty);
            pManager.AddBooleanParameter("Add", "Add", "When true perform update on the working document", GH_ParamAccess.item, false);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("FilePath", "File", "Path to the written file (working copy)", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Success", "Ok", "True when write succeeded", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string docPath = string.Empty;
            string projectName = string.Empty;
            string client = string.Empty;
            string dateInput = string.Empty;
            string version = string.Empty;
            string author = string.Empty;
            string checkedBy = string.Empty;
            bool add = false;

            if (!DA.GetData(0, ref docPath)) return;
            DA.GetData(1, ref projectName);
            DA.GetData(2, ref client);
            DA.GetData(3, ref dateInput);
            DA.GetData(4, ref version);
            DA.GetData(5, ref author);
            DA.GetData(6, ref checkedBy);
            DA.GetData(7, ref add);

            bool success = false;
            string writtenPath = string.Empty;

            if (string.IsNullOrWhiteSpace(docPath))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "DocPath is empty.");
                DA.SetData(0, writtenPath);
                DA.SetData(1, success);
                return;
            }

            if (!File.Exists(docPath))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Document not found: {docPath}");
                DA.SetData(0, writtenPath);
                DA.SetData(1, success);
                return;
            }

            if (!add)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Add is false — no changes made.");
                DA.SetData(0, writtenPath);
                DA.SetData(1, success);
                return;
            }

            DateTime dateValue;
            if (string.IsNullOrWhiteSpace(dateInput))
            {
                dateValue = DateTime.Now;
            }
            else if (!DateTime.TryParse(dateInput, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out dateValue)
                     && !DateTime.TryParse(dateInput, out dateValue))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Date '{dateInput}' could not be parsed. Using current date.");
                dateValue = DateTime.Now;
            }
            string dateText = dateValue.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

            try
            {
                using (var doc = WordprocessingDocument.Open(docPath, true))
                {
                    var body = doc.MainDocumentPart?.Document?.Body;
                    if (body == null)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Document body is null or missing.");
                        DA.SetData(0, writtenPath);
                        DA.SetData(1, success);
                        return;
                    }

                    Table matchedTable = null;
                    foreach (var table in body.Elements<Table>())
                    {
                        var firstColTexts = table.Elements<TableRow>()
                                                 .Select(r => r.Elements<TableCell>().FirstOrDefault())
                                                 .Where(c => c != null)
                                                 .Select(GetCellPlainText)
                                                 .Select(t => t.Trim().ToLowerInvariant())
                                                 .ToList();

                        bool hasProject = firstColTexts.Any(t => t.Contains("projectname"));
                        bool hasClient = firstColTexts.Any(t => t.Contains("client"));
                        bool hasDateKey = firstColTexts.Any(t => t.Contains("date"));

                        if (hasProject && hasClient && hasDateKey)
                        {
                            matchedTable = table;
                            break;
                        }
                    }

                    if (matchedTable == null)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "No matching project-info table found.");
                        DA.SetData(0, writtenPath);
                        DA.SetData(1, success);
                        return;
                    }

                    void UpdateKey(string key, string value)
                    {
                        var row = matchedTable.Elements<TableRow>()
                                               .FirstOrDefault(r => r.Elements<TableCell>()
                                                                     .FirstOrDefault() != null &&
                                                                     GetCellPlainText(r.Elements<TableCell>().First())
                                                                         .Trim().ToLowerInvariant()
                                                                         .Contains(key.ToLowerInvariant()));
                        if (row == null) return;

                        var cells = row.Elements<TableCell>().ToList();
                        if (cells.Count < 2) return;

                        SetCellText(cells[1], value ?? string.Empty);
                    }

                    UpdateKey("projectname", projectName);
                    UpdateKey("client", client);
                    UpdateKey("date", dateText);
                    UpdateKey("version", version);
                    UpdateKey("author", author);
                    UpdateKey("checked by", checkedBy);

                    doc.MainDocumentPart.Document.Save();
                }

                writtenPath = docPath;
                success = true;
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"Successfully updated project info: {writtenPath}");
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Error: {ex.Message}");
            }

            DA.SetData(0, writtenPath);
            DA.SetData(1, success);
        }

        static string GetCellPlainText(TableCell cell)
        {
            return string.Concat(cell.Descendants<Text>().Select(t => t.Text));
        }

        static void SetCellText(TableCell cell, string text)
        {
            cell.RemoveAllChildren<Paragraph>();
            var para = new Paragraph(new Run(new Text(text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve }));
            cell.AppendChild(para);
        }

        public override GH_Exposure Exposure => GH_Exposure.primary;

        protected override System.Drawing.Bitmap Icon => null;

        public override Guid ComponentGuid => new Guid("9a9a2bde-3d2f-4b6b-9f9a-2d4b8d8e7f02");
    }
}
