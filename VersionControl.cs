using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Globalization;

using Grasshopper;
using Grasshopper.Kernel;
using Rhino.Geometry;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace BalkroosterExport
{
    public class VersionControl : GH_Component
    {
        public VersionControl()
          : base("VersionControl", "VersionControl",
            "Replace table data (except header) and insert version rows into the first matching .docx table.",
            "Document", "Export")
        {
        }

        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("DocPath", "Doc", "Path to working .docx file", GH_ParamAccess.item);
            pManager.AddTextParameter("Revisie", "Revisie", "Revisie strings (e.g. A). At least one required.", GH_ParamAccess.list);
            pManager.AddTextParameter("Datum", "Datum", "Data voor de rijen (yyyy-MM-dd). Indien leeg wordt 'nu' gebruikt.", GH_ParamAccess.list);
            pManager.AddTextParameter("Status", "Status", "Status voor de revisies", GH_ParamAccess.list);
            pManager.AddTextParameter("Toelichting", "Toelichting", "Toelichting bij de revisies", GH_ParamAccess.list);
            pManager.AddTextParameter("Opsteller", "Opsteller", "Naam van de opsteller", GH_ParamAccess.list);
            pManager.AddTextParameter("Controleur", "Controleur", "Naam van de controleur", GH_ParamAccess.list);
            pManager.AddBooleanParameter("Add", "Add", "When true replace table data and add the provided rows.", GH_ParamAccess.item, false);
        }

        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("FilePath", "File", "Path to the written file", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Success", "Ok", "True when write succeeded", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string docPath = string.Empty;
            var revisionInputs = new List<string>();
            var dateInputs = new List<string>();
            var statusInputs = new List<string>();
            var commentInputs = new List<string>();
            var authorInputs = new List<string>();
            var checkerInputs = new List<string>();
            bool add = false;

            if (!DA.GetData(0, ref docPath)) return;
            DA.GetDataList(1, revisionInputs);
            DA.GetDataList(2, dateInputs);
            DA.GetDataList(3, statusInputs);
            DA.GetDataList(4, commentInputs);
            DA.GetDataList(5, authorInputs);
            DA.GetDataList(6, checkerInputs);
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

            if (revisionInputs == null || revisionInputs.Count == 0)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Provide at least one Revisie entry.");
                DA.SetData(0, writtenPath);
                DA.SetData(1, success);
                return;
            }

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
                        var firstRow = table.Elements<TableRow>().FirstOrDefault();
                        if (firstRow == null) continue;

                        var headerCells = firstRow.Elements<TableCell>().Select(c => GetCellPlainText(c).Trim().ToLowerInvariant()).ToArray();
                        bool hasRevision = headerCells.Any(h => h.Contains("revisie"));
                        bool hasDate = headerCells.Any(h => h.Contains("datum"));
                        bool hasStatus = headerCells.Any(h => h.Contains("status"));
                        bool hasComment = headerCells.Any(h => h.Contains("toelichting"));
                        bool hasAuthor = headerCells.Any(h => h.Contains("opsteller"));
                        bool hasChecker = headerCells.Any(h => h.Contains("controleur"));

                        if (hasRevision && hasDate && hasStatus && hasComment && hasAuthor && hasChecker)
                        {
                            matchedTable = table;
                            break;
                        }
                    }

                    if (matchedTable == null)
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "No matching table found with Dutch headers: revisie, datum, status, toelichting, opsteller, controleur.");
                        DA.SetData(0, writtenPath);
                        DA.SetData(1, success);
                        return;
                    }

                    var allRows = matchedTable.Elements<TableRow>().ToList();
                    if (allRows.Count > 1)
                    {
                        foreach (var row in allRows.Skip(1).ToList())
                        {
                            row.Remove();
                        }
                    }

                    int rowCount = revisionInputs.Count;

                    for (int i = 0; i < rowCount; i++)
                    {
                        string revText = (i < revisionInputs.Count) ? revisionInputs[i] ?? string.Empty : string.Empty;
                        
                        string dateInput = (i < dateInputs.Count) ? dateInputs[i] : string.Empty;
                        DateTime dateValue;
                        if (string.IsNullOrWhiteSpace(dateInput))
                        {
                            dateValue = DateTime.Now;
                        }
                        else
                        {
                            if (!DateTime.TryParse(dateInput, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out dateValue))
                            {
                                if (!DateTime.TryParse(dateInput, out dateValue))
                                {
                                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"Datum '{dateInput}' op index {i} kon niet worden verwerkt. Huidige datum wordt gebruikt.");
                                    dateValue = DateTime.Now;
                                }
                            }
                        }
                        string dateText = dateValue.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

                        string statusText = (i < statusInputs.Count) ? statusInputs[i] ?? string.Empty : string.Empty;
                        string commentText = (i < commentInputs.Count) ? commentInputs[i] ?? string.Empty : string.Empty;
                        string authorText = (i < authorInputs.Count) ? authorInputs[i] ?? string.Empty : string.Empty;
                        string checkerText = (i < checkerInputs.Count) ? checkerInputs[i] ?? string.Empty : string.Empty;

                        var newRow = new TableRow();
                        newRow.Append(CreateTextCell(revText));
                        newRow.Append(CreateTextCell(dateText));
                        newRow.Append(CreateTextCell(statusText));
                        newRow.Append(CreateTextCell(commentText));
                        newRow.Append(CreateTextCell(authorText));
                        newRow.Append(CreateTextCell(checkerText));

                        matchedTable.AppendChild(newRow);
                    }

                    doc.MainDocumentPart.Document.Save();
                }

                writtenPath = docPath;
                success = true;
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"Successfully updated document: {writtenPath}");
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

        static TableCell CreateTextCell(string text)
        {
            var tc = new TableCell(
                new TableCellProperties(
                    new TableCellWidth() { Type = TableWidthUnitValues.Auto }
                ),
                new Paragraph(new Run(new Text(text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve }))
            );
            return tc;
        }

        public override GH_Exposure Exposure => GH_Exposure.primary;

        protected override System.Drawing.Bitmap Icon => null;

        public override Guid ComponentGuid => new Guid("b0c2a5b8-4b5e-4e6f-8b3c-9f2d5a7e6c01");
    }
}
