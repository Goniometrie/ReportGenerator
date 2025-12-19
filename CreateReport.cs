using System;
using System.IO;
using System.Globalization;

using Grasshopper;
using Grasshopper.Kernel;

namespace BalkroosterExport
{
    public class CreateReport : GH_Component
    {
        /// <summary>
        /// Central component that creates a working copy of a .docx template and exposes
        /// the working path for other components to consume. Template is never overwritten.
        /// Inputs:
        /// - TemplatePath     (string) : required .docx template path
        /// - OutputDirectory  (string) : optional; if empty copy is created next to the Grasshopper file
        /// - Filename         (string) : optional; if empty copy uses template name (ensures .docx)
        /// - Create           (bool)   : when true create the working copy
        /// Outputs:
        /// - WorkingPath      (string) : path to the copy (or output path) to be used by downstream components
        /// - TemplatePath     (string) : original template path (echo)
        /// - Success          (bool)   : true when copy was created (or output path validated)
        /// </summary>
        public CreateReport()
          : base("CreateReport", "CreateReport",
            "Create a working copy of a .docx template and expose the working path for other components.",
            "Document", "Utility")
        { }

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("TemplatePath", "Tmpl", "Full path to .docx template", GH_ParamAccess.item);
            pManager.AddTextParameter("OutputDirectory", "Dir", "Optional output directory. If empty, the Grasshopper file location or template location is used.", GH_ParamAccess.item, string.Empty);
            pManager.AddTextParameter("Filename", "File", "Optional filename for the copy. If empty, the template filename is used.", GH_ParamAccess.item, string.Empty);
            pManager.AddBooleanParameter("Create", "Create", "When true create the working copy (template is never overwritten).", GH_ParamAccess.item, false);

            pManager[0].Optional = false;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("WorkingPath", "Work", "Path to the working .docx file for downstream components", GH_ParamAccess.item);
            pManager.AddTextParameter("TemplatePath", "Tmpl", "Original template path (echo)", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Success", "Ok", "True when working file is available", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string templatePath = string.Empty;
            string outputDirectory = string.Empty;
            string filename = string.Empty;
            bool create = false;

            if (!DA.GetData(0, ref templatePath)) return;
            DA.GetData(1, ref outputDirectory);
            DA.GetData(2, ref filename);
            DA.GetData(3, ref create);

            bool success = false;
            string workingPath = string.Empty;

            if (string.IsNullOrWhiteSpace(templatePath))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "TemplatePath is empty.");
                DA.SetData(0, workingPath);
                DA.SetData(1, templatePath);
                DA.SetData(2, success);
                return;
            }

            if (!File.Exists(templatePath))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Template not found: {templatePath}");
                DA.SetData(0, workingPath);
                DA.SetData(1, templatePath);
                DA.SetData(2, success);
                return;
            }

            try
            {
                // 1. Determine target directory
                string targetDir = string.Empty;
                if (!string.IsNullOrWhiteSpace(outputDirectory))
                {
                    targetDir = outputDirectory;
                }
                else
                {
                    string ghFilePath = OnPingDocument()?.FilePath;
                    if (!string.IsNullOrWhiteSpace(ghFilePath))
                    {
                        targetDir = Path.GetDirectoryName(ghFilePath);
                    }
                    else
                    {
                        targetDir = Path.GetDirectoryName(Path.GetFullPath(templatePath));
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Grasshopper file is not saved. Copy will be created in template folder.");
                    }
                }

                // 2. Determine filename
                string finalFilename = filename;
                if (string.IsNullOrWhiteSpace(finalFilename))
                {
                    finalFilename = Path.GetFileName(templatePath);
                }
                else
                {
                    // Automatically add .docx if missing (case-insensitive check)
                    if (!finalFilename.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                    {
                        finalFilename += ".docx";
                    }
                }

                // 3. Resolve candidate path
                string candidatePath = Path.Combine(targetDir, finalFilename);

                // Verify candidate path is valid and not equal to template
                try
                {
                    string fullCandidate = Path.GetFullPath(candidatePath);
                    string fullTemplate = Path.GetFullPath(templatePath);

                    if (string.Equals(fullCandidate, fullTemplate, StringComparison.OrdinalIgnoreCase))
                    {
                        // If they are the same, we MUST change the name to avoid overwriting the template
                        string nameOnly = Path.GetFileNameWithoutExtension(candidatePath);
                        string ext = Path.GetExtension(candidatePath);
                        candidatePath = Path.Combine(targetDir, $"{nameOnly}_copy{ext}");
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Output path equals template path. Using a suffix to avoid overwriting template.");
                    }
                }
                catch
                {
                    // If Resolution fails, we'll hit an error later in CreateDirectory or Copy
                }

                if (!create)
                {
                    workingPath = candidatePath;
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Create is false — no copy created.");
                    DA.SetData(0, workingPath);
                    DA.SetData(1, templatePath);
                    DA.SetData(2, false);
                    return;
                }

                // Ensure directory exists
                if (!Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                }

                // If candidatePath already exists, add numeric suffix to avoid overwrite
                if (File.Exists(candidatePath))
                {
                    string nameOnly = Path.GetFileNameWithoutExtension(candidatePath);
                    string dir = Path.GetDirectoryName(candidatePath);
                    string ext = Path.GetExtension(candidatePath);
                    int i = 1;
                    string newPath;
                    do
                    {
                        newPath = Path.Combine(dir, $"{nameOnly}_{i}{ext}");
                        i++;
                    } while (File.Exists(newPath));
                    candidatePath = newPath;
                }

                // Copy template (never overwrite template)
                File.Copy(templatePath, candidatePath, false);

                workingPath = candidatePath;
                success = true;
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"Working copy created: {workingPath}");
            }
            catch (UnauthorizedAccessException uex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Access denied: {uex.Message}");
            }
            catch (IOException ioex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"IO error: {ioex.Message}");
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Unexpected error: {ex.Message}");
            }

            DA.SetData(0, workingPath);
            DA.SetData(1, templatePath);
            DA.SetData(2, success);
        }

        public override GH_Exposure Exposure => GH_Exposure.quarternary;

        protected override System.Drawing.Bitmap Icon => null;

        public override Guid ComponentGuid => new Guid("f7d2c3a7-6e4b-4b5a-9f1b-2c3d4e5f6a78");
    }
}
