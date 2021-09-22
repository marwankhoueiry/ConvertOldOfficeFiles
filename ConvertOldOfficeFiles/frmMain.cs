using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ConvertOldOfficeFiles
{
    public partial class FrmMain : Form
    {
        private int _fileCount;

        public FrmMain()
        {
            InitializeComponent();
        }

        public sealed override string Text { get; set; } =
            Application.ProductName + " Version " + Application.ProductVersion;

        private void ConvertPath(string path, bool bConvert)
        {
            try
            {
                // Search for Excel files with an old office format and convert them into the new office OpenXML format
                var fileNames = Directory.GetFiles(path, "*.xls");
                statusLabel.Text = path;
                System.Windows.Forms.Application.DoEvents();
                Cursor.Current = Cursors.WaitCursor;
                foreach (var fileName in fileNames)
                {
                    var ext = Path.GetExtension(fileName);
                    if (Path.GetExtension(fileName) != ".xls") continue;
                    // Check if the file is a file with Office 2003 format (header check)
                    if (!IsOldOfficeFormat(fileName))
                    {
                        tbOutput.AppendText("Error: the file " + fileName +
                                            " has a wrong format and therefore will not be converted !" +
                                            Environment.NewLine);
                        continue;
                    }

                    if (bConvert)
                    {
                        ConvertXls(fileName);
                    }
                    else
                    {
                        tbOutput.AppendText(fileName + Environment.NewLine);
                        _fileCount++;
                    }
                }

                // Search for Word files with an old office format and convert them into the new office OpenXML format
                fileNames = Directory.GetFiles(path, "*.doc");
                foreach (var fileName in fileNames)
                    if (Path.GetExtension(fileName) == ".doc")
                    {
                        // Check if the file is a file with Office 2003 format (header check)
                        if (!IsOldOfficeFormat(fileName))
                        {
                            tbOutput.AppendText("Error: the file " + fileName +
                                                " has a wrong format and therefore will not be converted !" +
                                                Environment.NewLine);
                            continue;
                        }

                        if (bConvert)
                        {
                            ConvertDoc(fileName);
                        }
                        else
                        {
                            tbOutput.AppendText(fileName + Environment.NewLine);
                            _fileCount++;
                        }
                    }

                // Now we have to search the sub dirs recursively
                var dirs = Directory.GetDirectories(path);
                foreach (var dir in dirs)
                    ConvertPath(dir, bConvert);
            }
            catch
            {
            }
        }

        private void ConvertXls(string fileName)
        {
            var saveFileName = fileName.Replace(".xls", ".xlsx");

            try
            {
                // Load Excel worksheet
                var wb = COMHandler.OpenExcelDocument(fileName);

                try
                {
                    // Check if the file contains macro code
                    var linesOfCode = wb.VBProject.VBComponents.Sum(component => component.CodeModule.CountOfLines);

                    // A file containing macros must have a different target format / file extension
                    if (linesOfCode > 0)
                    {
                        saveFileName = fileName.Replace(".xls", ".xlsm");
                        tbOutput.AppendText("Convert " + fileName + " to " + saveFileName + Environment.NewLine);
                        // Save in OpenXML format with macros (see https://docs.microsoft.com/de-de/office/vba/api/excel.xlfileformat)
                        wb.SaveAs(saveFileName, 52);
                    }
                    else
                    {
                        tbOutput.AppendText("Convert " + fileName + " to " + saveFileName + Environment.NewLine);
                        // Save in OpenXML format without macros  (see https://docs.microsoft.com/de-de/office/vba/api/excel.xlfileformat)
                        wb.SaveAs(saveFileName, 51);
                    }

                    _fileCount++;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                // Cleanup
                wb.Close();
                wb.DisposeChildInstances();

                // Reset the timestamp for "date modified" for the new created file
                var fi = new FileInfo(fileName);
                File.SetLastWriteTime(saveFileName, fi.LastWriteTime);
                File.SetCreationTime(saveFileName, fi.CreationTime);

                // Delete the source file
                File.Delete(fileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ConvertDoc(string fileName)
        {
            var saveFileName = fileName.Replace(".doc", ".docx");

            try
            {
                // Load Word document
                var doc = COMHandler.OpenWordDocument(fileName);

                try
                {
                    // Check if the file contains macro code
                    var linesOfCode = 0;
                    foreach (var component in doc.VBProject.VBComponents)
                        linesOfCode += component.CodeModule.CountOfLines;

                    // A file containing macros must have a different target format / file extension
                    if (linesOfCode > 0)
                    {
                        saveFileName = fileName.Replace(".doc", ".docm");
                        tbOutput.AppendText("Convert " + fileName + " to " + saveFileName + Environment.NewLine);
                        // Save in OpenXML format with macros (see https://docs.microsoft.com/de-de/office/vba/api/word.wdsaveformat)
                        doc.SaveAs2(saveFileName, 13);
                    }
                    else
                    {
                        tbOutput.AppendText("Convert " + fileName + " to " + saveFileName + Environment.NewLine);
                        // Save in OpenXML format without macros (see https://docs.microsoft.com/de-de/office/vba/api/word.wdsaveformat)
                        doc.SaveAs2(saveFileName, 16);
                    }

                    _fileCount++;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                // Cleanup
                doc.Close();
                doc.DisposeChildInstances();

                // Reset the timestamp for "date modified" for the new created file
                var fi = new FileInfo(fileName);
                File.SetLastWriteTime(saveFileName, fi.LastWriteTime);
                File.SetCreationTime(saveFileName, fi.CreationTime);

                // Delete the source file
                File.Delete(fileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btConvert_Click(object sender, EventArgs e)
        {
            var path = tbPath.Text.Trim();
            if (path.Length <= 0 || !Directory.Exists(path)) return;
            
            tbOutput.Clear();
            _fileCount = 0;
            ConvertPath(path, true);
            statusLabel.Text = "Ready";
            Cursor.Current = Cursors.Default;
            tbOutput.AppendText(_fileCount + " files converted" + Environment.NewLine);
        }

        private void btCheck_Click(object sender, EventArgs e)
        {
            var path = tbPath.Text.Trim();
            if (path.Length <= 0 || !Directory.Exists(path)) return;
            
            tbOutput.Clear();
            _fileCount = 0;
            ConvertPath(path, false);
            statusLabel.Text = "Ready";
            Cursor.Current = Cursors.Default;
            tbOutput.AppendText(_fileCount + " files found" + Environment.NewLine);
        }

        private void frmMain_Shown(object sender, EventArgs e)
        {
            tbPath.Focus();
        }

        private void btSelectPath_Click(object sender, EventArgs e)
        {
            var dlg = new FolderBrowserDialog
            {
                Site = null,
                Tag = null,
                AutoUpgradeEnabled = false,
                ShowNewFolderButton = false,
                SelectedPath = null,
                RootFolder = Environment.SpecialFolder.Desktop,
                Description = null,
                UseDescriptionForTitle = false
            };
            dlg.Description = "Select directory to be converted";
            if (dlg.ShowDialog() == DialogResult.OK)
                tbPath.Text = dlg.SelectedPath;
        }

        private bool IsOldOfficeFormat(string fileName)
        {
            var bIsOldFormat = true;
            try
            {
                // Header check, see https://www.loc.gov/preservation/digital/formats/fdd/fdd000509.shtml
                byte[] header = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

                // We're reading the first 512 Bytes of the file
                var buffer = new byte[512];
                using var fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                // If the file does not contain at least the header it can't be a file in an old Office 2003 format
                if (fs.Read(buffer, 0, buffer.Length) < header.Length)
                    return false;

                // Check if the files begins with an Office 2003 header
                for (var i = 0; i < header.Length; i++)
                    if (buffer[i] != header[i])
                        bIsOldFormat = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return bIsOldFormat;
        }
    }
}