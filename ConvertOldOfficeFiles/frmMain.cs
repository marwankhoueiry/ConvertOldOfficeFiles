using System;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace ConvertOldOfficeFiles
{
    public partial class frmMain : Form
    {
        private Excel.Application excel;
        private Word.Application word;
        private int fileCount = 0;
        public frmMain()
        {
            InitializeComponent();
            
            // Create window title
            this.Text = Assembly.GetExecutingAssembly().GetName().Name + " Version " + Assembly.GetExecutingAssembly().GetName().Version;
            
            // Create Excel COM object instance
            excel = new Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;

            // Create Word COM object instance
            word = new Word.Application();
            word.Visible = false;
            word.DisplayAlerts = WdAlertLevel.wdAlertsNone;
        }

        private void ConvertPath (string path, bool bConvert)
        {
            try
            {
                // Search for Excel files with an old office format and convert them into the new office OpenXML format
                string[] fileNames = Directory.GetFiles(path, "*.xls");
                statusLabel.Text = path;
                Application.DoEvents();
                Cursor.Current = Cursors.WaitCursor;
                foreach (string fileName in fileNames)
                {
                    string ext = Path.GetExtension(fileName);
                    if (Path.GetExtension(fileName) == ".xls")
                    {
                        // Check if the file is a file with Office 2003 format (header check)
                        if (!IsOldOfficeFormat(fileName))
                        {
                            tbOutput.AppendText("Error: the file " + fileName + " has a wrong format and therefore will not be converted !" + Environment.NewLine);
                            continue;
                        }

                        if (bConvert)
                            ConvertXLS(fileName);
                        else
                        {
                            tbOutput.AppendText(fileName + Environment.NewLine);
                            fileCount++;
                        }
                    }
                }

                // Search for Word files with an old office format and convert them into the new office OpenXML format
                fileNames = Directory.GetFiles(path, "*.doc");
                foreach (string fileName in fileNames)
                {
                    if (Path.GetExtension(fileName) == ".doc")
                    {
                        // Check if the file is a file with Office 2003 format (header check)
                        if (!IsOldOfficeFormat(fileName))
                        {
                            tbOutput.AppendText("Error: the file " + fileName + " has a wrong format and therefore will not be converted !" + Environment.NewLine);
                            continue;
                        }

                        if (bConvert)
                            ConvertDOC(fileName);
                        else
                        {
                            tbOutput.AppendText(fileName + Environment.NewLine);
                            fileCount++;
                        }
                    }
                }

                // Now we have to search the sub dirs recursively
                string[] dirs = Directory.GetDirectories(path);
                foreach (string dir in dirs)
                    ConvertPath(dir, bConvert);
            }
            catch { }
        }

        private void ConvertXLS(string fileName)
        {
            string saveFileName = fileName.Replace(".xls",".xlsx");
            
            try
            {
                // Load Excel worksheet
                Excel.Workbook wb = excel.Workbooks.Open(fileName);

                try
                {
                    // Read author
                    NetOffice.OfficeApi.DocumentProperties properties = (NetOffice.OfficeApi.DocumentProperties)wb.BuiltinDocumentProperties;
                    string author = "";
                    foreach (NetOffice.OfficeApi.DocumentProperty p in properties)
                        if (p.Name == "Author")
                            author = p.Value.ToString();
                    //MessageBox.Show(author);

                    // Check if the file contains macro code
                    int linesOfCode = 0;
                    foreach (NetOffice.VBIDEApi.VBComponent component in wb.VBProject.VBComponents)
                        linesOfCode += component.CodeModule.CountOfLines;

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
                    fileCount++;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                // Cleanup
                wb.Close();
                wb.DisposeChildInstances();

                // Reset the timestamp for "date modified" for the new created file
                FileInfo fi = new FileInfo(fileName);
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
        private void ConvertDOC(string fileName)
        {
            string saveFileName = fileName.Replace(".doc", ".docx");

            try
            {
                // Load Word document
                Word.Document doc = word.Documents.Open(fileName);

                try
                {
                    // Check if the file contains macro code
                    int linesOfCode = 0;
                    foreach (NetOffice.VBIDEApi.VBComponent component in doc.VBProject.VBComponents)
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
                    fileCount++;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                // Cleanup
                doc.Close();
                doc.DisposeChildInstances();

                // Reset the timestamp for "date modified" for the new created file
                FileInfo fi = new FileInfo(fileName);
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
            string path = tbPath.Text.Trim();
            if (path.Length > 0 && Directory.Exists(path))
            {
                tbOutput.Clear();
                fileCount = 0;
                ConvertPath(path, true);
                statusLabel.Text = "Ready";
                Cursor.Current = Cursors.Default;
                tbOutput.AppendText(fileCount + " files converted" + Environment.NewLine);
            }
        }

        private void btCheck_Click(object sender, EventArgs e)
        {
            string path = tbPath.Text.Trim();
            if (path.Length > 0 && Directory.Exists(path))
            {
                tbOutput.Clear();
                fileCount = 0;
                ConvertPath(path, false);
                statusLabel.Text = "Ready";
                Cursor.Current = Cursors.Default;
                tbOutput.AppendText(fileCount + " files found" + Environment.NewLine);
            }
        }

        private void frmMain_Shown(object sender, EventArgs e)
        {
            tbPath.Focus();
        }

        private void btSelectPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.Description = "Select directory to be converted";
            if (dlg.ShowDialog() == DialogResult.OK)
                tbPath.Text = dlg.SelectedPath;
        }

        private bool IsOldOfficeFormat (string fileName)
        {
            bool bIsOldFormat = true;
            try
            {
                // Header check, see https://www.loc.gov/preservation/digital/formats/fdd/fdd000509.shtml
                byte[] header = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

                // We're reading the first 512 Bytes of the file
                byte[] buffer = new byte[512];
                using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    // If the file does not contain at least the header it can't be a file in an old Office 2003 format
                    if (fs.Read(buffer, 0, buffer.Length) < header.Length)
                        return false;

                    // Check if the files begins with an Office 2003 header
                    for (int i = 0; i < header.Length; i++)
                        if (buffer[i] != header[i])
                            bIsOldFormat = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return bIsOldFormat;
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Clear Excel COM object
            excel.Quit();
            excel.Dispose();

            // Clear Word COM object
            word.Quit();
            word.Dispose();
        }
    }
}
