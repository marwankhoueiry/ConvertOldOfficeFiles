using System;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using Excel = NetOffice.ExcelApi;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace ConvertOldOfficeFiles
{
    public partial class frmMain : Form
    {
        private Excel.Application excel = null;
        private Word.Application word = null; 
        private int fileCount = 0;
        private bool bExit = false;
        public frmMain()
        {
            InitializeComponent();
            
            // Create window title
            this.Text = Assembly.GetExecutingAssembly().GetName().Name + " Version " + Assembly.GetExecutingAssembly().GetName().Version;
            
            try
            {
                // Create Excel COM object instance
                excel = new Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;

                try
                {
                    // Create Word COM object instance
                    word = new Word.Application();
                    word.Visible = false;
                    word.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                }
                catch { bExit = true; MessageBox.Show("Could not start a Word object instance on this computer", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
            catch { bExit = true; MessageBox.Show("Could not start an Excel object instance on this computer", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void ConvertPath (string path, bool bConvert)
        {

            string donePath = path + "\\Done\\";
            string failedPath = path + "\\Failed\\";
            if (checkFolders(donePath, failedPath) != true)
            {
                var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                tbOutput.AppendText(date + " Error: Unable to create folder: " + donePath + " Or Folder: " + failedPath + ".  Existing..." + Environment.NewLine);
                return;
            }

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
                    var t = Path.GetFileNameWithoutExtension(fileName);
                    var s = Path.GetFileName(fileName);
                    if (Path.GetExtension(fileName).ToLower() == ".xls")
                    {
                        // Check if the file is a file with Office 2003 format (header check)
                        if (!IsOldOfficeFormat(fileName))
                        {
                            var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            tbOutput.AppendText(date + "Error: the file " + fileName + " has a wrong format and therefore will not be converted !" + Environment.NewLine);
                            continue;
                        }

                        if (bConvert) {
                            string newPath = "C:\\OpenText\\Upload\\";
                            var newFilename = newPath + Path.GetFileNameWithoutExtension(fileName) + ".xlsx";
                            ConvertXLS(fileName, newFilename);
                        } 
                        else
                        {
                            var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            tbOutput.AppendText(date + ' ' + fileName + Environment.NewLine);
                            fileCount++;
                        }
                    }
                }

                // Search for Word files with an old office format and convert them into the new office OpenXML format
                fileNames = Directory.GetFiles(path, "*.doc");
                foreach (string fileName in fileNames)
                {
                  
                    if (Path.GetExtension(fileName).ToLower() == ".doc")
                    {
                        // Check if the file is a file with Office 2003 format (header check)
                        if (!IsOldOfficeFormat(fileName))
                        {
                            var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            tbOutput.AppendText(date + " Error: the file " + fileName + " has a wrong format and therefore will not be converted !" + Environment.NewLine);
                            continue;
                        }

                        if (bConvert) { 
                            string newPath = "C:\\OpenText\\Upload\\";
                            var newFilename = newPath + Path.GetFileNameWithoutExtension(fileName) + ".docx";
                            Boolean success = ConvertDOC(fileName, newFilename);
                      
                            if (success)
                            {
                                if (File.Exists(donePath + Path.GetFileName(fileName)) == false){
                                    File.Move(fileName, donePath + Path.GetFileName(fileName));
                                }
                                else
                                {
                                    File.Delete(fileName);  //Already proccessed - Get rid
                                }

                            } 
                            else
                            {
                                if (File.Exists(failedPath + Path.GetFileName(fileName)) == false)
                                {
                                    File.Move(fileName, failedPath + Path.GetFileName(fileName));
                                }
                                else
                                {
                                    File.Delete(fileName);  //Already proccessed - Get rid
                                }
                            }
                        }
                        else
                        {
                            var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            tbOutput.AppendText(date + ' ' + fileName + Environment.NewLine);
                            fileCount++;
                        }
                    }
                }
                string[] dirs;

                if (chkIncludeSubFolders.Checked)
                {
                    // Now we have to search the sub dirs recursively
                    dirs = Directory.GetDirectories(path);
                    foreach (string dir in dirs)
                        ConvertPath(dir, bConvert);

                }
            }
            catch { }
        }
        private Boolean checkFolders(string donePath, string  failedPath)
        {
            try
            {
                // Determine whether the directory exists.
                if (Directory.Exists(donePath) == false)
                {
                    Directory.CreateDirectory(donePath);
                }
                if (Directory.Exists(failedPath) == false)
                {
                    Directory.CreateDirectory(failedPath);
                }

            }
            catch (Exception e)
            {
                var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                tbOutput.AppendText(date + " Error: " + e.Message + Environment.NewLine);
                return false;
            }
            finally { }
            return true;
        }
        private void ConvertXLS(string fileName, string newFileName)
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
                            if (p != null) { 
                                author = p.Value.ToString();
                            }
                    //MessageBox.Show(author);

                    // Check if the file contains macro code
                    int linesOfCode = 0;
                    try
                    {
                        foreach (NetOffice.VBIDEApi.VBComponent component in wb.VBProject.VBComponents)
                            linesOfCode += component.CodeModule.CountOfLines;
                    }
                    catch
                    {
                        // Access to VBA object model ist not trusted, see https://support.microsoft.com/en-us/topic/programmatic-access-to-office-vba-project-is-denied-960d5265-6592-9400-31bc-b2ddfb94b445
                        var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        tbOutput.AppendText(date + " Error converting " + fileName + ": please enable access to the VBA object model within Excel (see https://support.microsoft.com/en-us/topic/programmatic-access-to-office-vba-project-is-denied-960d5265-6592-9400-31bc-b2ddfb94b445)" + Environment.NewLine);
                        return;
                    }

                    // A file containing macros must have a different target format / file extension
                    if (linesOfCode > 0)
                    {

                        saveFileName = newFileName.Replace(".xlsx", ".xlsm");
                        var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        tbOutput.AppendText(date + " Convert " + fileName + " to " + saveFileName + Environment.NewLine);
                        // Save in OpenXML format with macros (see https://docs.microsoft.com/de-de/office/vba/api/excel.xlfileformat)
                        wb.SaveAs(saveFileName, 52);
                    }
                    else
                    {
                        var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        tbOutput.AppendText(date + " Convert " + fileName + " to " + saveFileName + Environment.NewLine);
                        // Save in OpenXML format without macros  (see https://docs.microsoft.com/de-de/office/vba/api/excel.xlfileformat)

                        //wb.SaveAs(saveFileName, 51);
                        wb.SaveAs(newFileName, 51);
                    }
                    fileCount++;
                }
                catch (Exception ex)
                {
                    var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    tbOutput.AppendText(date + ex.Message);
                }

                // Cleanup
                wb.Close();
                wb.DisposeChildInstances();

                // Reset the timestamp for "date modified" for the new created file
                FileInfo fi = new FileInfo(fileName);
                File.SetLastWriteTime(newFileName, fi.LastWriteTime);
                File.SetCreationTime(newFileName, fi.CreationTime);

                // Delete the source file
                //File.Delete(fileName);
                File.Move(fileName, Path.GetDirectoryName(fileName) + "\\Done\\");
            }
            catch (Exception ex)
            {
                var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                tbOutput.AppendText(date + ' ' + ex.Message + Environment.NewLine);
                //MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private Boolean ConvertDOC(string fileName, string newFileName) 
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
                    try
                    {
                        foreach (NetOffice.VBIDEApi.VBComponent component in doc.VBProject.VBComponents)
                            linesOfCode += component.CodeModule.CountOfLines;
                    }
                    catch
                    {
                        // Access to VBA object model ist not trusted, see https://support.microsoft.com/en-us/topic/programmatic-access-to-office-vba-project-is-denied-960d5265-6592-9400-31bc-b2ddfb94b445
                        var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        tbOutput.AppendText(date + " Error converting " + fileName + ": please enable access to the VBA object model within Word (see https://support.microsoft.com/en-us/topic/programmatic-access-to-office-vba-project-is-denied-960d5265-6592-9400-31bc-b2ddfb94b445)" + Environment.NewLine);
                        return false;
                    }

                    // A file containing macros must have a different target format / file extension
                    if (linesOfCode > 0)
                    {
                        saveFileName = newFileName.Replace(".docx", ".docm");

                        var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        tbOutput.AppendText(date + " Convert " + fileName + " to " + saveFileName + Environment.NewLine);
                        // Save in OpenXML format with macros (see https://docs.microsoft.com/de-de/office/vba/api/word.wdsaveformat)
                        doc.SaveAs2(saveFileName, 13);
                    }
                    else
                    {
                        var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        tbOutput.AppendText(date + " Convert " + fileName + " to " + saveFileName + Environment.NewLine);
                        // Save in OpenXML format without macros (see https://docs.microsoft.com/de-de/office/vba/api/word.wdsaveformat)
                        saveFileName = newFileName;
                        doc.SaveAs2(saveFileName, 16);
                    }
                    fileCount++;
                }
                catch (Exception ex)
                {
                    var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    tbOutput.AppendText(date + " " + ex.Message + Environment.NewLine);
                    //MessageBox.Show(ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                // Cleanup
                doc.Close();
                doc.DisposeChildInstances();

                // Reset the timestamp for "date modified" for the new created file
                FileInfo fi = new FileInfo(fileName);
                File.SetLastWriteTime(saveFileName, fi.LastWriteTime);
                File.SetCreationTime(saveFileName, fi.CreationTime);

 
                return true;
            }
            catch (Exception ex)
            {
                var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                tbOutput.AppendText(date + ' ' + ex.Message + Environment.NewLine);
                return false;
                //MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                tbOutput.AppendText(date + ' ' + fileCount + " files converted" + Environment.NewLine);
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
                var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                tbOutput.AppendText(date + ' ' + fileCount + " files found" + Environment.NewLine);
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
                var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                tbOutput.AppendText(date + ' ' + ex.Message + Environment.NewLine);

                //MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return bIsOldFormat;
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Clear Excel COM object
            if (excel != null)
            {
                try
                {
                    excel.Quit();
                    excel.Dispose();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message, "Error while disposing Excel object instance", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }

            // Clear Word COM object
            if (word != null)
            {
                try
                {
                    word.Quit();
                    word.Dispose();
                }
                catch (Exception ex) { MessageBox.Show(ex.Message, "Error while disposing Word object instance", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
        }

        private void tbOutput_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
