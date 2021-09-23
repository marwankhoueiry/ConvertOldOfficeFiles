using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConvertOldOfficeFiles
{
    internal class Converter
    {
        public static int FileCount;
        public static string Output { get; set; }

        public static void ConvertPath(string path, bool bConvert)
        {
            // Reset file count
            FileCount = 0;

            try
            {
                // Search for Excel files with an old office format and convert them into the new office OpenXML format
                var fileNames = Directory.GetFiles(path, "*.xls");
                //TODO:statusLabel.Text = path;
                System.Windows.Forms.Application.DoEvents();
                Cursor.Current = Cursors.WaitCursor;
                foreach (var fileName in fileNames)
                {
                    var ext = Path.GetExtension(fileName);
                    if (Path.GetExtension(fileName) != ".xls") continue;
                    // Check if the file is a file with Office 2003 format (header check)
                    if (!IsOldOfficeFormat(fileName))
                    {
                        Output += "Error: the file " + fileName +
                                  " has a wrong format and therefore will not be converted !" +
                                  Environment.NewLine;
                        TextChanged?.Invoke(new object(), EventArgs.Empty);
                        continue;
                    }

                    if (bConvert)
                    {
                        ConvertXls(fileName);
                    }
                    else
                    {
                        Output += fileName + Environment.NewLine;
                        TextChanged?.Invoke(new object(), EventArgs.Empty);
                        FileCount++;
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
                            Output += "Error: the file " + fileName +
                                                " has a wrong format and therefore will not be converted !" +
                                                Environment.NewLine;
                            TextChanged?.Invoke(new object(), EventArgs.Empty);
                            continue;
                        }

                        if (bConvert)
                        {
                            ConvertDoc(fileName);
                        }
                        else
                        {
                            Output += fileName + Environment.NewLine;
                            TextChanged?.Invoke(new object(), EventArgs.Empty);
                            FileCount++;
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
        private static void ConvertXls(string fileName)
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
                        Output += "Convert " + fileName + " to " + saveFileName + Environment.NewLine;
                        TextChanged?.Invoke(new object(), EventArgs.Empty);
                        // Save in OpenXML format with macros (see https://docs.microsoft.com/de-de/office/vba/api/excel.xlfileformat)
                        wb.SaveAs(saveFileName, 52);
                    }
                    else
                    {
                        Output += "Convert " + fileName + " to " + saveFileName + Environment.NewLine;
                        TextChanged?.Invoke(new object(), EventArgs.Empty);
                        // Save in OpenXML format without macros  (see https://docs.microsoft.com/de-de/office/vba/api/excel.xlfileformat)
                        wb.SaveAs(saveFileName, 51);
                    }

                    FileCount++;
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

        private static void ConvertDoc(string fileName)
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
                        Output += "Convert " + fileName + " to " + saveFileName + Environment.NewLine;
                        TextChanged?.Invoke(new object(), EventArgs.Empty);
                        // Save in OpenXML format with macros (see https://docs.microsoft.com/de-de/office/vba/api/word.wdsaveformat)
                        doc.SaveAs2(saveFileName, 13);
                    }
                    else
                    {
                        Output += "Convert " + fileName + " to " + saveFileName + Environment.NewLine;
                        TextChanged?.Invoke(new object(), EventArgs.Empty);
                        // Save in OpenXML format without macros (see https://docs.microsoft.com/de-de/office/vba/api/word.wdsaveformat)
                        doc.SaveAs2(saveFileName, 16);
                    }

                    FileCount++;
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

        private static bool IsOldOfficeFormat(string fileName)
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

        public static event EventHandler TextChanged;

        public virtual void OnTextChanged(EventArgs e)
        {
            var handler = TextChanged;
            handler?.Invoke(this, e);
        }
    }
}
