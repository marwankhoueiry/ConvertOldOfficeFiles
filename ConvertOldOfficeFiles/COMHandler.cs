using System;
using NetOffice.WordApi.Enums;

namespace ConvertOldOfficeFiles
{
    class COMHandler: IDisposable
    {
        private COMHandler()
        {
            ExcelApplication = new NetOffice.ExcelApi.Application
            {
                Visible = false,
                DisplayAlerts = false
            };

            WordApplication = new NetOffice.WordApi.Application
            {
                Visible = false,
                DisplayAlerts = WdAlertLevel.wdAlertsNone
            };
        }

        private static NetOffice.ExcelApi.Application ExcelApplication { get; set; }

        private static NetOffice.WordApi.Application WordApplication { get; set; }

        /// <summary>
        /// Disposes all COM objects instantiated by this class
        /// </summary>
        public void Dispose()
        {
            ExcelApplication.Quit();
            ExcelApplication.Dispose();
            WordApplication.Quit();
            WordApplication.Dispose();
        }

        /// <summary>
        /// Opens a MS Word Document
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns> NetOffice.WordApi.Document</returns>
        public static NetOffice.WordApi.Document OpenWordDocument(string fileName)
        {
            return WordApplication.Documents.Open(fileName);
        }


        /// <summary>
        /// Opens a MS Excel Workbook
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns> NetOffice.ExcelApi.Workbook</returns>
        public static NetOffice.ExcelApi.Workbook OpenExcelDocument(string fileName)
        {
            return ExcelApplication.Workbooks.Open(fileName);
        }
    }
}