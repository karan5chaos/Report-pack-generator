using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using Report_pack_generator.Settings;
using System.Data.SqlClient;
using iTextSharp.text.pdf;
using Microsoft.Office.Interop.Word;
using Report_pack_generator.Modules;

namespace Report_pack_generator.Functions
{
    static class get_Functions
    {


        public static bool rotate_Pdf(string inputfile)
        {
            var isError = false;
            try
            {
                Get_Status.staus_messages.Add("Rotating pages - " + Path.GetFileNameWithoutExtension(inputfile) + " ..");

                PdfReader reader = new PdfReader(File.ReadAllBytes(inputfile));
                PdfName pdfName = new PdfName(inputfile);

                int n = reader.NumberOfPages;
                int rot;
                PdfDictionary pageDict;

                
                for (int i = 1; i <= n; i++)
                {
                    rot = reader.GetPageRotation(i);
                    pageDict = reader.GetPageN(i);
                    pageDict.Put(PdfName.ROTATE, new PdfNumber(rot + 270));
                }
                Get_Status.staus_messages.Add(Path.GetFileNameWithoutExtension(inputfile) + " pages rotated.");

                PdfStamper stamper = new PdfStamper(reader, new FileStream(inputfile, FileMode.Open));
                stamper.Close();
                reader.Close();
            }

            catch(Exception exception)
            {
                isError = true;
                Get_Status.staus_messages.Add("Error occured in rotating pages for "+ Path.GetFileNameWithoutExtension(inputfile) + " ..");
                Get_Status.set_error(exception, inputfile, "PDF");
            }
            return isError;
        }

        public static bool rm_Nodata(string file, Excel.Application app)
        {
            bool isError = false;

            try
            {
                Get_Status.staus_messages.Add("Checking 'No Data' in - " + Path.GetFileNameWithoutExtension(file) + " ..");

                Excel.Workbook wkb = app.Workbooks.Open(file);

                app.DisplayAlerts = false;

                foreach (Excel.Worksheet sheet in wkb.Sheets)
                {
                    Excel.Range range = (Excel.Range)sheet.Columns["A:AA", Type.Missing];

                    Excel.Range resultRange = range.Find
                    (
                    What: "No Data",
                    LookIn: Excel.XlFindLookIn.xlValues,
                    LookAt: Excel.XlLookAt.xlPart,
                    SearchOrder: Excel.XlSearchOrder.xlByRows,
                    SearchDirection: Excel.XlSearchDirection.xlNext
                    );

                    if (resultRange != null)
                    {
                        sheet.Delete();
                    }
                }

                wkb.Save();

                GC.Collect();
                GC.WaitForPendingFinalizers();

                wkb.Close(Type.Missing, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(wkb);

                Get_Status.staus_messages.Add("Sheets with 'No Data' removed.");
            }
            catch(Exception exception)
            {
                isError = true;
                Get_Status.set_error(exception, file, "-");
            }

            return isError;
        }

        public static bool unhide_Columns(string file, Excel.Application app)
        {
            bool isError = false;
            if (excel_settings.Default.unhide_column)
            {

                try
                {
                    Get_Status.staus_messages.Add("Checking hidden columns in - " + Path.GetFileNameWithoutExtension(file) + " ..");

                    Excel.Workbook wkb = app.Workbooks.Open(file);


                    foreach (Excel.Worksheet ws in wkb.Sheets)
                    {
                        var cols = ws.Columns["A:AA"];

                        cols.Hidden = false;

                        foreach (var col in cols)
                        {
                            if (col.Width < 1)
                            {
                                col.Width = 14;

                            }
                        }
                        ws.Columns["D:F"].AutoFit();

                        ws.PageSetup.CenterFooter = "";
                        ws.PageSetup.RightFooter = "";
                        ws.PageSetup.LeftFooter = "";

                        if (excel_settings.Default.wrap_text)
                        {
                            ws.Range["D10:D2000"].WrapText = true;
                        }

                    }
                    wkb.Save();

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    wkb.Close(Type.Missing, Type.Missing, Type.Missing);
                    Marshal.FinalReleaseComObject(wkb);
                }

                catch(Exception exception)
                {
                    isError = true;
                    Get_Status.staus_messages.Add("Error occured while resizing hiddden columns..");
                    Get_Status.set_error(exception, file, "-");
                }
            }
            return isError;

        }

        public static bool warp_text(string file, Excel.Application app)
        {
            bool isError = false;
            if (excel_settings.Default.unhide_column)
            {

                try
                {
                    Get_Status.staus_messages.Add("Warp text in - " + Path.GetFileNameWithoutExtension(file) + " ..");

                    Excel.Workbook wkb = app.Workbooks.Open(file);


                    foreach (Excel.Worksheet ws in wkb.Sheets)
                    {
                        if (excel_settings.Default.wrap_text)
                        {
                            ws.Range["D10:D2000"].WrapText = true;
                        }
                    }
                    wkb.Save();

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    wkb.Close(Type.Missing, Type.Missing, Type.Missing);
                    Marshal.FinalReleaseComObject(wkb);

                    Get_Status.staus_messages.Add("Warp text enabled");

                }

                catch(Exception exception)
                {
                    isError = true;
                    Get_Status.staus_messages.Add("Warp text cannot be enabled..");
                }
            }
            return isError;

        }

        public static bool remove_NewRenewColumn(string file, Excel.Application app)
        {
            bool isError = false;
            if (excel_settings.Default.new_renew_column)
            {
                try
                {
                    Get_Status.staus_messages.Add("Checking for New/Renewal column in - " + Path.GetFileNameWithoutExtension(file) + " ..");

                    if (!file.Contains("Additional"))
                    {
                        Excel.Workbook wkb = app.Workbooks.Open(file);

                        foreach (Excel.Worksheet ws in wkb.Sheets)
                        {
                            ws.Columns[excel_settings.Default.borderux_column].Delete();
                        }
                        wkb.Save();

                        GC.Collect();
                        GC.WaitForPendingFinalizers();

                        wkb.Close(Type.Missing, Type.Missing, Type.Missing);
                        Marshal.FinalReleaseComObject(wkb);

                        Get_Status.staus_messages.Add("New/Renew column removed");
                    }
                }
                catch(Exception exception)
                {
                   isError= true;
                   Get_Status.staus_messages.Add("Error occured in removing New/Renew column..");
                   Get_Status.set_error(exception, file, "Borderaux files");
                }
            }
            return isError;
        }

        public static bool MergePDFs(List<string> fileNames, string targetPdf)
        {
            var merged = false;
            Get_Status.staus_messages.Add("Initiating report pack compilation..");

            if (fileNames.Count > 0)
            {
                using (FileStream stream = new FileStream(targetPdf, FileMode.Create))
                {
                    iTextSharp.text.Document document = new iTextSharp.text.Document();
                    PdfCopy pdf = new PdfCopy(document, stream);
                    PdfReader reader = null;
                    try
                    {
                        document.Open();
                        Get_Status.staus_messages.Add("Appending blannk pages where required..");
                        foreach (string file in fileNames)
                        {
                            reader = new PdfReader(file);
                            if (reader.NumberOfPages > 0)
                            {
                                pdf.AddDocument(reader);

                                //check and add blank pages.
                                //checks if number of pages in pdf file is odd and add pages accordingly.
                                if (reader.NumberOfPages % 2 != 0 && !file.Contains("CLIENT"))
                                {
                                    pdf.AddPage(reader.GetPageSize(1), reader.GetPageRotation(1));
                                }
                                reader.Close();
                            }
                        }
                        merged = true;

                        Get_Status.staus_messages.Add("Report pack compilation successful.");

                    }
                    catch(Exception exception)
                    {
                        merged = false;
                        if (reader != null)
                        {
                            reader.Close();
                        }
                        Get_Status.staus_messages.Add("Error occured while report pack compilation");
                        

                    }
                    finally
                    {
                        if (document != null)
                        {
                            document.Close();
                            fileNames.Clear();
                        }
                    }
                }
            }
            return merged;
        }

        public static string convert_ToPDF(int fileType, string inputFile, string outputFile, string folder_name, Excel.Application app)
        {
            string filename = null;
            Get_Status.staus_messages.Add("Converting "+Path.GetFileNameWithoutExtension(inputFile)+ " ..");

            if (folder.Default.covert_pdf)
            {
                switch (fileType)
                {

                    case 0:
                        {
                            try
                            {
                                app.Visible = false;
                                app.DisplayStatusBar = false;
                                app.ScreenUpdating = false;

                                Microsoft.Office.Interop.Excel.Workbook wkb = app.Workbooks.Open(inputFile, ReadOnly: true);
                                wkb.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outputFile);

                                GC.Collect();
                                GC.WaitForPendingFinalizers();

                                wkb.Close(Type.Missing, Type.Missing, Type.Missing);
                                Marshal.FinalReleaseComObject(wkb);

                                if (folder_name.Contains("Pipeline"))
                                {
                                    rotate_Pdf(outputFile);
                                }

                                if (folder_name.Contains("Analysis") || folder_name.Contains("Pipeline") || folder_name.Contains("Commentaries"))
                                {
                                    filename = outputFile;
                                }
                                Get_Status.staus_messages.Add("Converted " + Path.GetFileNameWithoutExtension(inputFile) + " successfully.");

                            }
                            catch(Exception exception)
                            {
                                Get_Status.staus_messages.Add(Path.GetFileNameWithoutExtension(inputFile) + " cannot be converted..");
                                filename = null;
                                Get_Status.set_error(exception,inputFile,folder_name);
                            }
                        }
                        break;

                    case 1:
                        {
                            try
                            {
                                Microsoft.Office.Interop.Word.Application wordapp = new Microsoft.Office.Interop.Word.Application();
                                wordapp.Visible = false;
                                wordapp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                                wordapp.ScreenUpdating = false;
                                wordapp.DisplayDocumentInformationPanel = false;

                                Microsoft.Office.Interop.Word.Document wkb = wordapp.Documents.Open(inputFile, ReadOnly: true);
                                wkb.ExportAsFixedFormat(outputFile, WdExportFormat.wdExportFormatPDF);

                                GC.Collect();
                                GC.WaitForPendingFinalizers();

                                wkb.Close(Type.Missing, Type.Missing, Type.Missing);
                                Marshal.FinalReleaseComObject(wkb);

                                wordapp.Quit();
                                Marshal.FinalReleaseComObject(wordapp);
                                Get_Status.staus_messages.Add("Converted " + Path.GetFileNameWithoutExtension(inputFile) + " successfully.");
                            }
                            catch(Exception exception)
                            {
                                filename = null;
                                Get_Status.staus_messages.Add(Path.GetFileNameWithoutExtension(inputFile) + " cannot be converted..");
                                Get_Status.set_error(exception, inputFile, folder_name);
                            }

                        }
                        break;

                }
            }

            
            return filename;
        }

        public static bool rm_Sheets(string file, Excel.Application app)
        {
            var isError = false;
            try
            {
                Get_Status.staus_messages.Add("Deleting sheet from - " + Path.GetFileNameWithoutExtension(file) + " ..");

                Excel.Workbook wkb = app.Workbooks.Open(file);

                app.DisplayAlerts = false;
                if (excel_settings.Default.trim_2_page)
                {
                 
                    wkb.Sheets[2].Delete();
                  
                }
                else if (excel_settings.Default.trim_3_page)
                {
                    wkb.Sheets[3].Delete();
                }           

                wkb.Save();

                GC.Collect();
                GC.WaitForPendingFinalizers();

                wkb.Close(Type.Missing, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(wkb);

                Get_Status.staus_messages.Add("Deleted sheet from - " + Path.GetFileNameWithoutExtension(file));
            }
            catch(Exception exception)
            {
                isError = true;
                Get_Status.staus_messages.Add("Error occured while deleting sheet from  - " + Path.GetFileNameWithoutExtension(file) + " ..");
                Get_Status.set_error(exception, file, "Analysis files");
            }


            return isError;

        }

        public static bool rm_Sheet_Additional_Borderaux(string file, Excel.Application app)
        {
            var isError = false;
            try
            {
                Get_Status.staus_messages.Add("Deleting 'Not shown' sheet from - " + Path.GetFileNameWithoutExtension(file) + " ..");

                Excel.Workbook wkb = app.Workbooks.Open(file);

                app.DisplayAlerts = false;

                foreach (Excel.Worksheet sheet in wkb.Sheets)
                {
                    if (!sheet.Name.Contains("Not Shown"))
                    {
                        Excel.Range range = (Excel.Range)sheet.Columns["A:AA", Type.Missing];

                        Excel.Range resultRange = range.Find
                        (
                        What: "No Data",
                        LookIn: Excel.XlFindLookIn.xlValues,
                        LookAt: Excel.XlLookAt.xlPart,
                        SearchOrder: Excel.XlSearchOrder.xlByRows,
                        SearchDirection: Excel.XlSearchDirection.xlNext
                        );

                        if (resultRange != null)
                        {
                            sheet.Delete();
                        }
                    }
                    else
                    {
                        sheet.Delete();
                    }
                }

                wkb.Save();

                GC.Collect();
                GC.WaitForPendingFinalizers();

                wkb.Close(Type.Missing, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(wkb);

                Get_Status.staus_messages.Add("Deleting sheet from - " + Path.GetFileNameWithoutExtension(file) + " ..");
            }
            catch(Exception exception)
            {
                isError = true;
                Get_Status.staus_messages.Add("Error occured while deleting sheet from - " + Path.GetFileNameWithoutExtension(file) + " ..");
                Get_Status.set_error(exception, file, "Add. Bordx files");

            }
            return isError;

        }

        public static List<string> connect_ToDatabase(string connectionString, string command)
        {
            List<string> items = new List<string>();
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    SqlCommand cmd = new SqlCommand(command, connection);
                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            items.Add(rdr.GetString(0).Trim());
                        }
                    }
                }
            }
            catch(Exception exception)
            {
                items.Add("...");

            }
            return items;
        }

        public static void SetDoubleBuffering(System.Windows.Forms.Control control, bool value)
        {
            System.Reflection.PropertyInfo controlProperty = typeof(System.Windows.Forms.Control)
                .GetProperty("DoubleBuffered", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            controlProperty.SetValue(control, value, null);
        }

  
    }
}
