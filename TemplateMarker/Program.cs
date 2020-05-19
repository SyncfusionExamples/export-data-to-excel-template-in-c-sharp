﻿using Syncfusion.XlsIO;
using System.Data;
using System.IO;
using System.Reflection;

namespace TemplateMarker
{
    class Program
    {
        static void Main(string[] args)
        {
            //Code to read XML data to create a DataTable
            Assembly assembly = typeof(Program).GetTypeInfo().Assembly;
            Stream dataStream = assembly.GetManifestResourceStream("TemplateMarker.Data.customers.xml");
            DataSet customersDataSet = new DataSet();
            customersDataSet.ReadXml(dataStream, XmlReadMode.ReadSchema);
            DataTable northwindDt = customersDataSet.Tables[0];

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                //Open an existing spreadsheet which will be used as a template for generating the new spreadsheet.
                //After opening, the workbook object represents the complete in-memory object model of the template spreadsheet.
                IWorkbook workbook;

                //Open existing Excel template
                Stream cfFileStream = assembly.GetManifestResourceStream("TemplateMarker.Data.TemplateMarker.xlsx");
                workbook = excelEngine.Excel.Workbooks.Open(cfFileStream);

                //The first worksheet in the workbook is accessed.
                IWorksheet worksheet = workbook.Worksheets[0];

                //Create Template Marker processor.
                //Apply the marker to export data from datatable to worksheet.
                ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();
                marker.AddVariable("SalesList", northwindDt);
                marker.ApplyMarkers();

                //Saving and closing the workbook
                workbook.SaveAs("TemplateMarkerOutput.xlsx");

                //Close the workbook
                workbook.Close();
            }
        }
    }
}
