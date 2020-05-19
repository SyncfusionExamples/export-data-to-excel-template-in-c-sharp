using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.XlsIO;
using System.IO;
using System.Reflection;
using System.Data;
using System.Drawing;

namespace TemplateMarker
{
    class Program
    {
        # region Private Members
        private static DataTable northwindDt;
        private static DataTable numbersDt;
        public static IList<Customer> _customers = new List<Customer>();

        # endregion

        static void Main(string[] args)
        {
            
            #region Control data formatting

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                //Open an existing spreadsheet which will be used as a template for generating the new spreadsheet.
                //After opening, the workbook object represents the complete in-memory object model of the template spreadsheet.
                IWorkbook workbook;

                //Open existing workbook with data entered
                Assembly assembly = typeof(Program).GetTypeInfo().Assembly;
                Stream cfFileStream = assembly.GetManifestResourceStream("TemplateMarker.Data.TemplateMarkerFormatting.xlsx");
                workbook = excelEngine.Excel.Workbooks.Open(cfFileStream);

                //The first worksheet object in the worksheets collection is accessed.
                IWorksheet worksheet = workbook.Worksheets[0];

                //Create Template Marker Processor
                ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

                //Get first 5 items in the collection and bind them to the marker
                marker.AddVariable("SalesList", GetCustomerAsObjects().Take<Customer>(5).ToArray());

                //Process the markers in the template
                marker.ApplyMarkers();

                worksheet.Activate();

                //Saving and closing the workbook
                workbook.SaveAs("TemplateMarkerFormattingOutput.xlsx");

                //Close the workbook
                workbook.Close();
            }
            #endregion
                      
        }

        /// <summary>
        /// Gets the Collection of objects from the XML data.
        /// </summary>
        /// <returns>Collection of Customer Objects</returns>
        private static IList<Customer> GetCustomerAsObjects()
        {
            DataSet customersDataSet = new DataSet();

            //Open existing workbook with data entered
            Assembly assembly = typeof(Program).GetTypeInfo().Assembly;
            Stream dataStream = assembly.GetManifestResourceStream("TemplateMarker.Data.customers.xml");

            dataStream.Position = 0;
            customersDataSet.ReadXml(dataStream, XmlReadMode.ReadSchema);
            northwindDt = customersDataSet.Tables[0];
            IList<Customer> tmpCustomers = new List<Customer>();
            Customer customer = new Customer();
            numbersDt = GetTable();
            DataRowCollection rows = northwindDt.Rows;
            foreach (DataRow row in rows)
            {
                customer = new Customer();
                customer.SalesPerson = row[0].ToString();
                customer.SalesJanJune = Convert.ToInt32(row[1]);
                customer.SalesJulyDec = Convert.ToInt32(row[2]);
                customer.Image = GetImage(Convert.ToString(row[4]));
                tmpCustomers.Add(customer);
            }
            return tmpCustomers;
        }

        private static byte[] GetImage(string path)
        {
            Assembly assembly = typeof(Program).GetTypeInfo().Assembly;
            Stream imageStream = assembly.GetManifestResourceStream("TemplateMarker.Images." + path);
            using (BinaryReader reader = new BinaryReader(imageStream))
            {
                return reader.ReadBytes((int)imageStream.Length);
            }
        }

        private static DataTable GetTable()
        {
            Random r = new Random();
            DataTable dt = new DataTable("NumbersTable");

            int nCols = 4;
            int nRows = 10;

            for (int i = 0; i < nCols; i++)
                dt.Columns.Add(new DataColumn("Column" + i.ToString()));

            for (int i = 0; i < nRows; ++i)
            {
                DataRow dr = dt.NewRow();
                for (int j = 0; j < nCols; j++)
                    dr[j] = r.Next(0, 10);
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
}