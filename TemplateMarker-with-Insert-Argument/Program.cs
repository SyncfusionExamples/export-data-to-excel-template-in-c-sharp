using Syncfusion.XlsIO;
using System.Collections.Generic;
using System.IO;
using System.Drawing;

namespace TemplateMarker
{
    class Employee
    {
        private string m_name;
        private int m_id;
        private int m_age;

        public string Name
        {
            get
            {
                return m_name;
            }

            set
            {
                m_name = value;
            }
        }
        public int Id
        {
            get
            {
                return m_id;
            }

            set
            {
                m_id = value;
            }
        }
        public int Age
        {
            get
            {
                return m_age;
            }

            set
            {
                m_age = value;
            }
        }
    }

    class Program
    {
        public static List<Employee> GetEmployeeDetails()
        {
            List<Employee> employeeList = new List<Employee>();
            Employee emp = new Employee();
            emp.Name = "Andy Bernard";
            emp.Id = 1011;
            emp.Age = 35;

            employeeList.Add(emp);

            emp = new Employee();
            emp.Name = "Jim Halpert";
            emp.Id = 1012;
            emp.Age = 26;

            employeeList.Add(emp);

            emp = new Employee();
            emp.Name = "Karen Fillippelli";
            emp.Id = 1013;
            emp.Age = 28;

            employeeList.Add(emp);

            return employeeList;
        }

        public static void Main(string[] args)
        {
            InsertRows();
            InsertColumns();
        }

        static void InsertRows()
        {
            //Instantiate the spreadsheet creation engine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding header text
                worksheet["A1"].Text = "\"Insert\" Argument";
                worksheet["A3"].CellStyle.Font.RGBColor = Color.FromArgb(255, 0, 0);

                worksheet["A3"].Text = "\"Row\" Insertion with copy styles and copy merges";

                worksheet["A4:B4"].Merge();
                worksheet["A5:B5"].Merge();

                worksheet["A4"].Text = "Name";
                worksheet["C4"].Text = "Id";
                worksheet["D4"].Text = "Age";

                worksheet["A4:F4"].CellStyle.Font.Bold = true;

                worksheet["A5"].CellStyle.Font.Italic = true;

                //Adding markers dynamically with the arguments, 'insert','copystyles' and 'copymerges
                worksheet["A5"].Text = "%Employee.Name;insert:copystyles,copymerges";
                worksheet["C5"].Text = "%Employee.Id";
                worksheet["D5"].Text = "%Employee.Age";

                // This data will be moved to new row
                worksheet["A7"].Text = "Text in new row";

                worksheet["A9"].CellStyle.Font.RGBColor = Color.FromArgb(255, 0, 0);
                worksheet["A4:D4"].CellStyle.Color = Color.FromArgb(77, 176, 215);

                workbook.SaveAs("InsertRow-Template.xlsx");

                //Create template marker processor
                ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

                //Add marker variable
                marker.AddVariable("Employee", GetEmployeeDetails());

                //Apply markers
                marker.ApplyMarkers();

                //Save and close the workbook
                Stream stream = File.Create("InsertRows.xlsx");
                worksheet.UsedRange.AutofitColumns();
                workbook.SaveAs(stream);
            }
        }

        static void InsertColumns()
        {
            //Instantiate the spreadsheet creation engine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding header text
                worksheet["A1"].Text = "\"Insert\" Argument";

                worksheet["A3"].CellStyle.Font.RGBColor = Color.FromArgb(255, 0, 0);
                worksheet["A4:A6"].CellStyle.Color = Color.FromArgb(77, 176, 215);

                worksheet["A3"].Text = "\"Column\" Insertion with copy styles";
                worksheet["A4"].Text = "Name";
                worksheet["A5"].Text = "Id";
                worksheet["A6"].Text = "Age";

                worksheet["A4:A6"].CellStyle.Font.Bold = true;

                worksheet["B4"].CellStyle.Color = Color.FromArgb(189, 215, 238);

                //Adding markers dynamically with the arguments, 'insert' and 'copystyles' and. 'horizontal'
                worksheet["B4"].Text = "%Employee.Name;insert:copystyles;horizontal";
                worksheet["B5"].Text = "%Employee.Id;horizontal";
                worksheet["B6"].Text = "%Employee.Age;horizontal";

                // This data will be moved to new column
                worksheet["C6"].Text = "Text in new column";

                workbook.SaveAs("InsertColumn-Template.xlsx");

                //Create template marker processor
                ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

                //Add marker variable
                marker.AddVariable("Employee", GetEmployeeDetails());

                //Apply markers
                marker.ApplyMarkers();

                //Save and close the workbook
                Stream stream = File.Create("InsertColumns.xlsx");
                worksheet.UsedRange.AutofitColumns();
                workbook.SaveAs(stream);
            }
        }
    }
}
