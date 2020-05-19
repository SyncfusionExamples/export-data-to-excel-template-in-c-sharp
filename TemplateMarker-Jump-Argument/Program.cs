using Syncfusion.XlsIO;
using System.Collections.Generic;
using System.IO;

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
            //Instantiate the spreadsheet creation engine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding header text
                worksheet["A1"].Text = "\"Jump\" Argument";
                
                worksheet["A3"].Text = "Name";
                worksheet["B3"].Text = "Id";
                worksheet["C3"].Text = "Age";

                worksheet["A3:C3"].CellStyle.Font.Bold = true;
               
                //Adding markers dynamically with the argument, 'Jump'
                worksheet["A4"].Text = "%Employee.Name;jump:R[2]C";
                worksheet["B4"].Text = "%Employee.Id;jump:R[2]C";
                worksheet["C4"].Text = "%Employee.Age;jump:R[2]C";
               
                //Create template marker processor
                ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

                //Add marker variable
                marker.AddVariable("Employee", GetEmployeeDetails());

                //Apply markers
                marker.ApplyMarkers();

                //Save and close the workbook
                Stream stream = File.Create("Output.xlsx");
                worksheet.UsedRange.AutofitColumns();
                workbook.SaveAs(stream);
            }
        }
    }
}
