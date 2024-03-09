using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDemo.Models
{
    internal class ExcelFile
    {
        public void DeleteFileIfExists(FileInfo file)
        {
            if (file.Exists)
                file.Delete();
        }

        public async Task SaveExcelFile(List<Person> people, FileInfo file)
        {
            DeleteFileIfExists(file);

            using var package = new ExcelPackage(file);

            var workSheet = package.Workbook.Worksheets.Add("Main Report");

            var dataRange = workSheet.Cells["A2"].LoadFromCollection(people,true);

            dataRange.AutoFitColumns();

            //format the header
            workSheet.Cells["A1"].Value = "Our Main Report";
            workSheet.Cells["A1:C1"].Merge = true;

            //formating cells
            workSheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            workSheet.Column(1).Width = 30;
            workSheet.Column(2).Width = 30;
            workSheet.Column(3).Width = 30;

            workSheet.Row(1).Style.Font.Size = 30;
            workSheet.Row(2).Style.Font.Size = 16;
            workSheet.Row(1).Style.Font.Color.SetColor(Color.Blue);

            workSheet.Row(2).Style.Font.Bold = true;



            //save excel
            await package.SaveAsync();
        }

        public async Task<List<Person>> LoadExcelFile(FileInfo file)
        {
            List<Person> output = new();

            using var package = new ExcelPackage(file);

            await package.LoadAsync(file);

            var workSheet = package.Workbook.Worksheets[0];

            int row = 3;
            int column = 1;

            while (!string.IsNullOrWhiteSpace(workSheet.Cells[row,column].Value?.ToString()))
            {
                var person = new Person();
                person.Id = int.Parse(workSheet.Cells[row,column].Value.ToString());
                person.FirstName = workSheet.Cells[row, column + 1].Value.ToString();
                person.LastName = workSheet.Cells[row, column + 2].Value.ToString();

                output.Add(person);
                row++;
            }

            return output;
        }
    }
}
