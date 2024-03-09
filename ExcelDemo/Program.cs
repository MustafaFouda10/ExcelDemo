
using ExcelDemo.Models;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var file = new FileInfo(@"D:\ITI PROJECTS\ExcelDemo\ExcelDemo\Files\ExcelDemo.xlsx");

var person = new Person();

var people = person.GetSetupData();

var excelFile = new ExcelFile();

excelFile.SaveExcelFile(people, file);

var peopleFromExcelFile = await excelFile.LoadExcelFile(file);

foreach (var p in peopleFromExcelFile)
{
    Console.WriteLine($"{p.Id}. {p.FirstName} {p.LastName}");
}