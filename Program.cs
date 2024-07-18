using OfficeOpenXml;
using System.Reflection;
using System.Xml.Linq;

namespace ExcelReaderExample
{
    class Program
    {
        static void Main(string[] args)
        {
            //set the license
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Specify the file path
            string filePath = @"C:\REPORTES_ML\mlibre.xlsx";
            string newFilePath = @"C:\REPORTES_ML\Reporte Generado.xlsx";
                        

            // use the method to read the excel file
            ExcelReaderAndCleaner excelReaderAndCleaner = new ExcelReaderAndCleaner();

            // Define a data structure to store the data
            List<List<object>> data = new List<List<object>>();

            data = excelReaderAndCleaner.ReadExcelFile(filePath);

            // Print the extracted data (for demonstration)
            /*Console.WriteLine("Extracted data:");
            foreach (var row in data)
            {
                Console.WriteLine(string.Join(",", row));
            }*/

            // Create a new Excel file and write the data
            using (ExcelPackage newPackage = new ExcelPackage())
            {
                ExcelWorksheet newWorksheet = newPackage.Workbook.Worksheets.Add("Reporte1");

                // Write data to the new file
                // Loop through each row of the data
                for (int i = 0; i < data.Count; i++)
                {
                    // Loop through each cell in the row
                    for (int j = 0; j < data[i].Count; j++)
                    {
                        // Write the cell value to the worksheet
                        newWorksheet.Cells[i + 1, j + 1].Value = data[i][j];

                        if (j == data[i].Count - 1)
                        {
                            if (i == 0)
                                newWorksheet.Cells[i + 1, j + 3].Value = "Misma venta";
                            if (i > 0)
                                newWorksheet.Cells[i + 1, j + 3].Formula = "COUNTIF(K$2:K$"+ data.Count +",K" + (i + 1).ToString() + ")";
                        }

                    }
                    //especify column width for all columns
                    //newWorksheet.Column(i+1).Width = 25;
                }

                newWorksheet.Cells.AutoFitColumns();
                newPackage.SaveAs(newFilePath);
            }
            Console.WriteLine();
            Console.WriteLine("Data exported to new Excel file: {0}", newFilePath);
        }

        public class ExcelReaderAndCleaner
        {
            string masterFilePath = @"C:\REPORTES_ML\master.xlsx";
            
            //read the excel file
            public List<List<object>> ReadExcelFile(string filePath)
            {
                //set the license
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Read the Excel file using EPPlus library
                ExcelPackage package = new ExcelPackage(new FileInfo(filePath));

                // Get the first worksheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // Define a data structure to store the data
                List<List<object>> data = new List<List<object>>();

                // to process package rows
                bool isPackageSale = false;
                bool isSaleCancelled = false;
                int numberOfItemsInPackage = 0;
                int numberOfRowsInPackage = 0;
                int rowOfPackage = 0;
                float costo = 0;
                string ramo = "";

                // Loop through each row of the worksheet
                //for (int row = worksheet.Dimension.Start.Row; row <= worksheet.Dimension.End.Row; row++)
                for (int row = 6; row <= worksheet.Dimension.End.Row; row++)
                {
                    List<object> rowData = new List<object>();

                    // Loop through each cell in the row
                    for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                    {
                        if (col == 1 || col == 2 || col == 3 || col == 6 || col == 7 || col == 8 || col == 9 || col == 10 || col == 12 || col == 15 || col == 16 || col == 18 || col == 29)
                        // Get the cell value and add it to the list
                        {
                            if (col == 3 && row > 6)
                            {
                                string cellValue = worksheet.Cells[row, col].Value.ToString();
                                if (cellValue.Contains("Venta cancelada"))
                                {
                                    isSaleCancelled = true;
                                    numberOfItemsInPackage = 0;
                                    isPackageSale = false;
                                    break;
                                }
                                else if (cellValue.Contains("Paquete de"))
                                {
                                    //if the cell value is "Paquete de 3 productos" extract the number from the string as an int. The number can be from 1 to 99. Display such number in the console

                                    string[] words = cellValue.Split(' ');
                                    numberOfItemsInPackage = Convert.ToInt32(words[2]);
                                    Console.WriteLine("The number of products in the package is: " + numberOfItemsInPackage);
                                    isPackageSale = true;
                                    rowOfPackage = row;
                                    //row++;
                                    break;
                                }
                            }
                            if (col == 15 && row > 6)
                            {
                                // transform "MLM1490970065" into "#1490970065"
                                string cellValue = worksheet.Cells[row, col].Value.ToString();
                                string convertedID = "#" + cellValue.Substring(3, cellValue.Length - 3);
                                Tuple<float, string> costoEncontrado = FindStringInExcel(convertedID, masterFilePath);
                                costo = costoEncontrado.Item1;
                                ramo = costoEncontrado.Item2;
                                rowData.Add(convertedID);
                            }
                            else
                            {
                                rowData.Add(worksheet.Cells[row, col].Value);
                            }

                            if ((col == 7 || col == 8 || col == 9 || col == 10 || col == 12 || col == 29) && row > 6 && numberOfItemsInPackage > 0)                            
                            {
                                var cellValue = worksheet.Cells[rowOfPackage, col].Value;

                                switch (col)
                                {
                                    case 7:
                                        if (cellValue != "")
                                            rowData[4] = Convert.ToSingle(worksheet.Cells[row, 6].Value) * Convert.ToSingle(worksheet.Cells[row, 18].Value);
                                        break;
                                    case 8:
                                        if (cellValue != "")
                                            rowData[5] = (cellValue);
                                        break;
                                    case 9:
                                        if (cellValue != "")
                                            rowData[6] = Convert.ToSingle(worksheet.Cells[row, 18].Value) / Convert.ToSingle(worksheet.Cells[rowOfPackage, 7].Value) * Convert.ToSingle(worksheet.Cells[rowOfPackage, 9].Value);
                                        break;
                                    case 10:
                                        if (cellValue != "")
                                            rowData[7] = Convert.ToSingle(worksheet.Cells[row, 18].Value) / Convert.ToSingle(worksheet.Cells[rowOfPackage, 7].Value) * Convert.ToSingle(worksheet.Cells[rowOfPackage, 10].Value);
                                        break;
                                    case 12:
                                        if (cellValue != "")
                                            rowData[8] = Convert.ToSingle(worksheet.Cells[row, 18].Value) / Convert.ToSingle(worksheet.Cells[rowOfPackage, 7].Value) * Convert.ToSingle(worksheet.Cells[rowOfPackage, 12].Value);
                                        break;
                                    case 29:
                                        if (cellValue != "")
                                            rowData[12] = (cellValue);
                                        break;
                                }
                            }
                        }
                    }


                    if (!isPackageSale && !isSaleCancelled)
                    {
                        if (numberOfItemsInPackage > 0)
                        {
                            numberOfItemsInPackage--;
                        }
                        // Calculate the product of columns 3 and 4 (adjust indices if needed)
                        if (row > 6)
                        {
                            rowData.Add(ramo);
                            rowData.Add(costo);
                            rowData.Add(Convert.ToSingle(Convert.ToSingle(rowData[8])/ Convert.ToSingle(rowData[3]) - costo));
                            rowData.Add(Convert.ToSingle(rowData[3]) * Convert.ToSingle(rowData[15]));
                        }
                        //add the name for this new column
                        if (row == 6)
                        {
                            rowData.Add("Ramo");
                            rowData.Add("COSTO X ART");
                            rowData.Add("UTILIDAD X ART");
                            rowData.Add("UTILIDAD VENTA");
                        }

                        // Add the row data to the main list
                        data.Add(rowData);
                    }
                    else
                    {
                        isSaleCancelled = false;
                        isPackageSale = false;
                        continue;
                    }
                }

                return data;
            }
        }

        //class to find a string in an excel file which has several sheets and display the location by sheet, row and column
        static Tuple<float, string> FindStringInExcel(string searchValue, string filePath)
        {
            ExcelPackage package = new ExcelPackage(new FileInfo(filePath));
            bool found = false;

            // Loop through all worksheets
            foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
            {
                if(worksheet.Name == "Hoja1" || found)
                {
                    continue;
                }
                Console.WriteLine($"Searching {searchValue} in worksheet: {worksheet.Name}");
                
                // Loop through all cells in the current worksheet
                for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    for (int col = 1; col <= 1; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value;
                        if (cellValue == null || cellValue == "" || cellValue.ToString().Length < 7)
                        {
                            continue;
                        }

                        if (cellValue.ToString().Contains(searchValue))
                        {
                            Console.WriteLine($"Found match in worksheet '{worksheet.Name}', cell {row}, {col}: {cellValue}");
                            float costo = Convert.ToSingle(worksheet.Cells[row, 8].Value);

                            return Tuple.Create(costo, worksheet.Name);
                        }
                    }
                }
            }
            return Tuple.Create(0f, "");
        }
    }
}
