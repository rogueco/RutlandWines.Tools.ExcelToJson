using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using RutlandWinesCsvToJson.Model;
using Newtonsoft.Json;

namespace RutlandWinesCsvToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Starting Application");
            GetNewExelFile();
        }


        public static void GetNewExelFile()
        {
            using (var package = new ExcelPackage(new FileInfo("UploadSheet.xlsx")))
            {
                var workSheet = package.Workbook.Worksheets["Details"];
                var start = workSheet.Dimension.Start;
                var end = workSheet.Dimension.End;
                var test = workSheet.Cells;

                List<Product> productList = new List<Product>();

                for (int row = start.Row; row <= end.Row; row++)
                {

                    Product product = new Product();

                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        string cellValue = workSheet.Cells[row, col].Text;

                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            switch (col)
                            {
                                case 1:
                                    product.Vintage = cellValue == "NV" ? 0 : Int32.Parse(cellValue);
                                    break;
                                case 2:
                                    product.Title = cellValue;
                                    break;
                                case 3:
                                    product.Region = cellValue;
                                    break;
                                case 4:
                                    product.Country = cellValue;
                                    break;
                                case 5:
                                    product.FoodSuggestion = cellValue;
                                    break;
                                case 6:
                                    product.SellingStory = cellValue;
                                    break;
                                case 7:
                                    product.ABV = Double.Parse(cellValue);
                                    break;
                                case 8:
                                    product.Size = Int32.Parse(cellValue);
                                    break;
                                case 9:
                                    product.RestaurantPrice = Double.Parse(cellValue);
                                    break;
                                case 10:
                                    product.CaseSalePrice = Double.Parse(cellValue);
                                    break;
                                case 11:
                                    product.IndividualBottleSalePrice = Double.Parse(cellValue);
                                    break;
                                case 12:
                                    product.MilkUsed = cellValue == "Yes" ? true : false;
                                    break;
                                case 13:
                                    product.EggUsed = cellValue == "Yes" ? true : false;
                                    break;
                                case 14:
                                    product.Sulphites = cellValue == "Yes" ? true : false;
                                    break;
                                case 15:
                                    product.Vegatarain = cellValue == "Yes" ? true : false;
                                    break;
                                case 16:
                                    product.Vegan = cellValue == "Yes" ? true : false;
                                    break;
                                case 17:
                                    product.Organic = cellValue == "Yes" ? true : false;
                                    break;
                                case 18:
                                    product.Biodynamic = cellValue == "Yes" ? true : false;
                                    break;
                                case 19:
                                    product.WineType = cellValue;
                                    break;
                            }
                        }
                    }

                    productList.Add(product);
                }

                string json = JsonConvert.SerializeObject(productList.ToArray());

                Console.WriteLine("What would you like your file to be named?");
                string fileName = Console.ReadLine();

                System.IO.File.WriteAllTextAsync(@$"/Users/tomfletcher/Desktop/{fileName}.json", json);

                Console.WriteLine($"Program has finished, there has been {productList.Count} items added into the {fileName}.json. Press any button to exit program.");
                Console.ReadLine();
            }
        }
    }

}
