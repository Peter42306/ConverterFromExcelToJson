using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using static ConverterFromExcelToJson.Program;

namespace ConverterFromExcelToJson
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filePath = "D:\\MY РАБОТА\\BVSS\\2021 12 16, 2021 033, Johnny Cash, b+c+d+h\\vcfData.xlsx"; // путь к Excel файлу
            string jsonOutputPath = "D:\\MY РАБОТА\\BVSS\\2021 12 16, 2021 033, Johnny Cash, b+c+d+h\\vcfData.json";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Получаем первый лист
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // Пример поиска данных
                
                //FindData(worksheet, 14, 18); // Начинаем с D6 (4, 6)

                //GetTemperatures(worksheet);
                //GetDensities(worksheet);                
                GetVCFValue(worksheet, 21, .8228);               

                List<VcfData> vcfDatas = ExtractVCFDataFromExcel(filePath);

                SaveDataToJson(vcfDatas, jsonOutputPath);

                //PrintAllData(vcfDatas);
            }


            //string filePath = "D:\\MY РАБОТА\\BVSS\\2021 12 16, 2021 033, Johnny Cash, b+c+d+h\\vcfData.xlsx"; // путь к Excel файлу
            //string jsonOutputPath = "D:\\MY РАБОТА\\BVSS\\2021 12 16, 2021 033, Johnny Cash, b+c+d+h\\vcfData.json"; // путь для сохранения JSON файла

            //List<VcfData> vcfDataList = ExtractVCFDataFromExcel(filePath);
            //SaveDataToJson(vcfDataList, jsonOutputPath);

            //Console.WriteLine("JSON файл создан: " + jsonOutputPath);
        }
        
        //===================================================================================

        static void FindData(ExcelWorksheet worksheet, double targetTemperature, double targetDensity)
        {
            // Получаем диапазон температур и плотностей
            for (int col = 4; col <= worksheet.Dimension.Columns; col++) // Начинаем с колонки D (4)
            {
                double temperature = Convert.ToDouble(worksheet.Cells[6, col].Value); // Температура в строке 6
                double density = Convert.ToDouble(worksheet.Cells[7, col].Value); // Плотность в строке 7

                // Сравниваем с заданными целевыми значениями
                if (temperature == targetTemperature)
                {
                    Console.WriteLine($"Температура {targetTemperature} найдена в колонке {col}");
                }

                if (density == targetDensity)
                {
                    Console.WriteLine($"Плотность {targetDensity} найдена в колонке {col}");
                }

                // Если обе целевые величины найдены, ищем соответствующий VCF
                if (temperature == targetTemperature && density == targetDensity)
                {
                    // Значение VCF будет на пересечении температуры и плотности
                    double vcfValue = Convert.ToDouble(worksheet.Cells[8, col].Value); // VCF в строке 8
                    Console.WriteLine($"VCF для Температуры: {temperature} и Плотности: {density} = {vcfValue}");
                }
            }
        }

        public class VcfData
        {
            public double Temperature { get; set; }
            public double Density { get; set; }
            public double VCF { get; set; }
        }

        static List<VcfData>ExtractVCFDataFromExcel(string filePath)
        {
            var vcfDataList=new List<VcfData>();

            ExcelPackage.LicenseContext=LicenseContext.Commercial;            

            using(ExcelPackage package=new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                double[] temperatures = GetTemperatures(worksheet);
                double[] densities = GetDensities(worksheet);

                for (int i = 0; i < temperatures.Length; i++)
                {
                    for (int j = 0; j < densities.Length; j++)
                    {
                        double vcfValue = Convert.ToDouble(worksheet.Cells[j + 8, i + 4].Value); // D8 - это 8 строка, 4 столбец
                        vcfDataList.Add(new VcfData
                        {
                            Temperature = temperatures[i],
                            Density = densities[j],
                            VCF = vcfValue
                        });
                    }
                }
            }

            return vcfDataList;
        }

        static void SaveDataToJson(List<VcfData> vcfDataList, string jsonOutputPath)
        {
            string json = JsonConvert.SerializeObject(vcfDataList, Formatting.Indented);
            File.WriteAllText(jsonOutputPath, json);
            Console.WriteLine("JSON file was created: "+ jsonOutputPath);
        }

        static void PrintAllData(List<VcfData>vcfDataList)
        {
            foreach (var item in vcfDataList)
            {
                Console.WriteLine($"Temperature: {item.Temperature}, Density: {item.Density}, VCF: {item.VCF}");
            }
        }

        

        static double[] GetTemperatures(ExcelWorksheet worksheet)
        {
            int columnCount = worksheet.Dimension.Columns; // quantity of columns, which are not empty
            Console.WriteLine($"columnCount = {columnCount}");

            double[] temperatures = new double[columnCount];

            int firstCellsInRow = 4; // the first cell with data in our row
            int lastCellInRow = columnCount;
            int rowToTakeData = 6; // the row where we need data


            for (int col = firstCellsInRow; col <= lastCellInRow; col++)
            {
                temperatures[col - firstCellsInRow] = Convert.ToDouble(worksheet.Cells[rowToTakeData, col].Value);
                string cellAdress = worksheet.Cells[rowToTakeData, col].Address;
                //Console.WriteLine($"temperatures: {temperatures[col - firstCellsInRow]} in {cellAdress}");
            }
            return temperatures;
        }        

        static double[] GetDensities(ExcelWorksheet worksheet)
        {
            int rowCount = worksheet.Dimension.Rows;
            Console.WriteLine($"rowCount = {rowCount}");

            double[] densities = new double[rowCount];

            int firstCellInColumn = 8;
            int lastCellInColumn = rowCount;
            int columnToTakeData = 2;
            int totalRows = lastCellInColumn - firstCellInColumn + 1;

            for (int row = firstCellInColumn; row <= lastCellInColumn+1; row++)
            {
                densities[row - firstCellInColumn] = Convert.ToDouble(worksheet.Cells[row, columnToTakeData].Value);
                string cellAdress = worksheet.Cells[row, columnToTakeData].Address;
                //Console.WriteLine($"Density: {densities[row - firstCellInColumn]} in {cellAdress}");
            }
            return densities;
        }

        static double GetVCFValueByCellIndex(ExcelWorksheet worksheet,int temperatureColumn, int densityRow)
        {
            Console.WriteLine($"at {temperatureColumn} and {densityRow} VCF is {Convert.ToDouble(worksheet.Cells[densityRow, temperatureColumn].Value)}");
            return Convert.ToDouble(worksheet.Cells[densityRow, temperatureColumn].Value);
        }

        static double GetVCFValue(ExcelWorksheet worksheet, double targetTemperature, double targetDensity)
        {
            double[]temperatures=GetTemperatures(worksheet);
            double[]densities=GetDensities(worksheet);

            int firstCellsInRow = 4; // the first cell with data in our row
            int temperatureColumnIndex = -1;
            for (int i = 0; i < temperatures.Length; i++)
            {
                //Console.WriteLine($"Array with temperatures: {temperatures[i]} / {i+firstCellsInRow}");
                if (temperatures[i] == targetTemperature)
                {
                    temperatureColumnIndex = i + firstCellsInRow;
                    Console.WriteLine($"target column index: {temperatureColumnIndex}");
                    break;
                }
            }

            int firstCellInColumn = 8;
            int densityRowIndex = -1;            
            for (int i = 0; i < densities.Length; i++)
            {                
                if (Math.Abs(densities[i]-targetDensity)<0.0001)
                {
                    densityRowIndex = i + firstCellInColumn;
                    Console.WriteLine($"target row index: {densityRowIndex}");
                    break;
                }
            }

            // Проверяем, были ли найдены индексы
            if (temperatureColumnIndex == -1 || densityRowIndex == -1)
            {
                throw new Exception("Не удалось найти VCF для заданных температуры и плотности.");
            }

            double vcfValue = GetVCFValueByCellIndex(worksheet, temperatureColumnIndex, densityRowIndex);
            Console.WriteLine($"at temperature {targetTemperature} and at {targetDensity} VCF is {vcfValue}");

            return vcfValue;
        }        
    }
}
