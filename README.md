# bubble
using System;
using System.ComponentModel;
using System.Diagnostics.Metrics;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;

class Program
{
    static void Main(string[] args)
    {
        int comparisonsv = 0;
        int swapsv = 0;
        int comparisonsu = 0;
        int swapsu = 0;
        Random rnd = new Random();
        int rn=rnd.Next(1,10);
        // Создаем новый файл Excel и добавляем в него лист
        FileInfo file = new FileInfo("rez.xlsx");
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (ExcelPackage excel = new ExcelPackage(file))
        {
            var ws = excel.Workbook.Worksheets.FirstOrDefault(x => x.Name == "Sheet1");
            if (ws == null)
            {
                ws = excel.Workbook.Worksheets.Add("Sheet1");
            }
            else
            {
                excel.Workbook.Worksheets["Sheet1"].Cells.Clear();
                excel.Save();
            }
        }
        // Создание файла
        CreateRandomFile();
        // Открываем файл для чтения
        using (StreamReader reader = new StreamReader("input.txt"))
        {
            // Считываем содержимое файла
            string fileContent = reader.ReadToEnd();

            // Проверяем, есть ли значения в файле
            if (!string.IsNullOrEmpty(fileContent))
            {
               
                // Чтение массива из файла
                int[] arrv = ReadArrayFromFile("input.txt");
                int[] arru = ReadArrayFromFile("input.txt");
                int[] arrv1 = ReadArrayFromFile("input.txt");
                int[] arru1 = ReadArrayFromFile("input.txt");


                Console.WriteLine("Исходный массив:");
                Console.WriteLine("-----------------------");
                Console.WriteLine("| Индекс | Значение |");
                Console.WriteLine("-----------------------");
                for (int i = 0; i < arrv.Length; i++)
                {
                    Console.WriteLine($"| {i,6} | {arrv[i],8} |");
                }
                Console.WriteLine("-----------------------");
                using (ExcelPackage excel = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = excel.Workbook.Worksheets["Sheet1"];
                    worksheet.Cells[1, 1].Value = "Исходный массив:";
                    for (int i = 0; i < arrv.Length; i++)
                    {
                        worksheet.Cells[i + 2, 1].Value = arrv[i];
                        worksheet.Columns[1].Width = 10;
                        worksheet.Rows[1].Height = 30;
                        worksheet.Cells[1, 1].Style.WrapText = true;
                    }
                    excel.Save();
                }

                Console.WriteLine("Отсортированный массив в прямом порядке:");
                BubbleSortv(arrv, ref comparisonsv, ref swapsv);
                Console.WriteLine("-----------------------");
                Console.WriteLine("| Индекс | Значение |");
                Console.WriteLine("-----------------------");
                for (int i = 0; i < arrv.Length; i++)
                {
                    Console.WriteLine($"| {i,6} | {arrv[i],8} |");
                }
                Console.WriteLine("Количество сравнений: " + comparisonsv);
                Console.WriteLine("Количество перестановок: " + swapsv);
                Console.WriteLine("-----------------------");
                using (ExcelPackage excel = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = excel.Workbook.Worksheets["Sheet1"];
                    worksheet.Cells[1, 2].Value = "Отсортированный массив в прямом порядке:";
                    for (int i = 0; i < arrv.Length; i++)
                    {
                        worksheet.Cells[i + 2, 2].Value = arrv[i];
                        worksheet.Columns[2].Width = 25;
                        worksheet.Cells[1, 2].Style.WrapText = true;
                    }
                    worksheet.Cells[1, 3].Value = "Количество сравнений:";
                    worksheet.Cells[2, 3].Value = comparisonsv;
                    worksheet.Cells[1, 4].Value = "Количество перестановок:";
                    worksheet.Cells[2, 4].Value = swapsv;
                    worksheet.Cells[1, 3].Style.WrapText = true;
                    worksheet.Cells[1, 4].Style.WrapText = true;
                    worksheet.Columns[3].Width = 11;
                    worksheet.Columns[4].Width = 14;
                    excel.Save();
                }

                Console.WriteLine("Отсортированный массив в обратном порядке:");
                BubbleSortu(arru, ref comparisonsu, ref swapsu);
                Console.WriteLine("-----------------------");
                Console.WriteLine("| Индекс | Значение |");
                Console.WriteLine("-----------------------");
                for (int i = 0; i < arru.Length; i++)
                {
                    Console.WriteLine($"| {i,6} | {arru[i],8} |");
                }
                Console.WriteLine("Количество сравнений: " + comparisonsu);
                Console.WriteLine("Количество перестановок: " + swapsu);
                Console.WriteLine("-----------------------");
                using (ExcelPackage excel = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = excel.Workbook.Worksheets["Sheet1"];
                    worksheet.Cells[1, 5].Value = "Отсортированный массив в обратном порядке: ";
                    for (int i = 0; i < arru.Length; i++)
                    {
                        worksheet.Cells[i + 2, 5].Value = arru[i];
                        worksheet.Columns[5].Width = 25;
                        worksheet.Cells[1, 5].Style.WrapText = true;
                    }
                    worksheet.Cells[1, 6].Value = "Количество сравнений: ";
                    worksheet.Cells[2, 6].Value = comparisonsu;
                    worksheet.Cells[1, 7].Value = "Количество перестановок: ";
                    worksheet.Cells[2, 7].Value = swapsu;
                    worksheet.Cells[1, 6].Style.WrapText = true;
                    worksheet.Cells[1, 7].Style.WrapText = true;
                    worksheet.Columns[6].Width = 11;
                    worksheet.Columns[7].Width = 14;
                    excel.Save();
                }

                Console.WriteLine("Отсортированный массив в случайном порядке:");
                if (rn >= 1 && rn <= 5)
                {
                    comparisonsv = 0;
                    swapsv = 0;
                    BubbleSortv(arrv1, ref comparisonsv, ref swapsv);
                    Console.WriteLine("-----------------------");
                    Console.WriteLine("| Индекс | Значение |");
                    Console.WriteLine("-----------------------");
                    for (int i = 0; i < arrv1.Length; i++)
                    {
                        Console.WriteLine($"| {i,6} | {arrv1[i],8} |");
                    }
                    Console.WriteLine("Количество сравнений: " + comparisonsv);
                    Console.WriteLine("Количество перестановок: " + swapsv);
                    Console.WriteLine("-----------------------");
                    using (ExcelPackage excel = new ExcelPackage(file))
                    {
                        ExcelWorksheet worksheet = excel.Workbook.Worksheets["Sheet1"];
                        worksheet.Cells[1, 8].Value = "Отсортированный массив в случайном порядке: ";
                        for (int i = 0; i < arrv1.Length; i++)
                        {
                            worksheet.Cells[i + 2, 8].Value = arrv1[i];
                            worksheet.Columns[8].Width = 25;
                            worksheet.Cells[1, 8].Style.WrapText = true;
                        }
                        worksheet.Cells[1, 9].Value = "Количество сравнений: ";
                        worksheet.Cells[2, 9].Value = comparisonsv;
                        worksheet.Cells[1, 10].Value = "Количество перестановок:";
                        worksheet.Cells[2, 10].Value = swapsv;
                        worksheet.Cells[1, 9].Style.WrapText = true;
                        worksheet.Cells[1, 10].Style.WrapText = true;
                        worksheet.Columns[9].Width = 11;
                        worksheet.Columns[10].Width = 14;
                        excel.Save();
                    }
                }
                else if (rn >= 6 && rn <= 10)
                {
                    comparisonsu = 0;
                    swapsu = 0;
                    BubbleSortu(arru1, ref comparisonsu, ref swapsu);
                    Console.WriteLine("-----------------------");
                    Console.WriteLine("| Индекс | Значение |");
                    Console.WriteLine("-----------------------");
                    for (int i = 0; i < arru1.Length; i++)
                    {
                        Console.WriteLine($"| {i,6} | {arru1[i],8} |");
                    }
                    Console.WriteLine("Количество сравнений: " + comparisonsu);
                    Console.WriteLine("Количество перестановок: " + swapsu);
                    Console.WriteLine("-----------------------");
                    using (ExcelPackage excel = new ExcelPackage(file))
                    {
                        ExcelWorksheet worksheet = excel.Workbook.Worksheets["Sheet1"];
                        worksheet.Cells[1, 8].Value = "Отсортированный массив в случайном порядке:";
                        for (int i = 0; i < arru1.Length; i++)
                        {
                            worksheet.Cells[i + 2, 8].Value = arru1[i];
                            worksheet.Columns[8].Width = 25;
                            worksheet.Cells[1, 8].Style.WrapText = true;
                        }
                        worksheet.Cells[1, 9].Value = "Количество сравнений: ";
                        worksheet.Cells[2, 9].Value = comparisonsu;
                        worksheet.Cells[1, 10].Value = "Количество перестановок: ";
                        worksheet.Cells[2, 10].Value = swapsu;
                        worksheet.Cells[1, 9].Style.WrapText = true;
                        worksheet.Cells[1, 10].Style.WrapText = true;
                        worksheet.Columns[9].Width = 11;
                        worksheet.Columns[10].Width = 14;
                        excel.Save();
                    }
                }
            }
            else
            {
                Console.WriteLine("Файл не содержит значений.");
                using (ExcelPackage excel = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = excel.Workbook.Worksheets["Sheet1"];
                    worksheet.Cells[1, 1].Value = "Файл не содержит значений.";
                    excel.Save();
                }
            }
        }
        
    }
    // Создание файла с рандомными значениями
    static void CreateRandomFile()
    {
        try
        {
            Console.WriteLine("Создание файла со значениями");

            Console.WriteLine("Введите размерность массива:");
            int size = int.Parse(Console.ReadLine());
      
            Console.WriteLine("Выберите способ заполнения файла:");
            Console.WriteLine("1. Автоматически");
            Console.WriteLine("2. Вручную");
            int choice = int.Parse(Console.ReadLine());

            if (choice == 1)
            {
                
                    Console.WriteLine("Введите диапазон ОТ:");
                int min = int.Parse(Console.ReadLine());
                Console.WriteLine("Введите диапазон ДО:");
                int max = int.Parse(Console.ReadLine());
                if (min > max)
                {
                    Console.WriteLine("Ошибка: диапазон ОТ должен быть меньше или равен диапазону ДО.");
                    return;
                }
                    Random rand = new Random();
                    using (StreamWriter sw = new StreamWriter("input.txt"))
                    {
                        for (int i = 0; i < size; i++)
                        {
                            sw.WriteLine(rand.Next(min, max + 1));
                        }
                    }
               
                
            }
            else if (choice == 2)
            {
                 using (StreamWriter sw = new StreamWriter("input.txt"))
                 {
                    try
                    {
                        
                        for (int i = 0; i < size; i++)
                        {
                            Console.WriteLine("Введите {0} элемент массива:", i + 1);
                            string userInput = Console.ReadLine();
                            sw.WriteLine(userInput);
                            if (!int.TryParse(userInput, out int result))
                            {
                                throw new Exception("Вы ввели неверное значение.");
                            }
                        }
                    }
                    catch (Exception)
                    {
                        Console.WriteLine($"Вы ввели неверное значение.");
                        throw;
                    }
                 }
            }            
        }
        catch (Exception)
        {
            using (StreamWriter sw = new StreamWriter("input.txt"))
            {
                sw.Write("");
            }
            Console.WriteLine("неверное значение");
        }
    }
    static int[] ReadArrayFromFile(string filename)
    {
        try
        {
            string[] lines = File.ReadAllLines(filename);
            int[] arr = new int[lines.Length];
            for (int i = 0; i < lines.Length; i++)
            {
                arr[i] = int.Parse(lines[i]);
            }
            return arr;
        }
        catch (Exception)
        {
            Console.WriteLine($"неправильно");
            throw;
        }
    }
    static void BubbleSortv(int[] arr, ref int comparisonsv, ref int swapsv)
    {
        // проход по всем элементам массива
        for (int i = 0; i < arr.Length; i++)
        {
            // проход по всем элементам массива, кроме последнего
            for (int j = 0; j < arr.Length - 1; j++)
            {
                // увеличение счетчика сравнений
                comparisonsv++;
                // если текущий элемент больше следующего
                if (arr[j] > arr[j + 1])
                {
                    // меняем их местами
                    int temp = arr[j];
                    arr[j] = arr[j + 1];
                    arr[j + 1] = temp;
                    // увеличение счетчика перестановок
                    swapsv++;
                }
            }
        }
    }
    static void BubbleSortu(int[] arr, ref int comparisonsu, ref int swapsu)
    {
        for (int i = 0; i < arr.Length; i++)
        {
            for (int j = 0; j < arr.Length - 1; j++)
            {
                comparisonsu++;
                if (arr[j] < arr[j + 1])
                {
                    int temp = arr[j];
                    arr[j] = arr[j + 1];
                    arr[j + 1] = temp;
                    swapsu++;
                }
            }
        }
    }
 }
