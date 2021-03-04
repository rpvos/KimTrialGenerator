using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace KimsPosibilityCalculator
{
    class Program
    {
        public struct Trial
        {
            public int ID { get; set; }
            public string Initial { get; set; }
            public string Second { get; set; }
            public string Awnser { get; set; }
        }

        static void Main(string[] args)
        {
            // Creates and initializes a new ArrayList.
            ArrayList trials = new ArrayList();
            trials = intitialiseTrials(trials);
            Random random = new Random();

            List<Trial> sortedList = new List<Trial>();

            Trial t1 = (Trial)trials[random.Next(0, trials.Count)];
            trials.Remove(t1);
            t1.ID = 1;
            sortedList.Add(t1);
            Console.WriteLine($"{sortedList.Count}\tintitial = {t1.Initial}\t second = {t1.Second}");
            
            int attempts = 0;
            while (trials.Count != 0)
            {
                attempts++;

                Trial t2 = (Trial)trials[random.Next(0, trials.Count)];
                if (t2.Initial == t1.Second || attempts > 50)
                {
                    trials.Remove(t2);
                    t1 = t2;
                    t2.ID = sortedList.Count +1;
                    sortedList.Add(t2);
                    Console.WriteLine($"{t2.ID}\tintitial = {t2.Initial}\t second = {t2.Second}");
                    attempts = 0;
                }
            }

            exportToExcel(sortedList);
        }

        private static void exportToExcel(List<Trial> sortedList)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");
                excel.Workbook.Worksheets.Add("Worksheet2");
                excel.Workbook.Worksheets.Add("Worksheet3");

                var headerRow = new List<string[]>()
                    {
                        new string[] { "ID", "Initial", "Second", "Outcome" }
                    };

                // Target a worksheet
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];


                worksheet.Cells["A1"].LoadFromCollection(sortedList,true, TableStyles.Medium9);
                
                FileInfo excelFile = new FileInfo(@"Trial.xlsx");
                excel.SaveAs(excelFile);
            }
        }

        private static ArrayList intitialiseTrials(ArrayList trials)
        {
            ArrayList posibilities = new ArrayList() {
                "dk160",
                "dk200",
                "dk240",
                "dg200",
                "vk160",
                "vk200",
                "vk240",
                "vg200"};

            for (int i = 0; i < posibilities.Count; i++)
            {
                for (int j = 0; j < posibilities.Count; j++)
                {
                    if (i != j)
                        trials.Add(new Trial { Initial = (string)posibilities[i], Second = (string)posibilities[j] });
                }
            }


            return trials;
        }
    }
}
