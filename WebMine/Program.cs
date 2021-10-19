
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Threading;
using Xceed.Document.NET;
using Xceed.Words.NET;


namespace WebMine
{
    class Program
    {
        static void Main(string[] args)
        {
            string pathToFile = AppDomain.CurrentDomain.BaseDirectory + '\\';
            IWebDriver driver = new ChromeDriver(pathToFile);
            /*driver.Navigate().GoToUrl("http://root:root@192.168.9.103/#/minerstatus");

            Thread.Sleep(10000);

            DataTable dt1 = new DataTable();
            dt1.Columns.Add("Elapsed");
            dt1.Columns.Add("GH / S(RT)");
            dt1.Columns.Add("GH / S(avg)");
            dt1.Columns.Add("FoundBlocks");
            dt1.Columns.Add("LocalWork");
            dt1.Columns.Add("Utility");
            dt1.Columns.Add("WU");
            dt1.Columns.Add("BestShare");

            DataTable dt2 = new DataTable();
            dt2.Columns.Add("Pool");
            dt2.Columns.Add("URL");
            dt2.Columns.Add("User");
            dt2.Columns.Add("Status");
            dt2.Columns.Add("Type");
            dt2.Columns.Add("Diff");
            dt2.Columns.Add("GetWorks");
            dt2.Columns.Add("Priority");
            dt2.Columns.Add("Accepted");
            dt2.Columns.Add("Diff1#");
            dt2.Columns.Add("DiffA#");
            dt2.Columns.Add("DiffR#");
            dt2.Columns.Add("DiffS#");
            dt2.Columns.Add("Rejected");
            dt2.Columns.Add("Discarded");
            dt2.Columns.Add("Stale");
            dt2.Columns.Add("LSDiff");
            dt2.Columns.Add("LSTime");

            DataTable dt3 = new DataTable();
            dt3.Columns.Add("Chain#");
            dt3.Columns.Add("ASIC#");
            dt3.Columns.Add("Frequency(avg)");
            dt3.Columns.Add("Voltage");
            dt3.Columns.Add("Consumption(W)");
            dt3.Columns.Add("GH/S(ideal)");
            dt3.Columns.Add("GH/S(RT)");
            dt3.Columns.Add("Status");
            dt3.Columns.Add("Errors(HW)");
            dt3.Columns.Add("Temp(PCB)");
            dt3.Columns.Add("Temp(Chip)");
            dt3.Columns.Add("ASIC status");

            DataTable dt4 = new DataTable();
            dt4.Columns.Add("Fan#");
            dt4.Columns.Add("Fan1");
            dt4.Columns.Add("Fan2");
            dt4.Columns.Add("Fan3");
            dt4.Columns.Add("Fan4");
            dt4.Columns.Add("Fan5");
            dt4.Columns.Add("Fan6");
            dt4.Columns.Add("Fan7");
            dt4.Columns.Add("Fan8");



            List<IWebElement> cell;

            cell = driver.FindElements(By.XPath("//fieldset[1]/fieldset/table/thead/tr/tr")).ToList();
            for (int i = 0; i < cell.Count; i += 8)
                dt1.Rows.Add(cell[i].Text, cell[i + 1].Text, cell[i + 2].Text, cell[i + 3].Text, cell[i + 4].Text, cell[i + 5].Text, cell[i + 6].Text, cell[i + 7].Text);

            cell = driver.FindElements(By.XPath("//fieldset[2]/fieldset/table/thead/tr/tr")).ToList();
            for (int i = 0; i < cell.Count; i += 18)
                dt2.Rows.Add(cell[i].Text, cell[i + 1].Text, cell[i + 2].Text, cell[i + 3].Text, cell[i + 4].Text, cell[i + 5].Text, cell[i + 6].Text, cell[i + 7].Text, cell[i + 8].Text, cell[i + 9].Text, cell[i + 10].Text, cell[i + 11].Text, cell[i + 12].Text, cell[i + 13].Text, cell[i + 14].Text, cell[i + 15].Text, cell[i + 16].Text, cell[i + 17].Text);

            cell = driver.FindElements(By.XPath("//fieldset[3]/fieldset[1]/table/thead/tr/tr")).ToList();
            for (int i = 0; i < cell.Count; i += 12)
                dt3.Rows.Add(cell[i].Text, cell[i + 1].Text, cell[i + 2].Text, cell[i + 3].Text, cell[i + 4].Text, cell[i + 5].Text, cell[i + 6].Text, cell[i + 7].Text, cell[i + 8].Text, cell[i + 9].Text, cell[i + 10].Text, cell[i + 11].Text);

            cell = driver.FindElements(By.XPath("//fieldset[3]/fieldset[2]/table/thead/tr/tr")).ToList();
            for (int i = 0; i < cell.Count; i += 9)
                dt3.Rows.Add(cell[i].Text, cell[i + 1].Text, cell[i + 2].Text, cell[i + 3].Text, cell[i + 4].Text, cell[i + 5].Text, cell[i + 6].Text, cell[i + 7].Text, cell[i + 8].Text);
            */
            driver.Navigate().GoToUrl("http://root:root@192.168.9.103/#/minerprofiles");

            Thread.Sleep(5000);
            List<IWebElement> cell;
            DataTable dti1 = new DataTable();
            DataTable dti2 = new DataTable();
            DataTable dti3 = new DataTable();
            for(int i = 0; i < 22; i++)
            {
                dti1.Columns.Add("p"+i);
                dti2.Columns.Add("p"+i);
                dti3.Columns.Add("p"+i);
            }
            cell = driver.FindElements(By.XPath("/html/body/div/div[2]/main/div/div/div[1]/div/fieldset/div[2]/div/div/div[1]/fieldset/div/table/tbody/tr/td/div/div/div/span")).ToList();
            int z = 0;
            for(int i = 0; i < 6; i++)
            {
                dti1.Rows.Add();
                for (int y = 0; y < 22; y++)
                {
                    if ((i == 0 || i == 2 || i == 3) && (y == 0 || y == 1))
                        dti1.Rows[i][y] = "-";
                    else
                    {
                        dti1.Rows[i][y] = cell[z];
                        z++;
                    }
                }
            }
            cell = driver.FindElements(By.XPath("/html/body/div/div[2]/main/div/div/div[1]/div/fieldset/div[2]/div/div/div[2]/fieldset/div/table/tbody/tr/td/div/div/div/span")).ToList();
            z = 0;
            for (int i = 0; i < 6; i++)
            {
                dti2.Rows.Add();
                for (int y = 0; y < 22; y++)
                {
                    if ((i == 0 || i == 2 || i == 3) && (y == 0 || y == 1))
                        dti2.Rows[i][y] = "-";
                    else
                    {
                        dti2.Rows[i][y] = cell[z];
                        z++;
                    }
                }
            }
            cell = driver.FindElements(By.XPath("/html/body/div/div[2]/main/div/div/div[1]/div/fieldset/div[2]/div/div/div[3]/fieldset/div/table/tbody/tr/td/div/div/div/span")).ToList();
            z = 0;
            for (int i = 0; i < 6; i++)
            {
                dti3.Rows.Add();
                for (int y = 0; y < 22; y++)
                {
                    if ((i == 0 || i == 2 || i == 3) && (y == 0 || y == 1))
                        dti3.Rows[i][y] = "-";
                    else
                    {
                        dti3.Rows[i][y] = cell[z];
                        z++;
                    }
                }
            }
        }


    }
}
