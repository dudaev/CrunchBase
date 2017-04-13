using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
//using GemBox.Spreadsheet;
using excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.PhantomJS;
using OpenQA.Selenium.Support.UI;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;

namespace CrunchBaseReady
{
    public partial class Form1 : Form
    {

        IWebDriver driver;

        OpenFileDialog openFile;
        excel.Application excelApp;
        excel.Workbook wb;
        excel.Worksheet ws;
        excel.Application excelAppExp = new excel.Application();
        object misValue = System.Reflection.Missing.Value;
        List<string> alldata = new List<string>();
        List<string> rowsNames = new List<string>();
        string currJobPage = null;
        public Form1()
        {
            InitializeComponent();

        }

        private void uploadExcelBtn_Click(object sender, EventArgs e)
        {
            openFile = new OpenFileDialog();
            excelApp = new excel.Application();
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                excel.Workbooks excelWB = excelApp.Workbooks;
                excelWB.Open(openFile.FileName);
                excel._Worksheet sheets = excelApp.ActiveSheet;

                int columnCount = sheets.UsedRange.Columns.Count;
                List<string> columnNames = new List<string>();
                for (int c = 1; c < columnCount; c++)
                {
                    if (sheets.Cells[1, c].Value2 != null)
                    {
                        string columnName = sheets.Columns[c].Address;
                        Regex reg = new Regex(@"(\$)(\w*):");
                        if (reg.IsMatch(columnName))
                        {
                            Match match = reg.Match(columnName);
                            columnNames.Add(match.Groups[2].Value);
                        }
                    }
                }
                foreach (var item in columnNames)
                {
                    clmnNames.Items.Add(item);
                }
            }
        }



        private void excelColumnBtn_Click(object sender, EventArgs e)
        {



            //////////////////////////////////////////////////////
            excel._Worksheet sheets = excelApp.ActiveSheet;
            excel.Range last = sheets.Cells.SpecialCells(excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            excel.Range range = sheets.get_Range("A1", last);
            var iTotalRows = sheets.UsedRange.Rows.Count;

            int row = 1;

            driver = new FirefoxDriver();
            //Loop get all Urls in column
            for (int rows = 1; rows < iTotalRows - 1; rows++)
            {
                //
                driver.Navigate().GoToUrl(sheets.Cells[row++, clmnNames.Text].value);
                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(5.00));
                IWebElement elem = driver.FindElement(By.XPath(".//*[@id='profile_header_heading']/a"));
                string s = elem.Text;
                listBox1.Items.Add(s); // Company Name
                alldata.Add(s);



                IList<IWebElement> founderList = driver.FindElements(By.XPath(".//*[@id='info-card-overview-content']/div/dl/div[2]/dd[3]/a"));
                IList<IWebElement> founderList1 = driver.FindElements(By.XPath(".//*[@id='info-card-overview-content']/div/dl/div[2]/dd[3]/a"));

                List<string> urls = new List<string>();
                List<string> founderName = new List<string>();

                foreach (var item in founderList)
                {
                    urls.Add(item.GetAttribute("href"));


                }
                foreach (var item in founderList)
                {
                    founderName.Add(item.Text);
                }

                bool present;
                foreach (var item2 in urls)
                {


                    driver.Navigate().GoToUrl(item2);
                    driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(3.00));
                    //1////////////////////////////////////////////////////founder Name//////////////////////////////////////////////////////////////////////////
                    listBox1.Items.Add(founderName);
                    alldata.Add(founderName.ToString());
                    //2////////////////////////////////////////////////////Founder crunch Page//////////////////////////////////////////////////////////////////////////
                    listBox1.Items.Add(urls);
                    alldata.Add(urls.ToString());

                    //3////////////////////////////////////////////////////Role//////////////////////////////////////////////////////////////////////////
                    try
                    {
                        driver.FindElement(By.XPath(".//*[@id='info-card-overview-content']/div/dl/div/dd[1]"));
                        listBox1.Items.Add(driver.FindElement(By.XPath(".//*[@id='info-card-overview-content']/div/dl/div/dd[1]")).Text); // Role
                        alldata.Add(driver.FindElement(By.XPath(".//*[@id='info-card-overview-content']/div/dl/div/dd[1]")).Text);
                        present = true;
                    }
                    catch (NoSuchElementException)
                    {
                        listBox1.Items.Add("Role element doesn't exist");
                        alldata.Add("Role element doesn't exist");
                        present = false;
                    }

                    //4////////////////////////////////////////////////////# of investments//////////////////////////////////////////////////////////////////////////
                    if (driver.FindElements(By.XPath(".//*[@id='info-card-overview-content']/div/dl/div/dd[2]/a")).Any())
                    {
                        listBox1.Items.Add(driver.FindElement(By.XPath(".//*[@id='info-card-overview-content']/div/dl/div/dd[2]/a")).Text); //# of investments
                        alldata.Add(driver.FindElement(By.XPath(".//*[@id='info-card-overview-content']/div/dl/div/dd[2]/a")).Text);
                    }
                    else
                    {
                        listBox1.Items.Add("Element doesn't exist.");
                        alldata.Add("# of investments element doesn't exist");
                    }
                    //5////////////////////////////////////////////////////Gender//////////////////////////////////////////////////////////////////////////
                    if (driver.FindElements(By.XPath(".//*[@id='info-card-overview-content']/div/dl/dd[2]")).Any())
                    {
                        listBox1.Items.Add(driver.FindElement(By.XPath(".//*[@id='info-card-overview-content']/div/dl/dd[2]")).Text); //Gender
                        alldata.Add(driver.FindElement(By.XPath(".//*[@id='info-card-overview-content']/div/dl/dd[2]")).Text);
                    }
                    else
                    {
                        listBox1.Items.Add("Element doesn't exist.");
                        alldata.Add("Gender element doesn't exist.");
                    }
                    //6////////////////////////////////////////////////////Location//////////////////////////////////////////////////////////////////////////
                    if (driver.FindElements(By.XPath(".//*[@id='info-card-overview-content']/div/dl/dd[3]")).Any())
                    {
                        listBox1.Items.Add(driver.FindElement(By.XPath(".//*[@id='info-card-overview-content']/div/dl/dd[3]")).Text);//Location
                        alldata.Add(driver.FindElement(By.XPath(".//*[@id='info-card-overview-content']/div/dl/dd[3]")).Text);
                    }
                    else
                    {
                        listBox1.Items.Add("Element doesn't exist.");
                        alldata.Add("Element location doesn't exist.");
                    }
                    //7//////////////////////////////////////////////////WebSite//////////////////////////////////////////////////////////////////////////
                    if (driver.FindElements(By.XPath(".//*[@id='info-card-overview-content']/div/dl/dd[4]/a")).Any())
                    {
                        listBox1.Items.Add(driver.FindElement(By.XPath(".//*[@id='info-card-overview-content']/div/dl/dd[4]/a")).Text);//Website
                        alldata.Add(driver.FindElement(By.XPath(".//*[@id='info-card-overview-content']/div/dl/dd[4]/a")).Text);
                    }
                    else
                    {
                        listBox1.Items.Add("Element doesn't exist.");
                        alldata.Add("Element Website doesn't exist.");
                    }
                    //8//9//10////////////////////////////////////////Social//////////////////////////////////////////////////////////////////////////
                    if (driver.FindElements(By.XPath(".//*[@id='info-card-overview-content']/div/dl/dd[5]/a")).Any())
                    {
                        List<IWebElement> socials = driver.FindElements(By.XPath(".//*[@id='info-card-overview-content']/div/dl/dd[5]/a")).ToList();
                        List<string> socialsUrls = new List<string>();
                        foreach (var item in socials)
                        {
                            socialsUrls.Add(item.GetAttribute("href"));
                        }
                        foreach (var item in socialsUrls)
                        {
                            listBox1.Items.Add(item);//Social
                            alldata.Add(item);
                        }
                    }
                    else
                    {
                        listBox1.Items.Add("Element doesn't exist.");
                        alldata.Add("Element Website doesn't exist.");
                    }
                    //11////////////////////////////////////////////////Personal details//////////////////////////////////////////////////////////////////////////
                    if (driver.FindElements(By.ClassName("description-ellipsis")).Any())
                    {
                        List<IWebElement> personal = driver.FindElements(By.ClassName("description-ellipsis")).ToList();
                        List<string> personalText = new List<string>();

                        foreach (var item in personal)
                        {
                            personalText.Add(item.Text);
                        }

                        foreach (var item in personalText)
                        {
                            listBox1.Items.Add(item);
                            alldata.Add(item);
                        }

                    }
                    else
                    {
                        listBox1.Items.Add("Element personal details doesn't exist.");
                        alldata.Add("Element personal details doesn't exist.");
                    }
                    //12//////////////////////////////////////////////////Current Job Company//////////////////////////////////////////////////////////////////////////
                    if (driver.FindElements(By.XPath(".//*[@id='main-content']/div[2]/div[2]/div[3]/div[4]/div[2]/div[2]/div[2]/h5[1]/a")).Any())
                    {
                        listBox1.Items.Add(driver.FindElement(By.XPath(".//*[@id='main-content']/div[2]/div[2]/div[3]/div[4]/div[2]/div[2]/div[2]/h5[1]/a")).Text);//Current job Company
                        alldata.Add(driver.FindElement(By.XPath(".//*[@id='main-content']/div[2]/div[2]/div[3]/div[4]/div[2]/div[2]/div[2]/h5[1]/a")).Text);
                    }
                    else
                    {
                        listBox1.Items.Add("Element Current Job doesn't exist.");
                        alldata.Add("Element Current Job doesn't exist.");
                    }
                    //13//////////////////////////////////////////////////Current Job Company Crunch URL/////////////////////////////////////////////////////////////////
                    if (driver.FindElements(By.XPath(".//*[@id='main-content']/div[2]/div[2]/div[3]/div[4]/div[2]/div[2]/div[2]/h5/a")).Any())
                    {
                        listBox1.Items.Add(driver.FindElement(By.XPath(".//*[@id='main-content']/div[2]/div[2]/div[3]/div[4]/div[2]/div[2]/div[2]/h5/a")).GetAttribute("href"));//Current job Company CrunchUrl
                        alldata.Add(driver.FindElement(By.XPath(".//*[@id='main-content']/div[2]/div[2]/div[3]/div[4]/div[2]/div[2]/div[2]/h5/a")).GetAttribute("href"));
                    }
                    else
                    {
                        listBox1.Items.Add("Element Current Job doesn't exist.");
                        alldata.Add("Element Current Job doesn't exist.");
                    }
                    //14//////////////////////////////////////////////////Current Job Company WebSite/////////////////////////////////////////////////////////////////
                    
                    if (driver.FindElements(By.XPath(".//*[@id='main-content']/div[2]/div[2]/div[3]/div[4]/div[2]/div[2]/div[2]/h5/a")).Any())
                    {
                        currJobPage = driver.FindElement(By.XPath(".//*[@id='main-content']/div[2]/div[2]/div[3]/div[4]/div[2]/div[2]/div[2]/h5/a")).GetAttribute("href");//Current job Company CrunchUrl
                        driver.Navigate().GoToUrl(currJobPage);
                        if(driver.FindElements(By.XPath(".//*[@id='info-card-overview-content']/div/dl/div[2]/dd[5]/a")).Any())
                        {
                            listBox1.Items.Add(driver.FindElement(By.XPath(".//*[@id='info-card-overview-content']/div/dl/div[2]/dd[5]/a")).GetAttribute("href"));
                            alldata.Add(driver.FindElement(By.XPath(".//*[@id='info-card-overview-content']/div/dl/div[2]/dd[5]/a")).GetAttribute("href"));
                        }
                        else
                        {
                            listBox1.Items.Add("Element Current Job Page doesn't exist.");
                            alldata.Add("Element Current Job Page doesn't exist.");
                        }
                        
                    }
                    else
                    {
                        listBox1.Items.Add("Element Current Job Page doesn't exist.");
                        alldata.Add("Element Current Job Page doesn't exist.");
                    }
                    ////////////////////////////////////////////////////End of page parse//////////////////////////////////////////////////////////////////////////
                    listBox1.Items.Add("-----------------------------------------------");
                }

                //wb = excelAppExp.Workbooks.Add(misValue);
                //ws = (excel.Worksheet)wb.Worksheets.get_Item(1);

                //rowsNames.Add("Investor Name");
                //rowsNames.Add("Investor Crunchbase Page");
                //rowsNames.Add("Primary Role");
                //rowsNames.Add("# Of Investments");
                //rowsNames.Add("Gender");
                //rowsNames.Add("Location");
                //rowsNames.Add("Website");
                //rowsNames.Add("LinkedIn");
                //rowsNames.Add("Facebook");
                //rowsNames.Add("ITwitter");
                //rowsNames.Add("Person details");
                //rowsNames.Add("Current Job Company");
                //rowsNames.Add("Current Job Company crunchbase URl Page");
                //rowsNames.Add("Current Job company website");

                //int alldatarow = 1;
                //for (int i = 1; i <= founderName.Count; i++)
                //{
                //    for (int j = 0; j < 14; j++)
                //    {
                //        ws.Cells[j, 1] = rowsNames;
                //        //ws.Cells[j, 2] = alldata;
                //    }
                //}


            }
            driver.Quit();
            excelApp.Quit();
            listBox1.Items.Add("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx");
            foreach (var item in alldata)
            {
                listBox1.Items.Add(alldata.ToList());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {



            wb.SaveAs("d:\\crunchbase.xls", excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);



            wb.Close(true, misValue, misValue);
            excelAppExp.Quit();

            Marshal.ReleaseComObject(ws);
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(excelAppExp);

            MessageBox.Show("Excel file created , you can find it!");


        }


    }
}


/////////////////////////

