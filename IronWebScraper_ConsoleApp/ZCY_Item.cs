using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Collections.Generic;
using System.Net;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace IronWebScraper_ConsoleApp
{
    public class ZCY_Item
    {
        public string url;
        public string root_path;
        public string file_name;
        public string item_name;
        public string purchase_catalog;
        public List<string> sku_names = new List<string>();
        public List<string> sell_prices = new List<string>();
        public List<string> platform_prices = new List<string>();
        public List<string> attri_keys = new List<string>();
        public List<string> attri_values = new List<string>();
        public List<string> thumbnail_urls = new List<string>();
        public List<string> img_urls = new List<string>();

        public ZCY_Item(string target_url, string _root_path = @"D:\网超\")
        {
            url = target_url;
            root_path = _root_path;
        }

        public void Go()
        {
            if (!url_validate())
            {
                Console.WriteLine("URL failed");
            }
            //后台模式
            var options = new ChromeOptions();
            options.AddArguments("headless");
            options.SetLoggingPreference(LogType.Client, LogLevel.Off);

            ChromeDriver ZCY_driver = new ChromeDriver(options);

            for(int limit = 0; limit < 10; limit++)
            {
                try
                {
                    ZCY_driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
                    ZCY_driver.Navigate().GoToUrl(url);
                    break;
                }
                catch(OpenQA.Selenium.WebDriverTimeoutException)
                {
                    // Say something
                }
            }



            //获取商品名称
            item_name = ZCY_driver.FindElementById("js-item-name").Text;
            //验证文件名，去除非法字符
            Regex folder_name_validator = new Regex("[/\\?*<>:|]");
            file_name = folder_name_validator.Replace(item_name, " ");

            //获取 可申请采购目录
            try
            {
                purchase_catalog = ZCY_driver.FindElementByClassName("catalog-list").Text;
            }
            catch (OpenQA.Selenium.NoSuchElementException)
            {
                purchase_catalog = "暂无目录";
            }

            

            //获取 SKUs
            var skus = ZCY_driver.FindElementsByClassName("js-sku-attr");
            foreach (var element in skus)
            {
                // 获取商品属性名称
                sku_names.Add(element.GetAttribute("title"));
                element.Click();

                string disabled = element.GetAttribute("disabled");
                if (disabled == null || !disabled.Equals("true"))
                {
                    // 获取属性对应价格
                    sell_prices.Add(ZCY_driver.FindElementById("js-item-price").Text);
                    platform_prices.Add(ZCY_driver.FindElementById("js-item-platform-price").Text);
                }
                else
                {
                    // 0库存，精确价格未知
                    sell_prices.Add("无库存");
                    platform_prices.Add("无库存");
                }
            }

            //----------------------------------获取额外属性-------------------------------
            var attributes = ZCY_driver.FindElementByClassName("js-other-attributes-box");
            string data = attributes.GetAttribute("data-attrs");

            Regex regex = new Regex("(attrKey\":\".+?\",)");    //获取属性名称
            foreach (Match match in regex.Matches(data))
            {
                string result = Regex.Replace(match.Value, "^(attrKey\":\")", "");
                result = Regex.Replace(result, "(\",)$", "");
                attri_keys.Add(result);
            }
            Regex regex2 = new Regex("(attrVal\":\".+?\",)");   //获取属性内容
            foreach (Match match in regex2.Matches(data))
            {
                string result = Regex.Replace(match.Value, "^(attrVal\":\")", "");
                result = Regex.Replace(result, "(\",)$", "");
                attri_values.Add(result);
            }
            //---------------------------获取图片地址-----------------------------------
            var thumbnails = ZCY_driver.FindElementsByClassName("thumbnail");
            foreach (var thumbnail in thumbnails)
            {
                string url = thumbnail.GetAttribute("data-src");
                if (url != null && url.Length > 0)
                {
                    thumbnail_urls.Add(url);
                }
            }
            var imgs = ZCY_driver.FindElementByClassName("goods-img").FindElements(By.TagName("img"));
            foreach (var img in imgs)
            {
                string url = img.GetAttribute("src");
                if (url != null)
                {
                    Uri uriResult;
                    bool result = Uri.TryCreate(url, UriKind.Absolute, out uriResult)
                        && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
                    if (result)
                    {
                        img_urls.Add(url);
                    }
                }
            }
            ZCY_driver.Close();
            ZCY_driver.Dispose();

        }

        public void Save()
        {
            //-----------------创建文件夹--------------------------------------
            root_path = root_path + file_name;
            System.IO.Directory.CreateDirectory(root_path); //DEBUG: check folder exits

            //-----------------创建Excel--------------------------------------
            //SpreadsheetDocument document = CreateWorkbook(@"D:\网超\" + item_name + "\\" + item_name + ".xlsx");

            // BUG: 正在使用excel Exception
            CreateExcelDoc();

            //-------------------下载图片--------------------------TODO
            WebClient client = new WebClient();     // 图片下载器

            int number = 1;
            foreach (var src in thumbnail_urls)
            {
                string appedix = System.IO.Path.GetExtension(src);    //获取图片格式
                client.DownloadFile(new Uri(src), root_path + "\\" + "a商品预览图" + number + appedix);
                number++;
            }
            number = 1;
            foreach (var src in img_urls)
            {
                string appedix = System.IO.Path.GetExtension(src);    //获取图片格式
                client.DownloadFile(new Uri(src), root_path + "\\" + "b商品详情" + number + appedix);
                number++;
            }
            client.Dispose();
        }
        private Boolean url_validate()
        {
            //确认网页地址是政采云商品地址 TODO
            return true;
        }
        private void CreateExcelDoc()
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(root_path + "\\" + file_name + ".xlsx", SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "商品名称" };

                sheets.Append(sheet);

                workbookPart.Workbook.Save();


                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                // Constructing header
                Row row1 = new Row();
                row1.Append(
                    ConstructCell("商品名称", CellValues.String),
                    ConstructCell(item_name, CellValues.String));

                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row1);

                Row row2 = new Row();
                row2.Append(
                    ConstructCell("可申请采购目录", CellValues.String),
                    ConstructCell(purchase_catalog, CellValues.String));
                sheetData.AppendChild(row2);

                sheetData.AppendChild(new Row());

                Row row3 = new Row();
                row3.Append(
                    ConstructCell("属性名称", CellValues.String),
                    ConstructCell("销售价", CellValues.String),
                    ConstructCell("电商平台价", CellValues.String));
                sheetData.AppendChild(row3);

                int count = 0;
                foreach (string sku in sku_names)
                {
                    Row row = new Row();
                    row.Append(
                    ConstructCell(sku, CellValues.String),
                    ConstructCell(sell_prices[count], CellValues.String),
                    ConstructCell(platform_prices[count], CellValues.String));
                    sheetData.AppendChild(row);
                    count++;
                }


                sheetData.AppendChild(new Row());
                Row row4 = new Row();
                row4.Append(
                    ConstructCell("额外参数", CellValues.String));
                sheetData.AppendChild(row4);

                count = 0;
                foreach (string attri in attri_keys)
                {
                    Row row = new Row();
                    row.Append(
                    ConstructCell(attri_keys[count], CellValues.String),
                    ConstructCell(attri_values[count], CellValues.String));
                    sheetData.AppendChild(row);
                    count++;
                }

                worksheetPart.Worksheet.Save();
            }
        }
        private Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }
    }
}
