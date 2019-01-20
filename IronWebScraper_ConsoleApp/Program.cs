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
using OpenQA.Selenium.Support.UI;
using TheArtOfDev.HtmlRenderer;

namespace IronWebScraper_ConsoleApp
{

    class Program
    {
        // TODO 有信息，无库存
        // DEBUG: 不在购买范围，没有颜色分类
        // TODO timeout exceptions
        static void Main(string[] args)
        {
            //ZCY_Item one = new ZCY_Item("https://www.zcy.gov.cn/items/14855435");
            //one.Go();
            //one.Save();
            Console.WriteLine("抓取单个商品输入1后回车， 从搜索结果批量抓取输入2后回车: ");
            int mode = Convert.ToInt32(Console.ReadLine());
            while (!(mode == 1 || mode == 2))
            {
                Console.WriteLine("输入格式不正确：抓取单个商品输入1， 从搜索结果批量抓取输入2: ");
                mode = Convert.ToInt32(Console.ReadLine());
            }

            Console.WriteLine("输入政采云地址：");
            String ZCY_url = Console.ReadLine();
            while (ZCY_url.Length == 0)
            {
                Console.WriteLine("输入政采云地址：");
                ZCY_url = Console.ReadLine();
            }

            Console.WriteLine("输入保存到本地的路径，不输入则默认 D:\\网超\\");
            String destPath = Console.ReadLine();
            if (destPath.Length > 0)
            {
                bool exists = System.IO.Directory.Exists(destPath);
                if (!exists)
                {
                    System.IO.Directory.CreateDirectory(destPath);
                }
            } else
            {
                destPath = @"D:\网超\";
                bool exists = System.IO.Directory.Exists(destPath);
                if (!exists)
                {
                    System.IO.Directory.CreateDirectory(destPath);
                }
            }
            /*  https://www.zcy.gov.cn/items/14855435 */
            if (mode == 1)
            {
                ZCY_Item one;
                if (destPath.Length > 0)
                {
                    one = new ZCY_Item(ZCY_url, destPath);
                }
                 else
                {
                    one = new ZCY_Item(ZCY_url);
                }
                one.Go();
                one.Save();
                Console.WriteLine("抓取结束");
            }
            else if (mode == 2)
                {
            
                string ZCY_root = @"www.zcy.gov.cn";
                // string search_page_url = "https://www.zcy.gov.cn/eevees/shop?searchType=1&shopId=101149";
                string search_page_url = ZCY_url;
                //后台模式
                var options = new ChromeOptions();
                options.AddArguments("headless");
                options.SetLoggingPreference(LogType.Client, LogLevel.Off);

                var driver = new ChromeDriver(options);
                driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(30);
                driver.Navigate().GoToUrl(search_page_url);

                var page_selection = driver.FindElementByClassName("js-pagination");
                int total = Int32.Parse(page_selection.GetAttribute("data-total"));
                int size = Int32.Parse(page_selection.GetAttribute("data-size"));

                if (total == 0)
                {
                    Console.WriteLine("No item found");
                    return;
                }

                List<ZCY_Item_bref> list_complete = new List<ZCY_Item_bref>();

                int pages = (total + size - 1) / size; // Round UP, genius
                for (int i = 0; i < pages; i++)
                {
                    for (int limit = 0; limit < 2; limit++)
                    {
                        try
                        {
                            var item_table = driver.FindElementsByClassName("product");

                            foreach (var item in item_table)
                            {
                                IWebElement p_desc = item.FindElement(By.ClassName("product-desc")); ;
                                IWebElement link = p_desc.FindElement(By.TagName("a"));

                                string name = p_desc.Text;
                                string price = item.FindElement(By.ClassName("currency")).Text;
                                string url = link.GetAttribute("href");

                                //Regex regex = new Regex("^(.+?)");    //?后面的都不要
                                url = Regex.Replace(url, @"(\?.+)$", "");

                                ZCY_Item_bref item_bref = new ZCY_Item_bref(name, price, url);
                                list_complete.Add(item_bref);
                            }
                            break;// 尝试加载2次，每次失败等待1秒
                        }
                        catch (OpenQA.Selenium.StaleElementReferenceException e)
                        {
                            // TODO: 这里还没搞完啊
                            Console.WriteLine("你他妈这里又不对了,搞一下啊");
                            return;
                        }
                    }
                    var next_btn = driver.FindElementByClassName("next");
                    next_btn.Click();
                    System.Threading.Thread.Sleep(1000); // 这里等一秒，让网页加载完 TODO:把上面catch搞完
                }



                if (list_complete.Count() != total)
                {
                    // Debug purpose only
                    Console.WriteLine("Debug: list total number doesn't match " + list_complete.Count() + " : " + total);
                    return;
                }
            

                CreateExcelDoc_ZCY_Search_Page(list_complete, destPath);

                driver.Close();

                foreach (ZCY_Item_bref item in list_complete)
                {
                    ZCY_Item to_be_saved = new ZCY_Item(item.url, destPath);
                    try
                    {
                        to_be_saved.Go();
                        to_be_saved.Save();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                        Console.WriteLine("Dealing with: " + item.item_name);
                        throw e;
                    }
                }



                //==========================================================================================
                //var driver = new ChromeDriver();
                //driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(15);
                //driver.Navigate().GoToUrl("https://item.jd.com/112840.html");
                //img_jd(driver);

                //IJavaScriptExecutor js = driver;
                //var html2canvasJs = File.ReadAllText($"{GetAssemblyDirectory()}\\html2canvas.js");
                //js.ExecuteScript(html2canvasJs);
                //string generateScreenshotJS = @"function genScreenshot () {
                //                                     var canvasImgContentDecoded;
                //                                     html2canvas(document.body, {
                //                                       onrendered: function (canvas) {           

                //                                       window.canvasImgContentDecoded = canvas.toDataURL(""image/png"");
                //                                            window.open(myImage)
                //                                     }});
                //                                    }
                //                                    genScreenshot();";
                //js.ExecuteScript(generateScreenshotJS);

                //var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                //wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
                //wait.Until(
                //    wd =>
                //    {
                //        string response = (string)js.ExecuteScript
                //            ("return (typeof canvasImgContentDecoded === 'undefined' || canvasImgContentDecoded === null)");
                //        if (string.IsNullOrEmpty(response))
                //        {
                //            return false;
                //        }

                //        return bool.Parse(response);
                //    });
                //wait.Until(wd => !string.IsNullOrEmpty((string)js.ExecuteScript("return canvasImgContentDecoded;")));
                //var pngContent = (string)js.ExecuteScript("return canvasImgContentDecoded;");
                //pngContent = pngContent.Replace("data:image/png;base64,", string.Empty);
                //byte[] data = Convert.FromBase64String(pngContent);
                //var tempFilePath = Path.GetTempFileName().Replace(".tmp", ".png");
                //Image image;
                //using (var ms = new MemoryStream(data))
                //{
                //    image = Image.FromStream(ms);
                //}
                //image.Save(tempFilePath, ImageFormat.Png);
            }

        }


        private static string GetAssemblyDirectory()
        {
            string codeBase = Assembly.GetExecutingAssembly().CodeBase;
            var uri = new UriBuilder(codeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            return Path.GetDirectoryName(path);
        }
        public static void img_jd(ChromeDriver driver)
        {
            string root_path = @"D:\jd\";
            
            Regex file_name_validator = new Regex("[/\\?*<>:|]");
            WebClient client = new WebClient();     // 图片下载器



            //-----------------------文件夹-------------------------
            string item_name = driver.FindElementByClassName("sku-name").Text;  // 文件夹命名
            item_name =file_name_validator.Replace(item_name, " ");
            string folder = root_path + item_name;

            System.IO.Directory.CreateDirectory(folder);
            folder = folder + "\\";

            //=====================================================
            //System.IO.File.WriteAllText(@"D:\jd\hp\z.html", driver.PageSource);

            //=====================================================
            int number = 1;     // 命名图片

            //---------------------预览图------------------------------
            string img_small_id = "spec-list";
            var img_small_node = driver.FindElementById(img_small_id);
            IList<IWebElement> spec_list = img_small_node.FindElements(By.TagName("img"));

            foreach(var img in spec_list)
            {
                string img_small_src = img.GetAttribute("src");
                string img_large_src = img_small_src.Replace("com/n5", "com/n1");
                Regex img_url_finder = new Regex("s[0-9]+x[0-9]+_jfs");
                img_large_src = img_url_finder.Replace(img_large_src, "jfs");

                string appedix = System.IO.Path.GetExtension(img_large_src);    //获取图片格式

                // 处理图片
                //System.IO.File.AppendAllText(@"D:\test.txt", img_large_src + Environment.NewLine);
                Console.WriteLine(img_large_src + "|" + folder + "a商品预览图" + number + appedix);
                client.DownloadFile(new Uri(img_large_src), folder + "a商品预览图" + number + appedix);
                number++;
            }

            //----------------------商品信息图片-----------------------
            var details = driver.FindElementById("J-detail-content");

            //TODO 处理广告
            number = 1;
            if (details.FindElements(By.ClassName("ssd-module-wrap")).Count() == 0)
            {
                IList<IWebElement> imgs_detail = details.FindElements(By.TagName("img"));
                foreach (var img in imgs_detail)
                {
                    string src = img.GetAttribute("data-lazyload");
                    string appedix = System.IO.Path.GetExtension(src);  //获取图片格式
                    if (src.Substring(0, 2).Equals("//"))
                    {
                        src = "http:" + src;
                    }
                    Console.WriteLine(src + "|" + folder + number + appedix);
                    client.DownloadFile(new Uri(src), folder + number + appedix);
                    number++;
                }
            }
            else
            {            // TODO: 有ssd-module: 有ssd-widget-text？
                         
                         
                var ssd_module_wrap = details.FindElement(By.ClassName("ssd-module-wrap"));
                if(ssd_module_wrap.FindElements(By.ClassName("ssd-widget-text")).Count() == 0)
                {// 没有：找css
                    IList<IWebElement> ssd_module = ssd_module_wrap.FindElements(By.ClassName("ssd-module"));
                    foreach (var img in ssd_module)
                    {
                        string src = img.GetCssValue("background-image");
                        src = Regex.Replace(src, @"^(url)", "");
                        src = src.Trim('(', ')', '"');
                        Console.WriteLine(src);
                        if (src != "none" && src != null)
                        {
                            string appedix = System.IO.Path.GetExtension(src);  //获取图片格式
                            if (src.Substring(0, 2).Equals("//"))
                            {
                                src = "http:" + src;
                            }
                            Console.WriteLine(src + "|" + folder + "b商品详情" + number + appedix);
                            client.DownloadFile(new Uri(src), folder + "b商品详情" + number + appedix);
                            number++;
                        }
                        else
                        {
                            Console.WriteLine("DEBUG: ssd-module has no bg img.");
                            //这个ssd_module没有背景图片
                        }

                    }
                }
                else
                {//  有：截图？
                    // gif?
                    Console.WriteLine("???");

                    // 截图吧 截取全图，失败
                    //// Get the location of element on the page
                    //Screenshot screenshot = ((ITakesScreenshot)driver).GetScreenshot();
                    //screenshot.SaveAsFile(@"D:\testa.jpg");

                    //Rectangle rect = new Rectangle();
                    //if (details != null)
                    //{
                    //    Point point = details.Location;

                    //    // Get width and height of the element
                    //    int eleWidth = details.Size.Width;
                    //    int eleHeight = details.Size.Height;
                    //    // Create a rectangle using Width, Height and element location
                    //    rect = new Rectangle(point.X, point.Y, eleWidth, eleHeight);

                    //    Console.WriteLine("X: " + point.X + ", Y: " + point.Y + ", W: " + eleWidth + ", H: " + eleHeight);
                    //}
                    //// croping the image based on rect.
                    //Bitmap bmpImage = new Bitmap(screenshot_img);
                    //Image cropedImag = bmpImage.Clone(rect, bmpImage.PixelFormat);
                    //cropedImag.Save(@"D:\testb");


                    //-------------Html2Canvas 方案
                }
            }
        }
        private static void CreateExcelDoc_ZCY_Search_Page(List<ZCY_Item_bref> full_list, string root_path, string file_name = "商品列表")
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(root_path + "\\" + file_name + ".xlsx", SpreadsheetDocumentType.Workbook))
            {
                // TODO 如果root_path 不存在，创建文件夹

                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = file_name };

                sheets.Append(sheet);

                workbookPart.Workbook.Save();


                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                foreach(ZCY_Item_bref item in full_list)
                {
                    Row row = new Row();
                    row.Append(
                        ConstructCell(item.item_name, CellValues.String),
                        ConstructCell(item.price, CellValues.String),
                        ConstructCell(item.url, CellValues.String));

                    // Insert the header row to the Sheet Data
                    sheetData.AppendChild(row);
                }
                

                worksheetPart.Worksheet.Save();
            }
        }
        private static Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }
    }
}
