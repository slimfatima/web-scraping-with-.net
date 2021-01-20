using System;
using System.Linq;
using System.Collections.Generic;
using HtmlAgilityPack;
using ScrapySharp.Extensions;
using ScrapySharp.Network;
using scrapy.models;
using OfficeOpenXml;
using System.IO;

namespace scrapy
{
    class Program
    {
        static List<Security> securities;
        static Security security;
        static void Main(string[] args)
        {
            GetUrls("https://www.cbn.gov.ng/rates/GovtSecuritiesDrillDown.asp");
            SaveToExcel();
        }

        static void GetAuctionRows(string url){
            var page = new HtmlWeb();
            HtmlDocument doc = page.Load(url);
            security = new Security();

            var rows = doc.DocumentNode
            .Descendants("table")
            .FirstOrDefault(_=>_.Id.Equals("mytables"))
            .Descendants("tr");

            foreach(HtmlNode row in rows){
                saveRow(row);
            }
        }

        static void GetUrls(string url){
            var page = new HtmlWeb();
            HtmlDocument doc = page.Load(url);
            securities = new List<Security>();

            var links = doc.DocumentNode
            .Descendants("div")
            .FirstOrDefault(_=>_.Id.Equals("ContentTextinner"))
            .Descendants("a");

            foreach(HtmlNode link in links){
                GetAuctionRows("https://www.cbn.gov.ng/rates/" + link.Attributes["href"].Value);
            }

        }

        static void saveRow(HtmlNode row){
            var cell = GetCellValue(row);

            if(cell == ""){
                securities.Add(security);
                security = new Security();
            }else{
                var rowHeader = row.CssSelect("th b").FirstOrDefault().InnerText;

                switch (rowHeader){
                    case "Auction Date":
                        security.AuctiionDate = cell;
                        break;
                    case "Security Type":
                        security.SecurityType = cell;
                        break;
                    case "Tenor":
                        security.Tenor = cell;
                        break;
                    case "Auction No":
                        security.AuctionNumber = cell;
                        break;
                    case "Maturity Date":
                        security.MaturityDate = cell;
                        break;
                    case "Total Subscription":
                        security.TotalSubscription = cell;
                        break;
                    case "Total Successful":
                        security.TotalSuccessful = cell;
                        break;
                    case "Range Bid":
                        security.RangeBid = cell;
                        break;
                    case "Successful Bid Rates":
                        security.SuccessfulBidRate = cell;
                        break;
                    case "Description":
                        security.Description = cell;
                        break;
                    case "Rate":
                        security.Rate = cell;
                        break;
                    case "True Yield":
                        security.TrueYield = cell;
                        break;
                    case "Amount Offered (mn)":
                        security.AmountOffered = cell;
                        break;
                    default:

                        break;                
                }

            }
        }

        static string GetCellValue(HtmlNode row){
            return row.CssSelect("td").FirstOrDefault().InnerText;
        }

        public static void SaveToExcel(){
            using(ExcelPackage package = new ExcelPackage()){
                package.Workbook.Properties.Author = "Abdulgafar Jagun";

                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Securities");

                int sheetRow = 2;

                sheet.Cells[1,1].Value = "Auction Date";
                sheet.Cells[1,2].Value = "Security Type";
                sheet.Cells[1,3].Value = "Tenor";
                sheet.Cells[1,4].Value = "Auction No";
                sheet.Cells[1,5].Value = "Auction";
                sheet.Cells[1,6].Value = "Maturity Date";
                sheet.Cells[1,7].Value = "Total Subscription";
                sheet.Cells[1,8].Value = "Total Successful";
                sheet.Cells[1,9].Value = "Range Bid";
                sheet.Cells[1,10].Value = "Successful Bid Rates";
                sheet.Cells[1,11].Value = "Description";
                sheet.Cells[1,12].Value = "Rate";
                sheet.Cells[1,13].Value = "True Yield";
                sheet.Cells[1,14].Value = "Amount Offered (mn)";

                foreach(var security in securities){
                    sheet.Cells[sheetRow,1].Value = security.AuctiionDate;
                    sheet.Cells[sheetRow,2].Value = security.SecurityType;
                    sheet.Cells[sheetRow,3].Value = security.Tenor;
                    sheet.Cells[sheetRow,4].Value = security.AuctionNumber;
                    sheet.Cells[sheetRow,5].Value = security.Auction;
                    sheet.Cells[sheetRow,6].Value = security.MaturityDate;
                    sheet.Cells[sheetRow,7].Value = security.TotalSubscription;
                    sheet.Cells[sheetRow,8].Value = security.TotalSuccessful;
                    sheet.Cells[sheetRow,9].Value = security.Description;
                    sheet.Cells[sheetRow,10].Value = security.RangeBid;
                    sheet.Cells[sheetRow,11].Value = security.SuccessfulBidRate;
                    sheet.Cells[sheetRow,12].Value = security.Description;
                    sheet.Cells[sheetRow,13].Value = security.TrueYield;
                    sheet.Cells[sheetRow,14].Value = security.AmountOffered;

                    sheetRow++;
                }

                string fileName = "Securities.xlsx";
                //string path = Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.FullName;
                string path = Path.Combine(Directory.GetParent(Environment.CurrentDirectory).Parent.Parent.FullName, fileName);

                if(!File.Exists(path)){
                    FileInfo file = new FileInfo(path);
                    package.SaveAs(file);
                }else{
                    File.Delete(path);

                    FileInfo file = new FileInfo(path);
                    package.SaveAs(file);
                }

                
            }
        }

    }
}
