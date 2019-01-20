namespace IronWebScraper_ConsoleApp
{
    public class ZCY_Item_bref
    {
        public string url;
        public string item_name;
        public string price;
        public ZCY_Item_bref(string target_item_name, string target_price, string target_url)
        {
            url = target_url;
            item_name = target_item_name;
            price = target_price;

            // TODO validate URL: has to be ZCY item page
        }
    }
}
