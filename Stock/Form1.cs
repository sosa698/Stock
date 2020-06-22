/* 基本功能:
 * 用近兩個月畫出均線與三價線
 *
 * 自動找出均線多頭排列的股票。 
 * 
 * RowData:
 * 所有ID : https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E9%A3%9F%E5%93%81&form=menu&form_id=stock_id&form_name=stock_name&domain=0  //可更改食品攔
 * 各ID資料:https://tw.quote.finance.yahoo.net/quote/q?type=ta&perd=m&mkt=10&sym=00677U&v=1&callback=jQuery111308610472701998613_1584972510198&_=1584972510200  //可改ID
 *                                                                                                                                                    //perd=m :月線 , d:日線 , w:均線
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using System.Collections; // IComparer interface 在 System.Collections namespace 內
using System.Windows.Forms.DataVisualization.Charting;

namespace Stock
{

    public partial class Form1 : Form
    {
        /****************************************************************************全域變數************************************************************/

        string[] url =
                {
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E6%B0%B4%E6%B3%A5&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//水泥
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E9%A3%9F%E5%93%81&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//食品
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E5%A1%91%E8%86%A0&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//塑膠
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E7%B4%A1%E7%B9%94&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//紡織
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E9%9B%BB%E6%A9%9F&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//電機
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E9%9B%BB%E5%99%A8%E9%9B%BB%E7%BA%9C&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//電器電纜
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E5%8C%96%E5%AD%B8&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//化學
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E7%94%9F%E6%8A%80%E9%86%AB%E7%99%82&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//生技醫療
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E7%8E%BB%E7%92%83&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//玻璃
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E9%80%A0%E7%B4%99&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//造紙
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E9%8B%BC%E9%90%B5&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//鋼鐵     10
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E6%A9%A1%E8%86%A0&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//橡膠
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E6%B1%BD%E8%BB%8A&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//汽車
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E5%8D%8A%E5%B0%8E%E9%AB%94&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//半導體
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E9%9B%BB%E8%85%A6%E9%80%B1%E9%82%8A&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//電腦周邊
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E5%85%89%E9%9B%BB&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//光電
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E9%80%9A%E4%BF%A1%E7%B6%B2%E8%B7%AF&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//通信網路
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E9%9B%BB%E5%AD%90%E9%9B%B6%E7%B5%84%E4%BB%B6&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//電子零組件
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E9%9B%BB%E5%AD%90%E9%80%9A%E8%B7%AF&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//電子通路
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E8%B3%87%E8%A8%8A%E6%9C%8D%E5%8B%99&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//資訊服務
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E5%85%B6%E5%AE%83%E9%9B%BB%E5%AD%90&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//其他電子    20
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E7%87%9F%E5%BB%BA&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//營建
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E8%88%AA%E9%81%8B&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//航運
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E8%A7%80%E5%85%89&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//觀光
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E9%87%91%E8%9E%8D&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//金融
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E8%B2%BF%E6%98%93%E7%99%BE%E8%B2%A8&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//貿易百貨
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E6%B2%B9%E9%9B%BB%E7%87%83%E6%B0%A3&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//油電燃氣
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E5%AD%98%E8%A8%97%E6%86%91%E8%AD%89&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//存託憑證
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=ETF&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//ETF
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E5%8F%97%E7%9B%8A%E8%AD%89%E5%88%B8&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//受益證卷
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=ETN&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//ETN    30
                    "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E5%85%B6%E4%BB%96&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//其他
            //        "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E5%B8%82%E8%AA%8D%E8%B3%BC&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//市認證
           //         "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E5%B8%82%E8%AA%8D%E5%94%AE&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//市認售
           //         "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E6%8C%87%E6%95%B8%E9%A1%9E&form=menu&form_id=stock_id&form_name=stock_name&domain=0",//指數類
           //         "https://tw.stock.yahoo.com/h/kimosel.php?tse=1&cat=%E5%B8%82%E7%89%9B%E8%AD%89&form=menu&form_id=stock_id&form_name=stock_name&domain=0"//市牛證
            };  //各種分類股票，如果未來yahoo分類有增加，就在這邊新增網址

        //上櫃
        string[] url2 =
        {
            "no",                                                                                                                                      //no 水泥                 
            "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E9%A3%9F%E5%93%81&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //食品
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E5%A1%91%E8%86%A0&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //塑膠
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E7%B4%A1%E7%B9%94&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //紡織
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E9%9B%BB%E6%A9%9F&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //電機
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E9%9B%BB%E5%99%A8&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //電器
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E5%8C%96%E5%B7%A5&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //化工
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E7%94%9F%E6%8A%80&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //生技
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E7%8E%BB%E7%92%83&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //玻璃
            "no",                                                                                                                                      //no造紙
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E9%8B%BC%E9%90%B5&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //鋼鐵
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E6%A9%A1%E8%86%A0&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //橡膠
            "no",                                                                                                                                      //no汽車
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E5%8D%8A%E5%B0%8E&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //半導體
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E9%9B%BB%E8%85%A6&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //電腦
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E5%85%89%E9%9B%BB&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //光電
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E9%80%9A%E4%BF%A1&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //通信
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E9%9B%BB%E9%9B%B6&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //電零
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E9%80%9A%E8%B7%AF&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //通路
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E8%B3%87%E6%9C%8D&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //資訊服務
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E4%BB%96%E9%9B%BB&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //其他電子
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E7%87%9F%E5%BB%BA&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //營建
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E8%88%AA%E9%81%8B&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //航運
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E8%A7%80%E5%85%89&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //觀光
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E9%87%91%E8%9E%8D&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //金融
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E8%B2%BF%E6%98%93&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //貿易
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E6%B2%B9%E9%9B%BB&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //油電
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E6%96%87%E5%89%B5&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //文創
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E8%BE%B2%E7%A7%91&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //農科
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E9%9B%BB%E5%95%86&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //電商
	        "https://tw.stock.yahoo.com/h/kimosel.php?tse=2&cat=%E6%AB%83%E5%85%B6%E4%BB%96&form=menu&form_id=stock_id&form_name=stock_name&domain=0", //其他

        };
        public struct Stock_Parameter
        {
            public string data_source;
            public string d_w_m;
            public int MA1;
            public int MA2;
            public int MA3;
            public int MA4;
            //---MACD
            public int EMA1;
            public int EMA2;
            public int MACD;
            //---CCI
            public int CCI_day;
            public int CCI_condition;
            //---Quantity
            public UInt32 Quantity;
        };
        Stock_Parameter stock_parameter;



        enum Grid
        {
            ID = 0,
            Date,
            Value,
            Find_days,
            day_status,
            MA15_status,
            CCI_last,
            EPS,
        }

        enum Grid_NoBackTest
        {
            ID = 0,
            Value,
            day_condition,
            week_condition,
            Today_status,
            EPS,
            Quantity,
        };


        private class stock
        {
            public string ID;
            public string name;
        };
        private List<stock> all_stock = new List<stock>();    //存所有股票ID+名字，不分類

        class data                         //每筆資料包含日期,收盤價,均線
        {
            public double high;
            public double low;
            public double open; //開盤
            public string date;
            public double value;
            public double average_MA1;
            public double average_MA2;
            public double average_MA3;
            public double average_MA4;

            //----For MACD
            public double DI;
            public double EMA12;
            public double EMA26;
            public double DIF;
            public double MACD;
            public double OSC;
            //-----For KD
            public double RSV;
            public double K;
            public double D;

            //-----For CCI
            public double CCI;

        }
        List<data> day_data = new List<data>();     //日線
        List<data> week_data = new List<data>();     //週線
        List<data> month_data = new List<data>();     //年線
                                                      //所有股票共用三個buffer，每一支算完均線大小後就釋放

        class TLB
        {
            public string date;
            public double up;
            public double down;
            public int status;
        }
        List<TLB> TLB_data = new List<TLB>();

        public delegate void Simulation_Way(string ID, string d_w_m);
        public const double earn_percent = 1.013;
        /****************************************************************************全域變數************************************************************/

        public Form1()
        {
            InitializeComponent();

            cbx_simulation_way.Items.Add("3日線轉正");
            cbx_simulation_way.Items.Add("新三價線");
            cbx_simulation_way.Items.Add("多頭排列");
            cbx_simulation_way.Items.Add("低於5%日線");
            cbx_simulation_way.Items.Add("MACD綠");
            cbx_simulation_way.Items.Add("MACD紅");
            cbx_simulation_way.Items.Add("MACD線轉正");
            cbx_simulation_way.Items.Add("CCI(順勢指標)");
            cbx_simulation_way.Items.Add("不回測");
            cbx_simulation_way.SelectedIndex = 0;

            cbx_line.Items.Add("日線");
            cbx_line.Items.Add("周線");
            cbx_line.Items.Add("月線");
            cbx_line.SelectedIndex = 1;

            cbx_data_source.SelectedIndex = 1;

            tbx_MA1.Text = "4";
            tbx_MA2.Text = "6";
            tbx_MA3.Text = "8";
            tbx_MA4.Text = "15";
            tbx_CCI.Text = "10";
            tbx_CCI_condition.Text = "-135";

            tbx_quantity.Text = "500";

            cbx_IsFilter.Checked = true;



        }

        /****
        * Initial 
        * @Function:按鈕按下後，先初始參數
        * @input    
        * @output   
        ****/
        public void Initial()
        {
            string[] line_tmp = { "d", "w", "m" };

            int line_index_tmp = cbx_line.SelectedIndex;
            int data_source_index = cbx_data_source.SelectedIndex;

            stock_parameter.data_source = cbx_data_source.Text;
            stock_parameter.d_w_m = line_tmp[line_index_tmp];
            stock_parameter.MA1 = Convert.ToInt16(tbx_MA1.Text);
            stock_parameter.MA2 = Convert.ToInt16(tbx_MA2.Text);
            stock_parameter.MA3 = Convert.ToInt16(tbx_MA3.Text);
            stock_parameter.MA4 = Convert.ToInt16(tbx_MA4.Text);

            stock_parameter.EMA1 = 8;
            stock_parameter.EMA2 = 13;
            stock_parameter.MACD = 9;

            stock_parameter.CCI_day = Convert.ToInt16(tbx_CCI.Text); ;
            stock_parameter.CCI_condition = Convert.ToInt16(tbx_CCI_condition.Text);

            stock_parameter.Quantity = Convert.ToUInt32(tbx_quantity.Text);
            //-----Clear all data
            day_data.Clear();
            week_data.Clear();
            month_data.Clear();
            TLB_data.Clear();
            all_stock.Clear();
            Grid_TLB.Rows.Clear();
            Grid_NoBack.Rows.Clear();
        }


        /****
        * openTxt 
        * @Function:讀取txt黨內容，並解析日線
        * @input    filePath:txt路徑
        * @output   day_data將被更新  return flase: fail
        ****/
        private bool openTxt(string filePath)
        {
            //透過路徑讀取檔案
            try
            {
                StreamReader sr = new StreamReader(filePath);
                string line = sr.ReadLine(); //第一行總是跳過
                while ((line = sr.ReadLine()) != null)
                {
                    data tmp = new data();
                    string date = line.Substring(0, 3);
                    date = (Convert.ToInt64(date) + 1911).ToString();
                    date += line.Substring(4, 2) + line.Substring(7, 2);
                    line = line.Substring(9);
                    double value = Convert.ToDouble(line);
                    tmp.date = date;
                    tmp.value = value;

                    day_data.Add(tmp);
                }
                sr.Close();
                return true;
            }
            catch { return false; }

        }






        /****
        * GetWebContent 
        * @Function:取得url的原始碼
        * @input    Url:欲取得原始碼網址  
        * @output   網址原始碼(string)
        ****/
        public static string GetWebContent(string Url)
        {
            string urlAddress = Url;
            string result = "";
            try
            {

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlAddress);
                request.Method = "GET";
                request.ContentType = "text/html;charset=big5";
                request.Timeout = 2500;

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                if (response.StatusCode == HttpStatusCode.OK)
                {
                    Stream receiveStream = response.GetResponseStream();
                    StreamReader readStream = null;
                    if (response.CharacterSet == null)
                        readStream = new StreamReader(receiveStream);
                    else
                        readStream = new StreamReader(receiveStream, Encoding.GetEncoding(response.CharacterSet));
                    result = readStream.ReadToEnd();
                    response.Close();
                    readStream.Close();
                }

                /*
                WebClient wc = new WebClient();

                // 加上header，非必要，有些網站會吃header
                // https://zh.wikipedia.org/wiki/HTTP%E5%A4%B4%E5%AD%97%E6%AE%B5

                wc.Headers.Clear();
                var page = wc.DownloadString(Url);
                result = page.ToString();
                */
            }
            catch (Exception ex)
            {
                string err = ex.Message;
            }
            return result;
        }

        /****
        * Get_All_ID 
        * @Function:解析原始碼，取得股票所有ID
        * @input    會自動取全域變數url解析
        * @output   ID資料將存在 List<stock> all_stock;
        ****/
        private void Get_All_ID()
        {
            List<int> Select_ID = new List<int>();

            if (checkBox1.Checked == true) Select_ID.Add(0); if (checkBox2.Checked == true) Select_ID.Add(1); if (checkBox3.Checked == true) Select_ID.Add(2); if (checkBox4.Checked == true) Select_ID.Add(3);
            if (checkBox5.Checked == true) Select_ID.Add(4); if (checkBox6.Checked == true) Select_ID.Add(5); if (checkBox7.Checked == true) Select_ID.Add(6); if (checkBox8.Checked == true) Select_ID.Add(7);
            if (checkBox9.Checked == true) Select_ID.Add(8); if (checkBox10.Checked == true) Select_ID.Add(9); if (checkBox11.Checked == true) Select_ID.Add(10); if (checkBox12.Checked == true) Select_ID.Add(11);
            if (checkBox13.Checked == true) Select_ID.Add(12); if (checkBox14.Checked == true) Select_ID.Add(13); if (checkBox15.Checked == true) Select_ID.Add(14); if (checkBox16.Checked == true) Select_ID.Add(15);
            if (checkBox17.Checked == true) Select_ID.Add(16); if (checkBox18.Checked == true) Select_ID.Add(17); if (checkBox19.Checked == true) Select_ID.Add(18); if (checkBox20.Checked == true) Select_ID.Add(19);
            if (checkBox21.Checked == true) Select_ID.Add(20); if (checkBox22.Checked == true) Select_ID.Add(21); if (checkBox23.Checked == true) Select_ID.Add(22); if (checkBox24.Checked == true) Select_ID.Add(23);
            if (checkBox25.Checked == true) Select_ID.Add(24); if (checkBox26.Checked == true) Select_ID.Add(25); if (checkBox27.Checked == true) Select_ID.Add(26); if (checkBox28.Checked == true) Select_ID.Add(27);
            if (checkBox29.Checked == true) Select_ID.Add(28); if (checkBox30.Checked == true) Select_ID.Add(29); if (checkBox31.Checked == true) Select_ID.Add(30); if (checkBox32.Checked == true) Select_ID.Add(31);


            string[] use_url = null;
            if (checkBox36.Checked == false)
            {
                use_url = url;  //上市
            }
            else
            {
                use_url = url2;  //上櫃
            }

            for (int i = 0; i < Select_ID.Count; i++)
            {
                string ID_RowData = GetWebContent(use_url[Select_ID[i]]);
                while (true)
                {
                    stock tmp = new stock();
                    int java_index = ID_RowData.IndexOf("javascript:setid"); if (java_index == -1) break;
                    ID_RowData = ID_RowData.Substring(java_index + 18);
                    int before_ID_Index = ID_RowData.IndexOf(',');
                    int before_name_Index = ID_RowData.IndexOf("');");
                    string ID_tmp = ID_RowData.Substring(0, before_ID_Index - 1);
                    string name_tmp = ID_RowData.Substring(before_ID_Index + 2, before_ID_Index - 3);
                    tmp.ID = ID_tmp;
                    tmp.name = name_tmp;

                    all_stock.Add(tmp);
                }

            }



        }

        /****
        * Get_All_data 
        * @Function:解析原始碼，取得該股票ID的日線周線月線資料(yahoo股市大多只有到過去3年資料)
        * @input    ID: 股票ID
        * @output   該ID資料將存在 List<data> day_data//日線 ,  List<data> day_week//週線    ,   List<data> day_year//年線
        * Note      因為buffer共用，取完資料就要對資料做parser，不能直接在取下一筆ID資料，否則上筆資料會被覆蓋。
        ****/
        private void Get_All_data(string ID)
        {
            string parsing_source = stock_parameter.data_source;

            string url = "";
            string html = "";

            string value = "";
            string date = "";

            string[] perd = { "d", "w", "m" };

            data data_tmp = null;

            bool Is_Have_Txt = false;

            //先刪掉所有資料  否則讀下一個ID時會繼續新增在後面
            day_data.Clear();
            week_data.Clear();
            month_data.Clear();

            //           perd[0] = stock_parameter.d_w_m;

            for (int times = 0; times < 2; times++)     //一種ID做三次 因為有日週月線
            {
                if (parsing_source == "Yahoo")
                {
                    if (times == 0) Is_Have_Txt = openTxt(@"d:\stock_data\" + ID + ".txt");

                    url = "https://tw.quote.finance.yahoo.net/quote/q?type=ta&perd=" + perd[times] + "&mkt=10&sym=" + ID + "&v=1&callback=jQuery111308610472701998613_1584972510198&_=1584972510200";   // yahoo
                    html = GetWebContent(url);

                    int tmp = html.IndexOf(":[{"); if (tmp == -1) break;
                    html = html.Substring(tmp);         //a從這開始是 :日期,value...
                    if (perd[times] == "d")    //日線分開做
                    {
                        if (Is_Have_Txt == true)//有TW資料
                        {
                            int last_TWindex = day_data.Count - 1;
                            string TW_date = day_data[last_TWindex].date;

                            tmp = html.IndexOf(TW_date); if (tmp == -1) return;
                            html = html.Substring(tmp);
                        }
                    }
                    while (true)            //新增找最高與最低
                    {
                        data_tmp = new data();

                        int time_index = html.IndexOf("t"); //尋找time 
                        if (time_index == -1) break;
                        html = html.Substring(time_index);
                        date = html.Substring(3, 8);        //get date

                        int value_index = html.IndexOf("c"); //尋找收盤價
                        html = html.Substring(value_index);
                        int before_value = html.IndexOf(",");

                        value = html.Substring(3, before_value - 3);

                        data_tmp.date = date;
                        data_tmp.value = Convert.ToDouble(value);

                        if (perd[times] == "d") day_data.Add(data_tmp);
                        else if (perd[times] == "w") week_data.Add(data_tmp);
                        else if (perd[times] == "m") month_data.Add(data_tmp);

                        if (perd[times] == "d" && Is_Have_Txt == true) //更新txt資料
                        {
                            StreamWriter sw = new StreamWriter(@"d:\stock_data\" + ID + ".txt", true);
                            string year_tmp = date.Substring(0, 4);
                            string month_tmp = date.Substring(4, 2);
                            string day_tmp = date.Substring(6, 2);
                            string write_data = (Convert.ToInt64(year_tmp) - 1911).ToString() + "/" + month_tmp + "/" + day_tmp + " " + value;
                            sw.WriteLine(write_data);
                            sw.Close();
                        }

                    }
                }
                else  // Money DJ
                {
                    //      url = "https://stock.capital.com.tw/z/BCD/czkc1.djbcd?a=" + ID + "&b=" + perd[times] + "&c=2880&E=1&ver=5";   // money dj 他還有ver 如果以後他更新可能會錯
                    url = "https://moneydj.emega.com.tw/z/BCD/czkc1.djbcd?a=" + ID + "&b=" + perd[times] + "&c=2880&E=1&ver=5";
                    html = GetWebContent(url);

                    //--step 1 先取所有日期
                    int tmp_index = 0;
                    int data_length = 0;
                    while (true)
                    {
                        data_tmp = new data();

                        tmp_index = html.IndexOf("/"); if (tmp_index == -1) break;  //確認後面還有日期

                        date = html.Substring(0, 4) + html.Substring(5, 2) + html.Substring(8, 2);

                        int comma = html.IndexOf(",");
                        html = html.Substring(11);

                        data_tmp.date = date;

                        if (perd[times] == "d") day_data.Add(data_tmp);
                        else if (perd[times] == "w") week_data.Add(data_tmp);
                        else if (perd[times] == "m") month_data.Add(data_tmp);
                        data_length++;
                    }


                    for (int kind = 0; kind < 4; kind++)  //取四種:開 高  低  收   //視情況拿掉open
                    {
                        for (int count = 0; count < data_length; count++)
                        {

                            int comma = html.IndexOf(",");
                            if (count == data_length - 1) comma = html.IndexOf(" ");   //最後一筆要特別處理

                            value = html.Substring(0, comma);
                            html = html.Substring(comma + 1);

                            if (perd[times] == "d")
                            {
                                if (kind == 0) day_data[count].open = Convert.ToDouble(value);
                                else if (kind == 1) day_data[count].high = Convert.ToDouble(value);
                                else if (kind == 2) day_data[count].low = Convert.ToDouble(value);
                                else if (kind == 3) day_data[count].value = Convert.ToDouble(value);
                            }
                            else if (perd[times] == "w")
                            {
                                if (kind == 0) week_data[count].open = Convert.ToDouble(value);
                                else if (kind == 1) week_data[count].high = Convert.ToDouble(value);
                                else if (kind == 2) week_data[count].low = Convert.ToDouble(value);
                                else if (kind == 3) week_data[count].value = Convert.ToDouble(value);
                            }
                            else if (perd[times] == "m")
                            {
                                if (kind == 0) month_data[count].open = Convert.ToDouble(value);
                                else if (kind == 1) month_data[count].high = Convert.ToDouble(value);
                                else if (kind == 2) month_data[count].low = Convert.ToDouble(value);
                                else if (kind == 3) month_data[count].value = Convert.ToDouble(value);
                            }
                        }
                    }
                    //-----------最後一筆 用yahoo的  因為他有即時更新 日線 周線 都更新       

                    {
                        int Is_Add_data = 0;    //0:不新增 1:新增 2:改在最後一筆
                        data_tmp = new data();

                        if (data_length <= 1) break;


                        //yahoo 不知道為啥 有時候會輸出資料不完整 try 3次
                        string html_yahoo = "";
                        int tmp = 0;

                        for (int try_ = 0; try_ < 3; try_++)
                        {
                            Thread.Sleep(500);

                            url = "https://tw.quote.finance.yahoo.net/quote/q?type=ta&perd=" + perd[times] + "&mkt=10&sym=" + ID + "&v=1";   // yahoo
                            html_yahoo = GetWebContent(url);

                            tmp = html_yahoo.LastIndexOf("t"); if (tmp == -1) break;
                            html_yahoo = html_yahoo.Substring(tmp);         //直接跳到最後一筆資料

                            if (html_yahoo[1] == 'a')
                                continue;
                            else
                                break;

                        }

                        try
                        {
                            date = html_yahoo.Substring(3, 8);
                        }
                        catch { continue; } //放棄使用yahoo最新資料


                        if (perd[times] == "d")
                        {
                            string date_MoneyDJ = day_data[data_length - 1].date;
                            if (date != date_MoneyDJ) //資料不一樣 新增
                            {
                                Is_Add_data = 1;
                            }
                        }
                        else if (perd[times] == "w")
                        {
                            string date_MoneyDJ = week_data[data_length - 1].date;
                            string date_test = date.Substring(0, 4) + "-" + date.Substring(4, 2) + "-" + date.Substring(6, 2);
                            string date_test1 = date_MoneyDJ.Substring(0, 4) + "-" + date_MoneyDJ.Substring(4, 2) + "-" + date_MoneyDJ.Substring(6, 2);
                            DateTime sDate = Convert.ToDateTime(date_test);
                            DateTime eDate = Convert.ToDateTime(date_test1);
                            TimeSpan ts = sDate - eDate;
                            double delta = ts.TotalDays;
                            if (delta > 6)
                            {
                                Is_Add_data = 1;
                            }
                            else
                            {
                                Is_Add_data = 2;
                            }
                        }

                        if (Is_Add_data != 0)
                        {
                            data_tmp.date = date;

                            tmp = html_yahoo.IndexOf("o");
                            html_yahoo = html_yahoo.Substring(tmp);
                            tmp = html_yahoo.IndexOf(",");
                            value = html_yahoo.Substring(3, tmp - 3);
                            data_tmp.open = Convert.ToDouble(value);

                            tmp = html_yahoo.IndexOf("h");
                            html_yahoo = html_yahoo.Substring(tmp);
                            tmp = html_yahoo.IndexOf(",");
                            value = html_yahoo.Substring(3, tmp - 3);
                            data_tmp.high = Convert.ToDouble(value);

                            tmp = html_yahoo.IndexOf("l");
                            html_yahoo = html_yahoo.Substring(tmp);
                            tmp = html_yahoo.IndexOf(",");
                            value = html_yahoo.Substring(3, tmp - 3);
                            data_tmp.low = Convert.ToDouble(value);

                            tmp = html_yahoo.IndexOf("c");
                            html_yahoo = html_yahoo.Substring(tmp);
                            tmp = html_yahoo.IndexOf(",");
                            value = html_yahoo.Substring(3, tmp - 3);
                            data_tmp.value = Convert.ToDouble(value);

                            if (perd[times] == "d") day_data.Add(data_tmp);
                            else if (perd[times] == "w")
                            {
                                if (Is_Add_data == 1)
                                    week_data.Add(data_tmp);
                                else if (Is_Add_data == 2)   //更改最後一列資料
                                {
                                    week_data[data_length - 1].open = data_tmp.open;
                                    week_data[data_length - 1].high = data_tmp.high;
                                    week_data[data_length - 1].low = data_tmp.low;
                                    week_data[data_length - 1].value = data_tmp.value;
                                }
                            }
                        }
                    }
                    //------------------------------------------------//
                }

            }
        }

        /****
        * Get_Average
        * @Function: 從當前日線月線周線buffer資料中算出移動平均
        * @input     day: 取幾筆資料平均
        * @output    算出的平均會存放在共用buffer的average_MA1,average_MA2,average_MA3
        * Note       請盡量保持 day1 < day2 < day3
        ****/
        private void Get_Average(int day1, int day2, int day3, int day4)
        {

            //排序成day1 < day2 < day3
            List<int> tmp = new List<int> { day1, day2, day3, day4 };
            tmp.Sort(); //默認是有小到大
            day1 = tmp[0];
            day2 = tmp[1];
            day3 = tmp[2];
            day4 = tmp[3];
            /*
                        List<data> use_data = new List<data>();
                        if (stock_parameter.d_w_m == "d") use_data = day_data;
                        else if (stock_parameter.d_w_m == "w") use_data = week_data;
                        else use_data = month_data;
            */
            //-------------------------
            int day = 0;
            List<data> use_data = new List<data>();

            string[] perd = { "d", "w", "m" };


            for (int perd_times = 0; perd_times < 2; perd_times++)
            {
                if (perd[perd_times] == "d") use_data = day_data;
                else if (perd[perd_times] == "w") use_data = week_data;
                else use_data = month_data;

                for (int times = 0; times < 4; times++)     //日線做4次 4種MA
                {
                    if (times == 0) day = day1;
                    else if (times == 1) day = day2;
                    else if (times == 2) day = day3;
                    else day = day4;

                    int data_length = use_data.Count; //假設10筆資料 出來會是11
                    double total = 0;

                    if (data_length < day) continue; //防呆，總資料<MA數  ex:只有10筆資料，卻要60筆平均
                    for (int i = 0; i < day - 1; i++) //先少加一天 , 假設3天就先加2天就好
                    {
                        total += use_data[i].value;
                    }
                    for (int first_index = 0, last_index = day - 1; last_index < data_length; last_index++, first_index++) //day-1: 假設day=3 ,第一筆平均應該放在[2]資料
                    {
                        total += use_data[last_index].value;    //加一天
                        double average = (double)(total / day);
                        total -= use_data[first_index].value;   //減掉最前面的

                        if (times == 0) use_data[last_index].average_MA1 = Math.Round(average, 2);
                        else if (times == 1) use_data[last_index].average_MA2 = Math.Round(average, 2);
                        else if (times == 2) use_data[last_index].average_MA3 = Math.Round(average, 2);
                        else use_data[last_index].average_MA4 = Math.Round(average, 2);
                    }
                }
            }


            /*
                        for (int times = 0; times < 4; times++)     //週線做三次 三種MA
                        {
                            if (times == 0) day = day1;
                            else if (times == 1) day = day2;
                            else if (times == 2) day = day3;
                            else day = day4;

                            int data_length = week_data.Count; //假設10筆資料 出來會是11
                            double total = 0;

                            if (data_length < day) continue; //防呆，總資料<MA數  ex:只有10筆資料，卻要60筆平均
                            for (int i = 0; i < day - 1; i++) //先少加一天 , 假設3天就先加2天就好
                            {
                                total += week_data[i].value;
                            }
                            for (int first_index = 0, last_index = day - 1; last_index < data_length; last_index++, first_index++) //day-1: 假設day=3 ,第一筆平均應該放在[2]資料
                            {
                                total += week_data[last_index].value;    //加一天
                                double average = (double)(total / day);
                                total -= week_data[first_index].value;   //減掉最前面的

                                if (times == 0) week_data[last_index].average_MA1 = Math.Round(average, 2);
                                else if (times == 1) week_data[last_index].average_MA2 = Math.Round(average, 2);
                                else if (times == 2) week_data[last_index].average_MA3 = Math.Round(average, 2);
                                else week_data[last_index].average_MA4 = Math.Round(average, 2);
                            }
                        }
            /*
                        for (int times = 0; times < 3; times++)     //月線做三次 三種MA
                        {
                            if (times == 0) day = day1;
                            else if (times == 1) day = day2;
                            else if (times == 2) day = day3;

                            int data_length = month_data.Count; //假設10筆資料 出來會是11
                            double total = 0;

                            if (data_length < day) continue; //防呆，總資料<MA數  ex:只有10筆資料，卻要60筆平均
                            for (int i = 0; i < day - 1; i++) //先少加一天 , 假設3天就先加2天就好
                            {
                                total += month_data[i].value;
                            }
                            for (int first_index = 0, last_index = day - 1; last_index < data_length; last_index++, first_index++) //day-1: 假設day=3 ,第一筆平均應該放在[2]資料
                            {
                                total += month_data[last_index].value;    //加一天
                                double average = (double)(total / day);
                                total -= month_data[first_index].value;   //減掉最前面的

                                if (times == 0) month_data[last_index].average_MA1 = Math.Round(average, 2);
                                else if (times == 1) month_data[last_index].average_MA2 = Math.Round(average, 2);
                                else if (times == 2) month_data[last_index].average_MA3 = Math.Round(average, 2);
                            }
                        }
            */
        }



        /****
        * Get_KD  
        * @Function: 取得KD
        * @input     ID: 為了printf用而已
        * @output   
        * Note      
        ****/
        private void Get_KD(string ID, string d_w_m)
        {
            List<data> use_data = new List<data>();

            string[] perd = { "d", "w", "m" };

            int KD_day = 9;

            int MAX = 1, MIN = 0;

            perd[0] = stock_parameter.d_w_m;
            for (int times = 0; times < 1; times++)     //日線周線月線
            {
                if (perd[times] == "d") use_data = day_data;
                else if (perd[times] == "w") use_data = week_data;
                else use_data = month_data;

                int data_length = use_data.Count; if (data_length < KD_day) return;

                //--step0 :  設定第八天的KD為0
                use_data[KD_day - 2].K = 50;
                use_data[KD_day - 2].D = 50;

                //--step1 :  從第九天開始 
                for (int day_index = KD_day - 1; day_index < data_length; day_index++)
                {
                    double end = use_data[day_index].value;

                    double[] Max_Min = new double[2];
                    double max_value = 0;
                    double min_value = 0;

                    Max_Min = Find_Max_Min(day_index, 9, ref use_data);
                    max_value = Max_Min[MAX];
                    min_value = Max_Min[MIN];

                    double RSV = Math.Round((end - min_value) / (max_value - min_value) * 100, 4);
                    double K = Math.Round(RSV / 3 + use_data[day_index - 1].K * 2 / 3, 4);
                    double D = Math.Round(K / 3 + use_data[day_index - 1].D * 2 / 3, 4);

                    use_data[day_index].RSV = RSV;
                    use_data[day_index].K = K;
                    use_data[day_index].D = D;

                }

            }
        }


        /****
        * Get_MACD  (MACD)
        * @Function: 計算MACD
        * @input     ID: 為了printf用而已
        * @output   
        * Note      
        ****/
        private void Get_MACD(string ID, string d_w_m)
        {
            List<data> use_data = new List<data>();

            string[] perd = { "d", "w", "m" };

            int EMA1 = stock_parameter.EMA1;
            int EMA2 = stock_parameter.EMA2;
            int MACD = stock_parameter.MACD;


            for (int times = 0; times < 2; times++)     //日線周線月線
            {
                if (perd[times] == "d") use_data = day_data;
                else if (perd[times] == "w") use_data = week_data;
                else use_data = month_data;

                int data_length = use_data.Count; if (data_length < 30) return;

                //--step1:  Get DI  , DI = (最高+最低+2*收盤)/4
                for (int i = 0; i < data_length; i++)
                {
                    double high = use_data[i].high;
                    double low = use_data[i].low;
                    double end = use_data[i].value;

                    double DI = Math.Round((high + low + 2 * end) / 4, 3);

                    use_data[i].DI = DI;
                }

                //---step 2  Get EMA12 & EMA26
                //先取得第一筆EMA12 與第一筆EAM26
                double DI_total = 0;
                for (int i = 0; i < EMA2; i++)
                {
                    DI_total += use_data[i].DI;
                    if (i == EMA1 - 1) use_data[i].EMA12 = Math.Round(DI_total / EMA1, 3);
                    if (i == EMA2 - 1)
                    {
                        use_data[i].EMA26 = Math.Round(DI_total / EMA2, 3);
                        use_data[i].DIF = use_data[i].EMA12 - use_data[i].EMA26;  //順便把第一筆DIF算出來
                    }
                }

                //取得之後的 EMA12 EMA26
                for (int i = EMA1; i < data_length; i++)
                {
                    double EMA12 = (use_data[i].DI * 2 + use_data[i - 1].EMA12 * (EMA1 - 1)) / (EMA1 + 1);
                    EMA12 = Math.Round(EMA12, 3);

                    use_data[i].EMA12 = EMA12;

                    if (i > (EMA2 - 1))
                    {
                        double EMA26 = (use_data[i].DI * 2 + use_data[i - 1].EMA26 * (EMA2 - 1)) / (EMA2 + 1);
                        EMA26 = Math.Round(EMA26, 3);
                        use_data[i].EMA26 = EMA26;

                        double DIF = use_data[i].EMA12 - use_data[i].EMA26;
                        use_data[i].DIF = Math.Round(DIF, 3);
                    }
                    if (i == EMA2 + MACD - 2)    //第一筆MACD
                    {
                        double DIF_Total = 0.0;
                        for (int j = 0; j < MACD; j++)
                            DIF_Total += use_data[EMA2 - 1 + j].DIF;
                        use_data[i].MACD = Math.Round(DIF_Total / MACD, 3);

                        use_data[i].OSC = Math.Round(use_data[i].DIF - use_data[i].MACD, 3);
                    }
                    if (i > EMA2 + MACD - 2)
                    {
                        double MACD_tmp = Math.Round((use_data[i].DIF * 2 + use_data[i - 1].MACD * (MACD - 1)) / (MACD + 1), 3);
                        use_data[i].MACD = MACD_tmp;

                        use_data[i].OSC = Math.Round(use_data[i].DIF - use_data[i].MACD, 3);
                    }
                }
            }
        }



        /****
        * Get_CCI  (CCI)
        * @Function: 計算CCI
        * @input     ID: 為了printf用而已
        * @output   
        * Note      CCI = (TP-MA) / MD / 0.015
        ****/
        private void Get_CCI(string ID, string d_w_m)
        {
            List<data> use_data = new List<data>();

            string[] perd = { "d", "w", "m" };

            int avg_day = stock_parameter.CCI_day;

            for (int times = 0; times < 2; times++)     //日線周線
            {
                if (perd[times] == "d") use_data = day_data;
                else if (perd[times] == "w") use_data = week_data;
                else use_data = month_data;

                int data_length = use_data.Count; if (data_length < 30) return;

                double[] TP = new double[data_length];
                double[] MA = new double[data_length];

                //--step1:  Get TP  , TP = (最高+最低+收盤)/3
                for (int i = 0; i < data_length; i++)
                {
                    double high = use_data[i].high;
                    double low = use_data[i].low;
                    double end = use_data[i].value;

                    double TP_ = Math.Round((high + low + end) / 3, 3);

                    TP[i] = TP_;
                }

                //--step2 : Get MA , MA = Tp加總/avg_day  

                double total_TP = 0;
                for (int i = 0; i < avg_day - 1; i++) total_TP += TP[i];  //Tp總和先少加一筆 之後每個loop就是先加今日的TP  計算完最後扣掉第一筆
                for (int i = avg_day - 1; i < data_length; i++)
                {
                    total_TP += TP[i];

                    double MA_ = Math.Round((total_TP / avg_day), 3);
                    MA[i] = MA_;

                    total_TP -= TP[i - avg_day + 1];
                }

                //step3--        Get MD , MD = abs(MA-Tp)/avg_day
                //--        Get CCI , CCI = (TP-MA) / MD / 0.015

                for (int i = 2 * avg_day - 2; i < data_length; i++)
                {
                    double MA_minus_TP = 0;
                    for (int j = 0; j <= avg_day - 1; j++)
                    {
                        MA_minus_TP += Math.Abs(MA[i - j] - TP[i - j]);
                    }
                    double MD = Math.Round(MA_minus_TP / avg_day, 3); //StDev(tp_tmp)

                    double CCI = (TP[i] - MA[i]) / MD / 0.015;
                    use_data[i].CCI = Math.Round(CCI, 3);
                }

            }
        }




        /****
        * Get_EPS
        * @Function: Get EPS
        * @input     ID
        * @output   
        * Note      
        ****/
        private double Get_EPS(string ID)
        {
            double EPS = 0.0;
            string[] value_tmp = new string[11];

            string url = "https://stock.wearn.com/financial.asp?kind=" + ID;
            string html = GetWebContent(url);

            for (int year = 0; year < 1; year++)
            {
                int tmp = html.IndexOf("stockalllistbg"); if (tmp == -1) break;
                html = html.Substring(tmp);

                for (int i = 0; i < 11; i++)
                {
                    tmp = html.IndexOf("right");
                    html = html.Substring(tmp);
                    tmp = html.IndexOf("&nbsp");
                    value_tmp[i] = html.Substring(7, tmp - 7);

                    html = html.Substring(tmp);
                }
                for (int i = 0; i < 11; i++)        //刪除所有%
                {
                    value_tmp[i] = value_tmp[i].Replace("%", "");
                }
                //   all_data.gross_margin[year] = Convert.ToDouble(value_tmp[0]);            //營業毛利率    
                //   all_data.operating_profit_margin[year] = Convert.ToDouble(value_tmp[1]); //營業利益率
                //   all_data.operating_expense_ratio[year] = all_data.gross_margin[year] - all_data.operating_profit_margin[year];           //營業費用率
                //   all_data.profit_margin[year] = Convert.ToDouble(value_tmp[2]);            //淨利率
                //   all_data.current_ratio[year] = Convert.ToDouble(value_tmp[3]);            //流動比率
                //   all_data.quick_ratio[year] = Convert.ToDouble(value_tmp[4]);              //速動比率
                //   all_data.debts_ratio[year] = Convert.ToDouble(value_tmp[5]);              //負債佔總資產比
                //   all_data.ROE[year] = Convert.ToDouble(value_tmp[9]);                      //ROE
                EPS = Convert.ToDouble(value_tmp[10]);                     //EPS
            }

            return EPS;

        }







        /****
      * Get_EPS
      * @Function: Get EPS
      * @input     ID
      * @output   
      * Note      
      ****/
        private UInt32 Get_Quantity(string ID)
        {
            UInt32 Quantity = 0;
            string[] value_tmp = new string[11];

            string url = "https://tw.quote.finance.yahoo.net/quote/q?type=ta&perd=d&mkt=10&sym=" + ID + "&v=1";
            string html = GetWebContent(url);

            int tmp = html.LastIndexOf("v");  if (tmp == -1) return 0;
            html = html.Substring(tmp);

            tmp = html.LastIndexOf("}]"); if (tmp == -1) return 0;

            string Quantity_tmp = html.Substring(3, tmp - 3);
            Quantity = Convert.ToUInt32(Quantity_tmp);
       
            return Quantity;

        }







        /****
        * Find_Max_Min
        * @Function: 從找出use_data 以day_index往前count天的最大or最小的收盤價
        * @input     day_index : 當前是哪天的index   count :可以改要近幾天的MAX/MIN   MAX/MIN   //MAX = 1 MIN = -1
        * @output    近count最MAX/MIN的value
        * Note      
        ****/
        private double[] Find_Max_Min(int day_index, int count, ref List<data> use_data)
        {
            count = count - 1;
            int MIN = 0, MAX = 1;
            double[] return_value = new double[2];

            int last_index = use_data.Count - 1;
            if (last_index < count) return null;

            return_value[MIN] = use_data[day_index].high;
            return_value[MAX] = use_data[day_index].low;

            for (int i = 1; i < count; i++)
            {
                if (use_data[day_index - i].high > return_value[MAX])
                    return_value[MAX] = use_data[day_index - i].high;

                if (use_data[day_index - i].low < return_value[MIN])
                    return_value[MIN] = use_data[day_index - i].low;
            }
            return return_value;

        }





        /****
         * Find_Earn1_5Percent  找出以當天
         * @Function: 找出以當天往後幾天  會賺1.5%  所以要看日線 ， 但如果日線資料不夠 就用當前的線
         * @input    
         * @output    回傳天數 -1:沒找到
         * Note      
         ****/

        private string Find_Earn1_5Percent(int day_index, string d_w_m)
        {
            List<data> use_data = new List<data>();
            if (d_w_m == "d") use_data = day_data;
            else if (d_w_m == "w") use_data = week_data;
            else use_data = month_data;

            double target_value = use_data[day_index].value * earn_percent;

            for (int after = 0; after + day_index < use_data.Count; after++)
            {
                if (use_data[after + day_index].value >= target_value) return after.ToString() + d_w_m;

            }

            return "NOOOOOOO" + d_w_m;


        }






        /****
         * Find_lower_8percent  
         * @Function: 找出低於五日線8%
         * @input     ID: 為了printf用而已
         * @output   
         * Note      
         ****/

        private void Find_lower_8percent(string ID, string d_w_m)
        {
            List<data> use_data = new List<data>();
            if (d_w_m == "d") use_data = day_data;
            else if (d_w_m == "w") use_data = week_data;
            else use_data = month_data;

            int data_length = use_data.Count;
            int last_index = use_data.Count - 1;

            int index = Grid_TLB.Rows.Add();
            Grid_TLB.Rows[index].Cells[(int)Grid.ID].Value = ID;

            for (int day_index = 21; day_index < data_length; day_index++)
            {
                if (use_data[day_index].average_MA3 * 0.85 > use_data[day_index].value)
                {
                    Grid_TLB.Rows[index].Cells[(int)Grid.Date].Value = use_data[day_index].date;
                    Grid_TLB.Rows[index].Cells[(int)Grid.Value].Value = use_data[day_index].value;
                    index = Grid_TLB.Rows.Add();
                }
            }
        }




        /****
        * Find_TLB_Max_Min
        * @Function: 從現有TLB_data中找出近count的最大or最小
        * @input     count :可以改要近幾天的MAX/MIN   MAX/MIN   //MAX = 1 MIN = -1
        * @output    近count最MAX/MIN的value
        * Note      
        ****/
        private double Find_TLB_Max_Min(int count, int MAX_MIN)
        {
            int MAX = 1, MIN = -1;
            double MAX_value = 0;
            double MIN_value = 9999.9;
            int last_index = TLB_data.Count - 1;

            if (MAX_MIN == MAX) //找最高UP
            {
                for (int i = 0; i < count; i++)
                {
                    if (last_index < 0) break;
                    if (TLB_data[last_index].up > MAX_value)
                        MAX_value = TLB_data[last_index].up;
                    last_index--;
                }
                return MAX_value;
            }
            else       //找最低down
            {
                for (int i = 0; i < count; i++)
                {
                    if (last_index < 0) break;
                    if (TLB_data[last_index].down < MIN_value)
                        MIN_value = TLB_data[last_index].down;

                    last_index--;
                }
                return MIN_value;
            }
        }



        //count 可以改要幾天為基準 一般是三天
        /****
        * Get_TLB
        * @Function: 從當前日線buffer算出TLB
        * @input     count :可以改要幾天為基準 一般是三天
        * @output    算出的TLB會存放在共用buffer TLB_data
        * Note      
        ****/
        private void Get_TLB(int count, string ID, string d_w_m)  //Three line break     //這裡只有寫日線
        {
            List<data> use_data = new List<data>();
            if (d_w_m == "d") use_data = day_data;
            else if (d_w_m == "w") use_data = week_data;
            else use_data = month_data;

            TLB_data.Clear();
            int MAX = 1, MIN = -1;
            int data_length = use_data.Count;
            int status = 0; //1:陽線 0:陰線
            double[] value_buffer = new double[3];
            int TLB_start_index = 0;
            int day_data_index = 0;

            //先畫出前count 根

            //第一根

            if (use_data[1].value > use_data[0].value)  //基準價小於第2
            {
                TLB tmp = new TLB();
                tmp.date = use_data[1].date;
                tmp.up = use_data[1].value;
                tmp.down = use_data[0].value;
                tmp.status = 1;
                TLB_data.Add(tmp);
            }
            else                                       //基準>第2天
            {
                TLB tmp = new TLB();
                tmp.date = use_data[1].date;
                tmp.up = use_data[0].value;
                tmp.down = use_data[1].value;
                tmp.status = 0;
                TLB_data.Add(tmp);
            }

            //------step 1 End
            //       int TLB_last_index = TLB_data.Count-1;

            for (day_data_index = 2; day_data_index < data_length; day_data_index++)
            {
                int TLB_last_index = TLB_data.Count - 1;

                if (TLB_data[TLB_last_index].status == 1)   //現在是陽線
                {
                    if (use_data[day_data_index].value > TLB_data[TLB_last_index].up)  //現在陽線 且收盤直接大於up 最簡單 直接加一條陽線
                    {
                        TLB tmp = new TLB();
                        tmp.date = use_data[day_data_index].date;
                        tmp.up = use_data[day_data_index].value;
                        tmp.down = TLB_data[TLB_last_index].up;
                        tmp.status = 1;
                        TLB_data.Add(tmp);
                        continue;
                    }

                    //若沒有直接大於 就要看有沒有低於近count天的最低down
                    double MIN_value = Find_TLB_Max_Min(count, MIN);

                    if (use_data[day_data_index].value < MIN_value)
                    {
                        TLB tmp = new TLB();
                        tmp.date = use_data[day_data_index].date;
                        tmp.up = TLB_data[TLB_last_index].down;
                        tmp.down = use_data[day_data_index].value;
                        tmp.status = 0;
                        TLB_data.Add(tmp);
                        continue;
                    }
                }

                else    //現在是陰線
                {
                    if (use_data[day_data_index].value < TLB_data[TLB_last_index].down)  //現在陰線 且收盤直接小於down 最簡單 直接加一條陰線
                    {
                        TLB tmp = new TLB();
                        tmp.date = use_data[day_data_index].date;
                        tmp.up = TLB_data[TLB_last_index].down;
                        tmp.down = use_data[day_data_index].value;
                        tmp.status = 0;
                        TLB_data.Add(tmp);
                        continue;
                    }

                    //若沒有直接小於 就要看有沒有高於近count天的最高up
                    double MAX_value = Find_TLB_Max_Min(count, MAX);

                    if (use_data[day_data_index].value > MAX_value)
                    {
                        TLB tmp = new TLB();
                        tmp.date = use_data[day_data_index].date;
                        tmp.up = use_data[day_data_index].value;
                        tmp.down = TLB_data[TLB_last_index].up;
                        tmp.status = 1;
                        TLB_data.Add(tmp);
                        continue;
                    }
                }
            }

        }

        /****
        * Find_TLB_UP
        * @Function: 從三價線中，找出開始翻紅的日期
        * @input     ID: 為了printf用而已
        * @output   
        * Note      
        ****/
        private void Find_TLB_UP(string ID, string d_w_m)
        {
            List<data> use_data = new List<data>();
            if (d_w_m == "d") use_data = day_data;
            else if (d_w_m == "w") use_data = week_data;
            else use_data = month_data;

            int last_data = TLB_data.Count - 1;

            bool condition = TLB_data[last_data].status == 1 && TLB_data[last_data - 1].status == 0;    //由綠翻紅

            if (cbx_IsFilter.Checked == true)
            {
                if (!condition) //今日由綠翻紅 直接離開
                {
                    return;
                }
            }


            int index = Grid_TLB.Rows.Add();
            Grid_TLB.Rows[index].Cells[(int)Grid.ID].Value = ID;

            for (int i = 1; i < TLB_data.Count; i++)
            {
                condition = TLB_data[last_data].status == 1 && TLB_data[last_data - 1].status == 0;


                if (condition)
                {

                    Grid_TLB.Rows[index].Cells[(int)Grid.Date].Value = TLB_data[i].date;
                    Grid_TLB.Rows[index].Cells[(int)Grid.Value].Value = TLB_data[i].up;

                    int find_day = 0;
                    //------尋找翻紅後幾天賺>1元

                    //先找出這是哪天

                    for (; find_day < use_data.Count; find_day++)
                    {
                        if (TLB_data[i].date == use_data[find_day].date) break;
                    }

                    string after_day = Find_Earn1_5Percent(find_day, d_w_m);

                    Grid_TLB.Rows[index].Cells[(int)Grid.Find_days].Value = after_day; //

                    index = Grid_TLB.Rows.Add();
                }
            }
        }



        /****
        * Find_TLB_Down_Nday
        * @Function: 從三價線中，找出已連續陰線N天
        * @input     ID: 為了printf用而已
        * @output   
        * Note      
        ****/
        private int Find_TLB_Down_Nday(string ID)
        {
            int TLB_last_index = TLB_data.Count - 1;
            int count = 0;
            while (true)
            {
                if (TLB_data[TLB_last_index].status == 1) break;
                TLB_last_index--;
                count++;
            }
            return count;
        }



        /****
        * Find_MA2_CrossMA3
        * @Function: 找出MA2向下穿越MA3的到現在天數
        * @input     ID: 為了printf用而已 int day_index : 模擬實當天的日期
        * @output   -1: MA10還在MA21上方
        * Note       
        ****/
        private int Find_MA2_CrossMA3(string ID, int day_index, string d_w_m)
        {
            List<data> use_data = new List<data>();
            if (d_w_m == "d") use_data = day_data;
            else if (d_w_m == "w") use_data = week_data;
            else use_data = month_data;
            int count = 0;

            //先確認MA2在MA3下方
            if (use_data[day_index].average_MA3 < use_data[day_index].average_MA2) return -1;

            for (; day_index >= 21; day_index--)
            {
                if (use_data[day_index].average_MA3 > use_data[day_index].average_MA2 && use_data[day_index - 1].average_MA3 < use_data[day_index - 1].average_MA2)
                    break;
                count++;
            }
            return count;
        }





        /****
        * Get_day_MACD_status
        * @Function: Get 目前日周MACD狀態
        * @input     ID: 為了printf用而已
        * @output   
        * Note      
        ****/
        private string Get_day_MACD_status()
        {
            string[] ret = new string[2];
            //-----------日線
            int last_index = day_data.Count() - 1;
            if (day_data[last_index].OSC < 0)
                ret[0] = "日:綠";
            else
                ret[0] = "日:紅";
            if (day_data[last_index].OSC > day_data[last_index - 1].OSC)
                ret[0] += "  正    ";
            else
                ret[0] += "  負    ";


            //-------------周線

            last_index = week_data.Count() - 1;
            if (week_data[last_index].OSC < 0)
                ret[0] += "週:綠";
            else
                ret[0] += "週:紅";
            if (week_data[last_index].OSC > week_data[last_index - 1].OSC)
                ret[0] += "  正";
            else
                ret[0] += "  負";



            return ret[0];
        }







        /****
        * Find_MA1_To_Negative
        * @Function: 找出，目前的值 與上次MA1正變負的值 的差異
        * @input     ID: 為了printf用而已
        * @output   
        * Note      
        ****/
        private double Find_MA1_To_Negative(int day_index, ref List<data> use_data)
        {
            day_index--;
            double value_delta = use_data[day_index].average_MA1;
            while (use_data[day_index].average_MA1 < use_data[day_index - 1].average_MA1)
            {
                day_index--;

                if (day_index <= 4) break;
            }
            value_delta = use_data[day_index].average_MA1 - value_delta;
            return Math.Round(value_delta, 2);
        }


        /****
        * Is_Long_Sort  
        * @Function: 檢查當天是否多頭排列
        * @input     ID: 為了printf用而已
        * @output   
        * Note      
        ****/

        private bool Is_Long_Sort(int index, ref List<data> use_data)
        {
            if (use_data[index].average_MA1 > use_data[index].average_MA2 && use_data[index].average_MA2 > use_data[index].average_MA3)
                return true;

            else return false;

        }


        /****
        * Find_MA1_To_Postive
        * @Function: 找出3日線由負轉正
        * @input     ID: 為了printf用而已
        * @output   
        * Note      
        ****/
        private void Find_MA1_To_Postive(string ID, string d_w_m)
        {
            List<data> use_data = new List<data>();
            if (d_w_m == "d") use_data = day_data;
            else if (d_w_m == "w") use_data = week_data;
            else use_data = month_data;


            int last_index = use_data.Count - 1;

            bool condition = use_data[last_index].average_MA1 > use_data[last_index - 1].average_MA1 && use_data[last_index - 1].average_MA1 < use_data[last_index - 2].average_MA1; //MA1負轉正

            if (cbx_IsFilter.Checked == true)
            {
                if (!condition)  //MA1負轉正
                {
                    return;
                }
            }

            int index = Grid_TLB.Rows.Add();
            double EPS = Get_EPS(ID);
            for (int i = stock_parameter.MA1 + 1; i < use_data.Count; i++)  //3日線 所以i = 4開始
            {
                condition = use_data[i].average_MA1 > use_data[i - 1].average_MA1 && use_data[i - 1].average_MA1 < use_data[i - 2].average_MA1;
                if (condition)
                {
                    int count_tmp = Find_MA2_CrossMA3(ID, i, d_w_m);
                    Grid_TLB.Rows[index].Cells[(int)Grid.ID].Value = ID;
                    Grid_TLB.Rows[index].Cells[(int)Grid.Date].Value = use_data[i].date;
                    Grid_TLB.Rows[index].Cells[(int)Grid.Value].Value = use_data[i].value;

                    //                    Grid_TLB.Rows[index].Cells[(int)Grid.MA1_To_Negative_Value_Delta].Value = Find_MA1_To_Negative(i, ref use_data);

                    //-----------找1元

                    string after_day = Find_Earn1_5Percent(i, d_w_m);
                    Grid_TLB.Rows[index].Cells[(int)Grid.Find_days].Value = after_day;

                    string day_status = Get_day_MACD_status();
                    Grid_TLB.Rows[index].Cells[(int)Grid.day_status].Value = day_status;

                    Grid_TLB.Rows[index].Cells[(int)Grid.EPS].Value = EPS;

                    index = Grid_TLB.Rows.Add();
                }

            }
        }


        /****
        * Is_slope_Pos_Neg  
        * @Function: 判斷當天斜率為正或負
        * @input     index : 要判斷的那天  , MA : 哪個MA
        * @output    true: 斜率為正  false: 斜率為負 
        * Note      
        ****/
        private bool Is_slope_Pos(int index, int MA, ref List<data> use_data)
        {
            double today_MA = 0, yesterday_MA = 0;
            if (MA == 1)
            {
                today_MA = use_data[index].average_MA1;
                yesterday_MA = use_data[index - 1].average_MA1;
            }
            else if (MA == 2)
            {
                today_MA = use_data[index].average_MA2;
                yesterday_MA = use_data[index - 1].average_MA2;
            }
            else if (MA == 3)
            {
                today_MA = use_data[index].average_MA3;
                yesterday_MA = use_data[index - 1].average_MA3;
            }
            else if (MA == 4)
            {
                today_MA = use_data[index].average_MA4;
                yesterday_MA = use_data[index - 1].average_MA4;
            }

            if (today_MA > yesterday_MA) return true;
            else return false;
        }


        /****
        * Find_Long_Sort  (多頭排列)
        * @Function: 找出多頭排列
        * @input     ID: 為了printf用而已
        * @output   
        * Note      
        ****/
        private void Find_Long_Sort(string ID, string d_w_m)
        {
            List<data> use_data = new List<data>();
            if (d_w_m == "d") use_data = day_data;
            else if (d_w_m == "w") use_data = week_data;
            else use_data = month_data;

            int data_length = use_data.Count;
            int last_index = use_data.Count - 1;

            bool condition1 = Is_slope_Pos(last_index, 1, ref use_data) && Is_slope_Pos(last_index, 2, ref use_data) /*&& Is_slope_Pos(last_index, 3, ref use_data)*/; //三線均正                                            
            bool condition2 = Is_Long_Sort(last_index, ref use_data);   //今日是否多頭

            bool today_condition = (condition1 == true && condition2 == true);

            condition1 = Is_slope_Pos(last_index - 1, 1, ref use_data) && Is_slope_Pos(last_index - 1, 2, ref use_data) /*&& Is_slope_Pos(last_index-1, 3, ref use_data)*/; //昨日是否三線均正                                            
            condition2 = Is_Long_Sort(last_index - 1, ref use_data);   //昨日是否多頭


            bool yesterday_condition = (condition1 == true && condition2 == true);

            bool totally_condition = today_condition == true && yesterday_condition == false;              //期望今天是三線均正或多頭開始的第一天
            if (cbx_IsFilter.Checked == true)
            {
                if (totally_condition == false)  ////第一個 三線均正 且 多頭
                {
                    return;
                }
            }

            int index = Grid_TLB.Rows.Add();
            double EPS = Get_EPS(ID);
            //----模擬過去第一天多頭排列時

            for (int day_index = 21; day_index < data_length; day_index++)
            {
                condition1 = Is_slope_Pos(day_index, 1, ref use_data) && Is_slope_Pos(day_index, 2, ref use_data) && Is_slope_Pos(day_index, 3, ref use_data); //三線均正                                            
                condition2 = Is_Long_Sort(day_index, ref use_data);   //今日是否多頭

                today_condition = (condition1 == true && condition2 == true);

                condition1 = Is_slope_Pos(day_index - 1, 1, ref use_data) && Is_slope_Pos(day_index - 1, 2, ref use_data) && Is_slope_Pos(day_index - 1, 3, ref use_data); //昨日是否三線均正                                            
                condition2 = Is_Long_Sort(day_index - 1, ref use_data);   //昨日是否多頭


                yesterday_condition = (condition1 == true && condition2 == true);

                totally_condition = today_condition == true && yesterday_condition == false;              //期望今天是三線均正或多頭開始的第一天

                if (totally_condition == false) continue;
                //是第一天多頭排列~~
                Grid_TLB.Rows[index].Cells[(int)Grid.ID].Value = ID;
                Grid_TLB.Rows[index].Cells[(int)Grid.Date].Value = use_data[day_index].date;
                Grid_TLB.Rows[index].Cells[(int)Grid.Value].Value = use_data[day_index].value;

                double tmp = use_data[day_index].value;

                string after_day_find = Find_Earn1_5Percent(day_index, d_w_m);
                Grid_TLB.Rows[index].Cells[(int)Grid.Find_days].Value = after_day_find;

                string day_status = Get_day_MACD_status();
                Grid_TLB.Rows[index].Cells[(int)Grid.day_status].Value = day_status;

                Grid_TLB.Rows[index].Cells[(int)Grid.EPS].Value = EPS;
                index = Grid_TLB.Rows.Add();
            }
        }






        /* Find_Green_MACD_Count
        * @Function: 找出，目前位置之前有幾根綠住
        * @input     ID: 為了printf用而已
        * @output   
        * Note      
        ****/
        private int Find_Green_MACD_Count(int day_index, ref List<data> use_data)
        {

            int count = 0;
            for (day_index--; day_index > 0; day_index--)
            {
                if (use_data[day_index].OSC >= 0) break;
                count++;
            }
            return count;
        }




        /****
        * Find_MACD_To_Pos  
        * @Function: 找出MACD由負轉正
        * @input     ID: 為了printf用而已
        * @output   
        * Note      
        ****/
        private void Find_MACD_To_Pos(string ID, string d_w_m)
        {
            List<data> use_data = new List<data>();
            if (d_w_m == "d") use_data = day_data;
            else if (d_w_m == "w") use_data = week_data;
            else use_data = month_data;

            int EMA1 = stock_parameter.EMA1;
            int EMA2 = stock_parameter.EMA2;
            int MACD = stock_parameter.MACD;

            int data_length = use_data.Count;
            int last_index = use_data.Count - 1;

            int day_index = 0;

            //TODO 簡單過濾
            //           if (use_data[last_index].OSC >= 0 || (use_data[last_index].OSC < use_data[last_index-1].OSC && use_data[last_index-1].OSC > use_data[last_index-2].OSC   )) return;
            Grid_TLB.Rows.Add();

            //先找出紅
            for (day_index = EMA2 + MACD - 2; day_index < data_length; day_index++)
            {
                if (use_data[day_index - 1].MACD < 0 && use_data[day_index].MACD >= 0)
                {
                    int index = Grid_TLB.Rows.Add();
                    Grid_TLB.Rows[index].Cells[(int)Grid.ID].Value = ID;
                    Grid_TLB.Rows[index].Cells[(int)Grid.Date].Value = use_data[day_index].date;
                    Grid_TLB.Rows[index].Cells[(int)Grid.Value].Value = use_data[day_index].value;

                    string after_day_find = Find_Earn1_5Percent(day_index, d_w_m);
                    Grid_TLB.Rows[index].Cells[(int)Grid.Find_days].Value = after_day_find;

                }
            }

        }






        /****
        * Find_MACD_Green  
        * @Function: 找出MACD 為
        * @input     ID: 為了printf用而已
        * @output   
        * Note      
        ****/
        private void Find_MACD_Green(string ID, string d_w_m)
        {
            List<data> use_data = new List<data>();
            if (d_w_m == "d") use_data = day_data;
            else if (d_w_m == "w") use_data = week_data;
            else use_data = month_data;

            int data_length = use_data.Count;
            int last_index = use_data.Count - 1;

            int day_index = 0;

            //TODO 簡單過濾
            bool condition1 = use_data[last_index].OSC > use_data[last_index - 1].OSC /*&& use_data[last_index-1].OSC < use_data[last_index-2].OSC*/;
            bool condition2 = use_data[last_index].OSC < 0;

            if (cbx_MACD_Onlyone.Checked == true) condition1 = use_data[last_index].OSC > use_data[last_index - 1].OSC && use_data[last_index - 1].OSC < use_data[last_index - 2].OSC;

            if (cbx_IsFilter.Checked == true)
            {
                if (!condition1 || !condition2)  ////綠回頭  (只要比昨日低 就算)
                {
                    return;
                }
            }

            Grid_TLB.Rows.Add();

            //先找出紅
            for (; day_index < data_length; day_index++)
            {
                if (use_data[day_index].OSC > 0) break;
            }

            //----------------------------------------------------//


            double EPS = Get_EPS(ID);

            for (; day_index < data_length; day_index++)
            {
                condition1 = use_data[day_index].OSC > use_data[day_index - 1].OSC && use_data[day_index - 1].OSC < use_data[day_index - 2].OSC;
                condition2 = use_data[day_index].OSC < 0;

                if (condition1 && condition2)
                {
                    int count_tmpA1 = Find_MA2_CrossMA3(ID, day_index, d_w_m); //MA10 向下穿越MA21
                    int count_tmp = Find_Green_MACD_Count(day_index, ref use_data);

                    int index = Grid_TLB.Rows.Add();

                    Grid_TLB.Rows[index].Cells[(int)Grid.ID].Value = ID;
                    Grid_TLB.Rows[index].Cells[(int)Grid.Date].Value = use_data[day_index].date;
                    Grid_TLB.Rows[index].Cells[(int)Grid.Value].Value = use_data[day_index].value;

                    string after_day_find = Find_Earn1_5Percent(day_index, d_w_m);
                    Grid_TLB.Rows[index].Cells[(int)Grid.Find_days].Value = after_day_find;

                    string day_status = Get_day_MACD_status();
                    Grid_TLB.Rows[index].Cells[(int)Grid.day_status].Value = day_status;
                    /*
                                        count_tmp = Find_MA2_CrossMA3(ID, day_index, d_w_m); //MA10 向下穿越MA21
                                        Grid_TLB.Rows[index].Cells[(int)Grid.MA10_Cross_MA21_days].Value = count_tmp;

                                        count_tmp = Find_Green_MACD_Count(day_index, ref use_data);
                                        Grid_TLB.Rows[index].Cells[(int)Grid.MA1_To_Negative_Value_Delta].Value = count_tmp;
                    */
                    Grid_TLB.Rows[index].Cells[(int)Grid.EPS].Value = EPS;
                }
                continue;

            }
        }



        /****
        * Find_MACD_Red  
        * @Function: 找出MACD 第一根轉紅(或下降又上去) 第二天買?
        * @input     ID: 為了printf用而已
        * @output   
        * Note      
        ****/
        private void Find_MACD_Red(string ID, string d_w_m)
        {
            List<data> use_data = new List<data>();
            if (d_w_m == "d") use_data = day_data;
            else if (d_w_m == "w") use_data = week_data;
            else use_data = month_data;

            int data_length = use_data.Count;
            int last_index = use_data.Count - 1;

            int day_index = 0;

            bool condition1 = use_data[last_index - 1].OSC <= 0 && use_data[last_index].OSC >= 0; //綠變紅
            bool condition2 = (use_data[last_index - 2].OSC > use_data[last_index - 1].OSC && use_data[last_index - 1].OSC < use_data[last_index].OSC ) && use_data[last_index].OSC >=0; //紅下降再次上升

            if (cbx_IsFilter.Checked == true)
            {
                if (!condition1 || !condition2)  ////第一根紅回頭 或 紅再次上升
                {
                    return;
                }
            }

            Grid_TLB.Rows.Add();

            //----------------------------------------------------//
            int EMA2 = stock_parameter.EMA2;
            int MACD = stock_parameter.MACD;
            double EPS = Get_EPS(ID);
            for (day_index = EMA2 + MACD - 2; day_index < data_length; day_index++)
            {
                if (use_data[day_index].OSC < 0) continue; //當前綠色 直接離開

                condition1 = use_data[day_index - 1].OSC <= 0 && use_data[day_index].OSC >= 0; //綠變紅
                condition2 = use_data[day_index - 2].OSC > use_data[day_index - 1].OSC && use_data[day_index - 1].OSC < use_data[day_index].OSC; //紅下降再次上升

                if (condition1 || condition2)
                {
                    int index = Grid_TLB.Rows.Add();
                    Grid_TLB.Rows[index].Cells[(int)Grid.ID].Value = ID;
                    Grid_TLB.Rows[index].Cells[(int)Grid.Date].Value = use_data[day_index].date;
                    Grid_TLB.Rows[index].Cells[(int)Grid.Value].Value = use_data[day_index].value;
                    string after_day_find = Find_Earn1_5Percent(day_index, d_w_m);
                    Grid_TLB.Rows[index].Cells[(int)Grid.Find_days].Value = after_day_find;
                    Grid_TLB.Rows[index].Cells[(int)Grid.EPS].Value = EPS;

                    string day_status = Get_day_MACD_status();
                    Grid_TLB.Rows[index].Cells[(int)Grid.day_status].Value = day_status;
                    /*
                    int count_tmp = Find_MA2_CrossMA3(ID, day_index, d_w_m); //MA10 向下穿越MA21
                    Grid_TLB.Rows[index].Cells[(int)Grid.day_status].Value = count_tmp;
                    */

                }
            }
        }



        /****
        * Find_CCI  (CCI)
        * @Function: 找出多CCI < -130
        * @input     ID: 為了printf用而已
        * @output   
        * Note      
        ****/
        private void Find_CCI(string ID, string d_w_m)
        {
            List<data> use_data = new List<data>();
            if (d_w_m == "d") use_data = day_data;
            else if (d_w_m == "w") use_data = week_data;
            else use_data = month_data;

            int CCI_condition = stock_parameter.CCI_condition;

            int data_length = use_data.Count;
            int last_index = use_data.Count - 1;

            bool condition1 = use_data[last_index].CCI > use_data[last_index - 1].CCI && use_data[last_index - 1].CCI < use_data[last_index - 2].CCI;   //CCI轉折                                       
            bool condition2 = use_data[last_index - 1].CCI < CCI_condition;

            bool total_condition = condition1 && condition2;

            if (cbx_IsFilter.Checked == true)
            {
                if (total_condition == false)
                {
                    return;
                }
            }

            int index = Grid_TLB.Rows.Add();
            double EPS = Get_EPS(ID);

            //----------

            for (int day_index = stock_parameter.CCI_day * 2; day_index < data_length; day_index++)
            {
                condition1 = use_data[day_index].CCI > use_data[day_index - 1].CCI && use_data[day_index - 1].CCI < use_data[day_index - 2].CCI;   //CCI轉折                                       
                condition2 = use_data[day_index - 1].CCI < CCI_condition;

                total_condition = condition1 && condition2;

                if (total_condition == false) continue;

                Grid_TLB.Rows[index].Cells[(int)Grid.ID].Value = ID;
                Grid_TLB.Rows[index].Cells[(int)Grid.Date].Value = use_data[day_index].date;
                Grid_TLB.Rows[index].Cells[(int)Grid.Value].Value = use_data[day_index].value;

                string after_day_find = Find_Earn1_5Percent(day_index, d_w_m);
                Grid_TLB.Rows[index].Cells[(int)Grid.Find_days].Value = after_day_find;

                string day_status = Get_day_MACD_status();
                Grid_TLB.Rows[index].Cells[(int)Grid.day_status].Value = day_status;

                bool MA15_status = Is_slope_Pos(day_index, 4, ref use_data);
                Grid_TLB.Rows[index].Cells[(int)Grid.MA15_status].Value = MA15_status;

                double CCI_last = use_data[day_index - 1].CCI;
                Grid_TLB.Rows[index].Cells[(int)Grid.CCI_last].Value = CCI_last;

                Grid_TLB.Rows[index].Cells[(int)Grid.EPS].Value = EPS;
                index = Grid_TLB.Rows.Add();
            }
        }


        private void NoBackTest(string ID, string d_w_m)
        {
            List<data> use_data = new List<data>();

            int index = Grid_NoBack.Rows.Add();
            bool IsNeedToPrintf = false;
            string eligible = "";
            

            //-------------------------------------------------日線多頭條件   
            for (int times = 0; times < 2; times++)
            {
                if (times == 0) use_data = day_data;
                else if (times == 1) use_data = week_data;

                eligible = "";
                IsNeedToPrintf = false;
                int last_index = use_data.Count - 1;


                //-----MACD紅
                bool condition1 = use_data[last_index - 1].OSC <= 0 && use_data[last_index].OSC >= 0; //綠變紅
                bool condition2 = (use_data[last_index - 2].OSC > use_data[last_index - 1].OSC && use_data[last_index - 1].OSC < use_data[last_index].OSC) && use_data[last_index].OSC>=0; //紅下降再次上升

                if (condition1 || condition2)
                {
                    eligible = "MACD紅 , ";
                    IsNeedToPrintf = true;
                }


                //-----MACD綠
                condition1 = use_data[last_index].OSC > use_data[last_index - 1].OSC /*&& use_data[last_index-1].OSC < use_data[last_index-2].OSC*/;
                condition2 = use_data[last_index].OSC < 0;
                if (cbx_MACD_Onlyone.Checked == true) condition1 = use_data[last_index].OSC > use_data[last_index - 1].OSC && use_data[last_index - 1].OSC < use_data[last_index - 2].OSC;
                if (condition1 && condition2)
                {
                    eligible += "MACD綠 , ";
                    IsNeedToPrintf = true;
                }

                //------多頭排列
                condition1 = Is_slope_Pos(last_index, 1, ref use_data) && Is_slope_Pos(last_index, 2, ref use_data) /*&& Is_slope_Pos(last_index, 3, ref use_data)*/; //三線均正                                            
                condition2 = Is_Long_Sort(last_index, ref use_data);   //今日是否多頭
                bool today_condition = (condition1 == true && condition2 == true);
                condition1 = Is_slope_Pos(last_index - 1, 1, ref use_data) && Is_slope_Pos(last_index - 1, 2, ref use_data) /*&& Is_slope_Pos(last_index-1, 3, ref use_data)*/; //昨日是否三線均正                                            
                condition2 = Is_Long_Sort(last_index - 1, ref use_data);   //昨日是否多頭
                bool yesterday_condition = (condition1 == true && condition2 == true);
                bool totally_condition = today_condition == true && yesterday_condition == false;              //期望今天是三線均正或多頭開始的第一天
                if (totally_condition == true)
                {
                    eligible += "多頭排列 , ";
                    IsNeedToPrintf = true;
                }

                //------MA3轉正
                condition1 = use_data[last_index].average_MA3 > use_data[last_index - 1].average_MA3 && use_data[last_index - 1].average_MA3 < use_data[last_index - 2].average_MA3; //MA1負轉正
                if (condition1 == true)
                {
                    eligible += "MA3轉正 , ";
                    IsNeedToPrintf = true;
                }

                if (IsNeedToPrintf == true)
                {
                    
                    Grid_NoBack.Rows[index].Cells[(int)Grid_NoBackTest.ID].Value = ID;
                    Grid_NoBack.Rows[index].Cells[(int)Grid_NoBackTest.Value].Value = use_data[last_index].value;
                    if (times == 0) Grid_NoBack.Rows[index].Cells[(int)Grid_NoBackTest.day_condition].Value = eligible;
                    else if (times == 1) Grid_NoBack.Rows[index].Cells[(int)Grid_NoBackTest.week_condition].Value = eligible;
                    string day_status = Get_day_MACD_status();
                    Grid_NoBack.Rows[index].Cells[(int)Grid_NoBackTest.Today_status].Value = day_status;

                    double EPS = Get_EPS(ID);
                    Grid_NoBack.Rows[index].Cells[(int)Grid_NoBackTest.EPS].Value = EPS;

                    UInt32 Quantity = Get_Quantity(ID);
                    Grid_NoBack.Rows[index].Cells[(int)Grid_NoBackTest.Quantity].Value = Quantity;
                    
                }
            }

        }

        //---------------------------------------------------

        private void button1_Click(object sender, EventArgs e)
        {

            Initial();

            var way = cbx_simulation_way.SelectedIndex;
            Simulation_Way[] simulation_way = { Find_MA1_To_Postive, Find_TLB_UP, Find_Long_Sort, Find_lower_8percent, Find_MACD_Green, Find_MACD_Red, Find_MACD_To_Pos , Find_CCI , NoBackTest };

            //------------------------前置 準備data

            Get_All_ID();

            progressBar1.Maximum = all_stock.Count;//設置最大長度值
            progressBar1.Value = 0;//設置當前值
            progressBar1.Step = 1;//設置沒次增長多少

            for (int i = 0; i < all_stock.Count; i++)
            {

                progressBar1.Value += progressBar1.Step;//讓進度條增加一次
//                Thread.Sleep(10);

                //----獲取data前 先去看量
                UInt32 Quantity = Get_Quantity(all_stock[i].ID);
                if (Quantity < stock_parameter.Quantity) continue;

                Get_All_data(all_stock[i].ID);
                Application.DoEvents();

                List<data> use_data = new List<data>();
                if (stock_parameter.d_w_m == "d") use_data = day_data;
                else if (stock_parameter.d_w_m == "w") use_data = week_data;
                else                             use_data = month_data;

                
                //加速 日線不滿100筆者  直接忽略
                int len = use_data.Count - 1;
                if (use_data.Count < 100 || use_data[len].value < 10 || use_data[len].value > 90) continue;
                Get_Average(stock_parameter.MA1, stock_parameter.MA2, stock_parameter.MA3 ,stock_parameter.MA4);
 //               Get_TLB(2, all_stock[i].ID, stock_parameter.d_w_m);
                Get_MACD(all_stock[i].ID, stock_parameter.d_w_m);
 //               Get_KD(all_stock[i].ID, stock_parameter.d_w_m);
                Get_CCI(all_stock[i].ID, stock_parameter.d_w_m);
                //----------------------------------------
                //跑模擬 選擇方式

                simulation_way[way](all_stock[i].ID, stock_parameter.d_w_m);

                
            }
        }



        private void button2_Click(object sender, EventArgs e)
        {
            chart1.Series.Clear();  //每次使用此function前先清除圖表
                                    /*
                                                ChartArea chartArea = new ChartArea();
                                                chartArea.Name = "FirstArea";

                                                chartArea.CursorX.IsUserEnabled = true;
                                                chartArea.CursorX.IsUserSelectionEnabled = true;
                                                chartArea.CursorX.SelectionColor = Color.SkyBlue;
                                                chartArea.CursorY.IsUserEnabled = true;
                                                chartArea.CursorY.AutoScroll = true;
                                                chartArea.CursorY.IsUserSelectionEnabled = true;
                                                chartArea.CursorY.SelectionColor = Color.SkyBlue;

                                                chartArea.CursorX.IntervalType = DateTimeIntervalType.Auto;
                                                chartArea.AxisX.ScaleView.Zoomable = false;
                                                chartArea.AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.All;//啓用X軸滾動條按鈕

                                                chartArea.BackColor = Color.AliceBlue;                      //背景色
                                                chartArea.BackSecondaryColor = Color.White;                 //漸變背景色
                                                chartArea.BackGradientStyle = GradientStyle.TopBottom;      //漸變方式
                                                chartArea.BackHatchStyle = ChartHatchStyle.None;            //背景陰影
                                                chartArea.BorderDashStyle = ChartDashStyle.NotSet;          //邊框線樣式
                                                chartArea.BorderWidth = 1;                                  //邊框寬度
                                                chartArea.BorderColor = Color.Black;


                                                chartArea.AxisX.MajorGrid.Enabled = true;
                                                chartArea.AxisY.MajorGrid.Enabled = true;

                                                // Axis
                                                chartArea.AxisY.Title = @"Value";
                                                chartArea.AxisY.LabelAutoFitMinFontSize = 5;
                                                chartArea.AxisY.LineWidth = 2;
                                                chartArea.AxisY.LineColor = Color.Black;
                                                chartArea.AxisY.Enabled = AxisEnabled.True;

                                                chartArea.AxisX.Title = @"Time";
                                                chartArea.AxisX.IsLabelAutoFit = true;
                                                chartArea.AxisX.LabelAutoFitMinFontSize = 5;
                                                chartArea.AxisX.LabelStyle.Angle = -15;


                                                chartArea.AxisX.LabelStyle.IsEndLabelVisible = true;        //show the last label
                                                chartArea.AxisX.Interval = 10;
                                                chartArea.AxisX.IntervalAutoMode = IntervalAutoMode.FixedCount;
                                                chartArea.AxisX.IntervalType = DateTimeIntervalType.NotSet;
                                                chartArea.AxisX.TextOrientation = TextOrientation.Auto;
                                                chartArea.AxisX.LineWidth = 2;
                                                chartArea.AxisX.LineColor = Color.Black;
                                                chartArea.AxisX.Enabled = AxisEnabled.True;
                                                chartArea.AxisX.ScaleView.MinSizeType = DateTimeIntervalType.Months;
                                                chartArea.AxisX.Crossing = 0;

                                                chartArea.Position.Height = 85;
                                                chartArea.Position.Width = 85;
                                                chartArea.Position.X = 5;
                                                chartArea.Position.Y = 7;

                                                chart.ChartAreas.Add(chartArea);
                                                chart.BackGradientStyle = GradientStyle.TopBottom;
                                                //圖表的邊框顏色、
                                                chart.BorderlineColor = Color.FromArgb(26, 59, 105);
                                                //圖表的邊框線條樣式
                                                chart.BorderlineDashStyle = ChartDashStyle.Solid;
                                                //圖表邊框線條的寬度
                                                chart.BorderlineWidth = 2;
                                                //圖表邊框的皮膚
                                                chart.BorderSkin.SkinStyle = BorderSkinStyle.Emboss;

                                                //Series1
                                                series1 = new Series();
                                                series1.ChartArea = "FirstArea";
                                                chart.Series.Add(series1);

                                                //Series1 style
                                                series1.ToolTip = "#VALX,#VALY";    //鼠標停留在數據點上，顯示XY值

                                                series1.Name = "series1";
                                                series1.ChartType = SeriesChartType.Spline;  // type
                                                series1.BorderWidth = 2;
                                                series1.Color = Color.Red;
                                                series1.XValueType = ChartValueType.Time;//x axis type
                                                series1.YValueType = ChartValueType.Int32;//y axis type

                                                //Marker
                                                series1.MarkerStyle = MarkerStyle.Square;
                                                series1.MarkerSize = 5;
                                                series1.MarkerColor = Color.Black;
                                     */

            //           ChartArea area = new ChartArea("AREA");
            //           chart1.ChartAreas.Add(area);


            //-------series_up 畫陽線
            Series series_up = new Series("Series_Up");
            series_up.BorderWidth = 1;
            series_up.ChartType = SeriesChartType.Candlestick;
            series_up.Color = Color.Red;

            this.chart1.Series.Add(series_up);//將線畫在圖上

            //------------TODO 要先找出TLB內最高跟最低
            /*
            double max = 0, min = 9999;
            for (int i = 1; i < TLB_data.Count; i++)
            {
                if (TLB_data[i].value > max)
                    max = TLB_data[i].value;
                if(TLB_data[i].value < min)
                    min = TLB_data[i].value;
            }
            chart1.ChartAreas[0].AxisY.Minimum = min-1;
            chart1.ChartAreas[0].AxisY.Maximum = max+1;
            chart1.ChartAreas[0].AxisX.Interval = 1;

            for (int i = 1; i < TLB_data.Count; i++)
            {
                chart1.Series["Series_Up"].Points.AddXY(TLB_data[i].date, TLB_data[i - 1].value, TLB_data[i - 1].value, TLB_data[i - 1].value, TLB_data[i].value);
            }
            */

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Get_Quantity("2303");
            string ID = txt_ID.Text;

            Initial();

            var way = cbx_simulation_way.SelectedIndex;
            Simulation_Way[] simulation_way = { Find_MA1_To_Postive, Find_TLB_UP, Find_Long_Sort, Find_lower_8percent, Find_MACD_Green, Find_MACD_Red, Find_MACD_To_Pos, Find_CCI, NoBackTest };

            Get_All_data(ID);
            //加速 日線不滿100筆者  直接忽略
            if (day_data.Count > 100 || week_data.Count > 50)
            {
                Get_Average(stock_parameter.MA1, stock_parameter.MA2, stock_parameter.MA3, stock_parameter.MA4);
                //    Get_TLB(3, ID, stock_parameter.d_w_m);
                Get_MACD(ID, stock_parameter.d_w_m);
            //    Get_CCI(ID, stock_parameter.d_w_m);
                //        Find_MACD_Green(ID, d_w_m);
                //               Find_Long_Sort(ID, d_w_m);
                //       NoBackTest(ID, d_w_m);
                simulation_way[way](ID, stock_parameter.d_w_m);
            }

        }

        private void checkBox33_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox33.Checked == true)
            {
                checkBox1.Checked = true; checkBox2.Checked = true; checkBox3.Checked = true; checkBox4.Checked = true; checkBox5.Checked = true;
                checkBox6.Checked = true; checkBox7.Checked = true; checkBox8.Checked = true; checkBox9.Checked = true; checkBox10.Checked = true;
            }
            else
            {
                checkBox1.Checked = false; checkBox2.Checked = false; checkBox3.Checked = false; checkBox4.Checked = false; checkBox5.Checked = false;
                checkBox6.Checked = false; checkBox7.Checked = false; checkBox8.Checked = false; checkBox9.Checked = false; checkBox10.Checked = false;
            }
            
        }
        private void checkBox34_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox34.Checked == true)
            {
                checkBox11.Checked = true; checkBox12.Checked = true; checkBox13.Checked = true; checkBox14.Checked = true; checkBox15.Checked = true;
                checkBox16.Checked = true; checkBox17.Checked = true; checkBox18.Checked = true; checkBox19.Checked = true; checkBox20.Checked = true;
            }
            else
            {
                checkBox11.Checked = false; checkBox12.Checked = false; checkBox13.Checked = false; checkBox14.Checked = false; checkBox15.Checked = false;
                checkBox16.Checked = false; checkBox17.Checked = false; checkBox18.Checked = false; checkBox19.Checked = false; checkBox20.Checked = false;
            }
            
        }
        private void checkBox35_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox35.Checked == true)
            {
                checkBox21.Checked = true; checkBox22.Checked = true; checkBox23.Checked = true; checkBox24.Checked = true; checkBox25.Checked = true;
                checkBox26.Checked = true; checkBox27.Checked = true; checkBox28.Checked = true; checkBox29.Checked = true; checkBox30.Checked = true;
                checkBox31.Checked = true; checkBox32.Checked = true;
            }
            else
            {
                checkBox21.Checked = false; checkBox22.Checked = false; checkBox23.Checked = false; checkBox24.Checked = false; checkBox25.Checked = false;
                checkBox26.Checked = false; checkBox27.Checked = false; checkBox28.Checked = false; checkBox29.Checked = false; checkBox30.Checked = false;
                checkBox31.Checked = false; checkBox32.Checked = false;
            }
            
        }






        /****
        * Get_All_Data_TW  
        * @Function: 解析證交所資料
        * @input     
        * @output   
        * Note      
        ****/
        private void Get_All_Data_TW(string ID_RowData)
        {
            int index = ID_RowData.IndexOf(":[[");
            ID_RowData = ID_RowData.Substring(index + 2);
            while (true)
            {
                data tmp = new data();
                int comma = ID_RowData.IndexOf(","); if (comma == -1) break;
                string date = ID_RowData.Substring(2, comma - 3);
                ID_RowData = ID_RowData.Substring(comma);
                int index_left = ID_RowData.IndexOf("]");
                string value = ID_RowData.Substring(2, index_left - 3);
                index = ID_RowData.IndexOf(",["); if (index == -1) break;
                ID_RowData = ID_RowData.Substring(index + 1);
                tmp.date = date;

                if (double.TryParse(value, out tmp.value) == false)
                {
                    int last_index = day_data.Count - 1;
                    if (last_index == -1) continue;

                    tmp.value = day_data[last_index].value;
                }
                day_data.Add(tmp);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string[] month = { "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12" };

            DateTime CurrTime = DateTime.Now;
            string year = CurrTime.Year.ToString();
            string week = CurrTime.Date.Month.ToString();
            string day = CurrTime.Date.Day.ToString();

            int start_year = 0, now_year = Convert.ToInt16(year);
            int need_year_count = 3;

            Get_All_ID();

            for (int id_index = 0; id_index < all_stock.Count; id_index++)
            {
                start_year = now_year - need_year_count;
                for (; start_year <= now_year; start_year++)
                {
                    for (int month_index = 0; month_index < 12; month_index++)
                    {
                        string date = start_year.ToString() + month[month_index] + "01";
                        string ID_RowData = GetWebContent("https://www.twse.com.tw/exchangeReport/STOCK_DAY_AVG?response=json&date=" + date + "&stockNo=" + all_stock[id_index].ID + "&_=1585623394656");
                        Thread.Sleep(2800); //Delay 3秒
                        Get_All_Data_TW(ID_RowData);
                    }
                }

                FileStream fileStream = new FileStream(@"d:\stock_data\" + all_stock[id_index].ID + ".txt", FileMode.Create);
                fileStream.Close();   //切記開了要關,不然會被佔用而無法修改喔!!!

                using (StreamWriter sw = new StreamWriter(@"d:\stock_data\" + all_stock[id_index].ID + ".txt"))
                {
                    // 欲寫入的文字資料 ~

                    sw.Write("ID:" + all_stock[id_index].ID + "\r\n");
                    for (int i = 0; i < day_data.Count; i++)
                    {
                        sw.Write(day_data[i].date + " " + day_data[i].value + "\r\n");
                    }
                    day_data.Clear();
                }
            }

        }

        private void checkBox36_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox36.Checked == true)
            {
                checkBox1.Enabled = false; checkBox1.Checked = false;
               checkBox10.Enabled = false; checkBox10.Checked = false;
                checkBox13.Enabled = false; checkBox13.Checked = false;
            }
            else
            {
                checkBox1.Enabled = true;
                checkBox10.Enabled = true;
                checkBox13.Enabled = true;
            }
        }
    }
}
