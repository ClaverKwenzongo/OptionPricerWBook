using OfficeOpenXml;
using System.Globalization;
using System;
using System.Diagnostics;

namespace OptionPricerWBook
{
    public class getSharePrice
    {
        static DateTime getDate(string _date)
        {
            CultureInfo cultureInfo = new CultureInfo("en-US");
            DateTime wsDate = DateTime.Parse(_date, cultureInfo);

            return wsDate;
        }


        public double _getSharePrice(string date, string shareName)
        {
            double share_price = 0;

            int date_row = 0;
            int row = 3;
            while (string.IsNullOrWhiteSpace(Globals.Sheet1.Cells[row, 1].Value?.ToString()) == false)
            {
                DateTime ws_date = getDate(Globals.Sheet1.Cells[row, 1].Value.ToString());
                Debug.WriteLine(ws_date);
                if (ws_date.ToString("dd/MM/yyyy") == date)
                {
                    date_row = row;
                    break;
                }
                row++;
            }


            int col = 2;
            int share_col = 0;
            while (string.IsNullOrWhiteSpace(Globals.Sheet1.Cells[2, col].Value?.ToString()) == false)
            {
                if (Globals.Sheet1.Cells[2, col].Value == shareName)
                {
                    share_col = col;
                    break;
                }

                col++;
            }

            share_price = Globals.Sheet1.Cells[date_row, share_col].Value;

            return share_price;
        }
    }
}
