using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

namespace OptionPricerWBook
{
    //This class finds the interest rate needed to be used in the black scholes merton to price the option.
    //There are two while loops: the first one finds the row index corresponding to the date that matches the exercise date entered by the user.
    //The second while loop will find the column index that corresponds to the tenor calculated between the exercise date and the maturity date.
    internal class getRates
    {
        static DateTime getDate(string _date)
        {
            CultureInfo cultureInfo = new CultureInfo("en-US");
            DateTime wsDate = DateTime.Parse(_date, cultureInfo);

            return wsDate;
        }

        public double getRate(double _op_tenor, string start_date)
        {

            //This loop will get the row index that corresponds to the exercise date entered by the user.
            int row = 3;
            int _date_row = 0;
            while (string.IsNullOrWhiteSpace(Globals.Sheet6.Cells[row, 1].Value?.ToString()) == false)
            {
                DateTime ws_date = getDate(Globals.Sheet6.Cells[row, 1].Value.ToString());
                if (ws_date.ToString("dd/MM/yyyy") == start_date)
                {
                    _date_row = row;
                    break;
                }
                row++;
            }

            double percent = 0;
            int col2 = 2;

            //This second loop finds the column index that matches the tenor as calculated between the exercise date and the maturity date entered by the user.
            while (string.IsNullOrWhiteSpace(Globals.Sheet6.Cells[_date_row, col2].Value?.ToString()) == false)
            {
                if (string.IsNullOrWhiteSpace(Globals.Sheet6.Cells[2, col2 + 1].Value?.ToString()) == false)
                {
                    if (_op_tenor == double.Parse(Globals.Sheet6.Cells[2, col2].Value.ToString()))
                    {
                        percent = double.Parse(Globals.Sheet6.Cells[_date_row, col2].Value.ToString());
                        break;
                    }
                    else if (_op_tenor == double.Parse(Globals.Sheet6.Cells[2, col2 + 1].Value.ToString()))
                    {
                        percent = double.Parse(Globals.Sheet6.Cells[_date_row, col2 + 1].Value.ToString());
                        break;
                    }
                    else if (double.Parse(Globals.Sheet6.Cells[2, col2].Value.ToString()) < _op_tenor && _op_tenor < double.Parse(Globals.Sheet6.Cells[2, col2 + 1].Value.ToString()))
                    {
                        percent = (double.Parse(Globals.Sheet6.Cells[_date_row, col2].Value.ToString()) + double.Parse(Globals.Sheet6.Cells[_date_row, col2 + 1].Value.ToString())) / 2;
                        break;
                    }
                }
                else
                {
                    percent = double.Parse(Globals.Sheet6.Cells[_date_row, col2].Value.ToString());
                    break;
                }

                col2++;
            }

            return percent;
        }
    }
}
