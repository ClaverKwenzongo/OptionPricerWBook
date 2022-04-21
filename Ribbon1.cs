//using Microsoft.Office.Tools.Ribbon;
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;

//namespace OptionPricerWBook
//{
//    public partial class Ribbon1
//    {
//        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
//        {

//        }
//    }
//}

using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.Globalization;
using System.Diagnostics;
using MathNet.Numerics.Statistics;


namespace OptionPricerWBook
{
    public partial class Ribbon1
    {
        int row_start = 5;

        getSharePrice getShare = new getSharePrice();
        getImpliedVol getVol = new getImpliedVol();
        getYield getDividentYield = new getYield();
        getRates getRates = new getRates();
        standAloneEquity getStandAloneEquity = new standAloneEquity();
        standAloneRates getStandAloneRates = new standAloneRates();
        standAloneVol getStandAloneVol = new standAloneVol();
        standAloneDividend getStandAloneDividend = new standAloneDividend();
        staticTotal getStaticTotal = new staticTotal();
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        static int row_count()
        {
            //count how many shares are in the portfolio: this is needed in the computation of the average P&L....
            int col = 2;
            int count = 0;
            while (string.IsNullOrWhiteSpace(Globals.Sheet1.Cells[2, col].Value?.ToString()) == false)
            {
                count++;
                col++;
            }

            return count;
        }

        static int getCol_Increase()
        {
            int col_count = 1;
            int col3 = 2;
            while (string.IsNullOrWhiteSpace(Globals.Sheet2.Cells[2, col3].Value?.ToString()) == false)
            {
                col_count++;
                col3++;
            }

            return col_count;
        }

        static double tenor(string start, string end)
        {
            CultureInfo culture = new CultureInfo("es-ES");

            DateTime end_date = DateTime.Parse(end, culture);
            DateTime start_date = DateTime.Parse(start, culture);

            TimeSpan interval = end_date - start_date;
            double days = (double)interval.TotalDays;

            return days;
        }

        static DateTime getDate(string _date)
        {
            CultureInfo cultureInfo = new CultureInfo("en-US");
            DateTime wsDate = DateTime.Parse(_date, cultureInfo);

            return wsDate;
        }

        private void ValuateBtn_Click(object sender, RibbonControlEventArgs e)
        {
            int col_inc = getCol_Increase();

            double portfolio_val = 0;

            int j = 4;
            while (string.IsNullOrWhiteSpace(Globals.Sheet4.Cells[row_start + 2, j].Value?.ToString()) == false)
            {
                string myStartDate = Globals.Sheet4.Cells[row_start + 3, j].Value.ToString();
                //Debug.WriteLine(myStartDate);
                string user_share = Globals.Sheet4.Cells[row_start + 2, j].Value;
                double S_0 = getShare._getSharePrice(myStartDate, user_share.ToUpper());

                double Tenor = tenor(Globals.Sheet4.Cells[row_start + 3, j].Value.ToString(), Globals.Sheet4.Cells[row_start + 4, j].Value.ToString());

                double impl_Vol = getVol.getImpl_Vol(Tenor, myStartDate, col_inc, user_share.ToUpper());

                double div_yield = getDividentYield.getDiv_Yield(Tenor, myStartDate, col_inc, user_share.ToUpper());

                double rate = getRates.getRate(Tenor, myStartDate);

                double K = Globals.Sheet4.Cells[row_start + 5, j].Value;

                string option_type = Globals.Sheet4.Cells[row_start + 7, j].Value;

                int psi = 0;
                if (option_type.ToUpper() == "PUT")
                {
                    psi = -1;
                }
                else
                {
                    psi = 1;
                }

                string option_pos = Globals.Sheet4.Cells[row_start + 6, j].Value;
                double size = Globals.Sheet4.Cells[row_start + 8, j].Value;

                if (option_pos.ToUpper() == "SHORT")
                {
                    size = -size;
                }
                else
                {
                    size = size;
                }


                EuropeanOptionPricer pricer = new EuropeanOptionPricer(K, psi, Tenor);

                double mkt_op_price = pricer.optionPrice(S_0, rate, impl_Vol, div_yield);
                ImpliedVolatility newton_implied_vol = new ImpliedVolatility(mkt_op_price, K, Tenor, S_0, rate, Math.Pow(10, -8), div_yield);

                Globals.Sheet4.Cells[row_start + 11, j].Value = pricer.optionPrice(S_0, rate, impl_Vol, div_yield);
                Globals.Sheet4.Cells[row_start + 12, j].Value = pricer.sensitivity(S_0, rate, div_yield, impl_Vol, "Delta");
                Globals.Sheet4.Cells[row_start + 13, j].Value = pricer.sensitivity(S_0, rate, div_yield, impl_Vol, "Gamma");
                Globals.Sheet4.Cells[row_start + 14, j].Value = pricer.sensitivity(S_0, rate, div_yield, impl_Vol, "Vega");
                Globals.Sheet4.Cells[row_start + 15, j].Value = pricer.sensitivity(S_0, rate, div_yield, impl_Vol, "Rho");
                Globals.Sheet4.Cells[row_start + 16, j].Value = pricer.sensitivity(S_0, rate, div_yield, impl_Vol, "Epsilon");
                Globals.Sheet4.Cells[row_start + 17, j].Value = newton_implied_vol.newton_vol(psi);

                portfolio_val += size * pricer.optionPrice(S_0, rate, impl_Vol, div_yield);

                j++;
            }

            var val = string.Format("{0:C}", portfolio_val);

            Globals.Sheet4.Cells[row_start, 3].Value = val;
        }

        private void HSVaRBtn_Click_1(object sender, RibbonControlEventArgs e)
        {
            int count_rows = row_count();
            int count_cols = getCol_Increase();

            getStandAloneEquity.standAlone_equity(row_start, count_rows, count_cols);
            getStandAloneRates.standAlone_rate(row_start, count_rows, count_cols);
            getStandAloneDividend.standAlone_div(row_start, count_rows, count_cols);
            getStandAloneVol.standAlone_vol(row_start, count_rows, count_cols);
            getStaticTotal.static_Total(row_start, count_rows, count_cols);
        }
    }
}
