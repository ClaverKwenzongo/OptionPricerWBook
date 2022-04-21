using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.Globalization;
using System.Diagnostics;
using MathNet.Numerics.Statistics;
using System.Windows.Forms;

namespace OptionPricerWBook
{
    public class standAloneDividend
    {
        getSharePrice getShare = new getSharePrice();
        getImpliedVol getVol = new getImpliedVol();
        getYield getDividentYield = new getYield();
        getRates getRates = new getRates();
        getPercentile percentile = new getPercentile();

        static DateTime getDate(string _date)
        {
            CultureInfo cultureInfo = new CultureInfo("en-US");
            DateTime wsDate = DateTime.Parse(_date, cultureInfo);

            return wsDate;
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
        public void standAlone_div(int row_start, int row_count, int col_up)
        {
            List<double> porfolio_pl = new List<double>();

            double sum_ = 0;
            double _sum_ = 0;

            //Lock the start date so that we can fix the other risk factors.
            DateTime _start_date = getDate(Globals.Sheet1.Cells[3, 1].Value.ToString());
            Debug.WriteLine(_start_date);

            int col_j = 4;

            while (string.IsNullOrWhiteSpace(Globals.Sheet4.Cells[row_start + 2, col_j].Value?.ToString()) == false)
            {
                double sensitivity = 0;
                double size = 0;

                if (string.IsNullOrWhiteSpace(Globals.Sheet4.Cells[row_start + 16, col_j].Value?.ToString()) == true)
                {
                    MessageBox.Show("To calculate risk metrics, you must valuate the portfolio first so the sensitivities are known.");
                }
                else
                {
                    //Get calculated sensititivity: delta for each share.
                    sensitivity = double.Parse(Globals.Sheet4.Cells[row_start + 16, col_j].Value.ToString());

                    //Get the amount of shares
                    size = double.Parse(Globals.Sheet4.Cells[row_start + 8, col_j].Value.ToString());
                }

                int pos = 0;
                string position = Globals.Sheet4.Cells[row_start + 6, col_j].Value;
                if (position.ToUpper() == "SHORT")
                {
                    pos = -1;
                }
                else
                {
                    pos = 1;
                }

                //Fix the other risk factors......................................................................................
                string lock_start_date = _start_date.ToString("dd/MM/yyyy");
                string user_share = Globals.Sheet4.Cells[row_start + 2, col_j].Value;
                string mat_date = Globals.Sheet4.Cells[row_start + 4, col_j].Value;

                double Tenor_ = tenor(lock_start_date, mat_date);
                double spot = getShare._getSharePrice(lock_start_date, user_share.ToUpper());
                double rf = getRates.getRate(Tenor_, lock_start_date);
                double vol = getVol.getImpl_Vol(Tenor_, lock_start_date, col_up, user_share.ToUpper());
                double K = Globals.Sheet4.Cells[row_start + 5, col_j].Value;
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                int date_r = 3;
                while (date_r < row_count + 1)
                {
                    //Check in the shares columns whether the next row is null or white space
                    if (string.IsNullOrWhiteSpace(Globals.Sheet1.Cells[date_r + 1, 2].Value?.ToString()) == false)
                    {
                        double p_l = 0;

                        DateTime date_start_1 = getDate(Globals.Sheet1.Cells[date_r, 1].Value.ToString());
                        DateTime date_start_2 = getDate(Globals.Sheet1.Cells[date_r + 1, 1].Value.ToString());

                        double Tenor_1 = tenor(date_start_1.ToString("dd/MM/yyyy"), mat_date);
                        double Tenor_2 = tenor(date_start_2.ToString("dd/MM/yyyy"), mat_date);

                        string option_type = Globals.Sheet4.Cells[row_start + 7, col_j].Value;


                        int psi = 0;
                        if (option_type.ToUpper() == "PUT")
                        {
                            psi = -1;
                        }
                        else
                        {
                            psi = 1;
                        }

                        EuropeanOptionPricer pricer_1 = new EuropeanOptionPricer(K, psi, Tenor_1);
                        EuropeanOptionPricer pricer_2 = new EuropeanOptionPricer(K, psi, Tenor_2);

                        //Change the volatility...............................................................................
                        double q_1 = getDividentYield.getDiv_Yield(Tenor_1, date_start_1.ToString("dd/MM/yyyy"), col_up, user_share.ToUpper());
                        double q_2 = getDividentYield.getDiv_Yield(Tenor_2, date_start_2.ToString("dd/MM/yyyy"), col_up, user_share.ToUpper());
                        //////////////////////////////////////////////////////////////////////////////////////////////////

                        double price_1 = pricer_1.optionPrice(spot, rf, vol, q_1);
                        double price_2 = pricer_2.optionPrice(spot, rf, vol, q_2);

                        p_l = Math.Log(price_1 / price_2);

                        porfolio_pl.Add(p_l);
                    }
                    else
                    {
                        break;
                    }

                    date_r++;
                }

                col_j++;

                double[] portfolio_pl_array = porfolio_pl.ToArray();

                double percentile_ = percentile.Percentile(portfolio_pl_array, 0.01);
                double _percentile_ = percentile.Percentile(portfolio_pl_array, 0.025);

                sum_ += percentile_ * size * pos;
                _sum_ += _percentile_ * size * pos;
            }

            Globals.Sheet4.Cells[row_start + 22, 8].Value = sum_;
            Globals.Sheet4.Cells[row_start + 23, 8].Value = _sum_;
        }

    }
}


