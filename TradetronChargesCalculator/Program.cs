using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;

namespace TradetronChargesCalculator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string path = args[0];
            List<PnL> pnls = new List<PnL>();
            if(Directory.Exists(path))
            {
                var files = Directory.GetFiles(path, "*.*", SearchOption.AllDirectories);
                foreach (var file in files)
                {
                    var pnl = HandleFile(file);
                    if(pnl != null)
                    {
                        pnl.Identifier = file.Remove(0, path.Length + 1);
                        pnls.Add(pnl);
                    }
                }
            }
            else if(File.Exists(path))
            {
                var pnl = HandleFile(path);
                if (pnl != null)
                {
                    pnl.Identifier = Path.GetFileName(path);
                    pnls.Add(pnl);
                }
            }

            PnL total = new PnL();
            total.Identifier = "Total";
            total.Gross = pnls.Sum(pnl => pnl.Gross);
            total.Charges = pnls.Sum(p => p.Charges);
            total.NetProfit = pnls.Sum(p => p.NetProfit);

            pnls.Add(total);

            int nameWidth = pnls.Max(p => p.Identifier.Length) + 2;
            int grossWidth = pnls.Max(p => p.Gross.ToString().Length) +2;
            int chargesWidth = pnls.Max(p => p.Charges.ToString().Length) + 2;
            int chargeRatioWidth = Math.Max("Charges Ratio".Length, pnls.Max(p => p.ChargeRatio.ToString().Length)) + 2;
            int netProfitWidth = pnls.Max(p => p.NetProfit.ToString().Length) + 2;

            pnls.Remove(total);

            pnls = pnls.OrderBy(p => p.Identifier).ToList();

            Console.WriteLine();
            Console.WriteLine();
            PrintWithWidth("Name", nameWidth, false);
            Console.Write(" |");
            PrintWithWidth("Gross", grossWidth, true);
            Console.Write(" |");
            PrintWithWidth("Charges", chargesWidth, true);
            Console.Write(" |");
            PrintWithWidth("Charges Ratio", chargeRatioWidth, true);
            Console.Write(" |");
            PrintWithWidth("Net", netProfitWidth, true);
            Console.WriteLine(" |");

            int cols = nameWidth + grossWidth + chargesWidth + chargeRatioWidth + netProfitWidth + 8;

            for(int i = 0; i < cols; i++)
            {
                Console.Write("-");
            }
            Console.WriteLine();

            foreach (var pnl in pnls)
            {
                PrintWithWidth(pnl.Identifier, nameWidth, false);
                Console.Write(" |");
                PrintWithWidth(pnl.Gross.ToString(), grossWidth, true);
                Console.Write(" |");
                PrintWithWidth(pnl.Charges.ToString(), chargesWidth, true);
                Console.Write(" |");
                PrintWithWidth(pnl.ChargeRatio.ToString(), chargeRatioWidth, true);
                Console.Write(" |");
                PrintWithWidth(pnl.NetProfit.ToString(), netProfitWidth, true);
                Console.WriteLine(" |");
            }
            for (int i = 0; i < cols; i++)
            {
                Console.Write("-");
            }
            Console.WriteLine();

            PrintWithWidth(total.Identifier, nameWidth, false);
            Console.Write(" |");
            PrintWithWidth(total.Gross.ToString(), grossWidth, true);
            Console.Write(" |");
            PrintWithWidth(total.Charges.ToString(), chargesWidth, true);
            Console.Write(" |");
            PrintWithWidth(total.ChargeRatio.ToString(), chargeRatioWidth, true);
            Console.Write(" |");
            PrintWithWidth(total.NetProfit.ToString(), netProfitWidth, true);
            Console.WriteLine(" |");

        }

        static void PrintWithWidth(string text, int width, bool righAligned) 
        {
            int space = width - text.Length;
            if (righAligned)
            {
                for (int i = 0; i < space; i++)
                {
                    Console.Write(" ");
                }
                Console.Write(text);
            } 
            else
            {
                Console.Write(text);
                for (int i = 0; i < space; i++)
                {
                    Console.Write(" ");
                }
            }
        }

        static PnL HandleFile(string path)
        {
            var ext = Path.GetExtension(path).ToUpperInvariant();

            if (ext == ".CSV")
            {
                Console.WriteLine("Reading: " + path);
                return ProcessCsvFile(path);
            }
            else if (ext == ".XLSX")
            {
                var name = Path.GetFileNameWithoutExtension(path);
                if (name.StartsWith("~$")) return null;
                Console.WriteLine("Reading: " + path);
                var csvFile = ConverExcelToCsv(path);
                try 
                { 
                    return ProcessCsvFile(csvFile); 
                }
                finally
                {
                    File.Delete(csvFile);
                }
            }
            return null;
        }

        static string ConverExcelToCsv(string filePath)
        {
            string tempPath;

            tempPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(filePath) + "_" + Guid.NewGuid() + ".csv");

            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
            _Worksheet xlWorksheet = xlWorkbook.Sheets[1];

            xlWorksheet.SaveAs(tempPath, XlFileFormat.xlCSV);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            return tempPath;
        }

        static void ProcessExcelFile(string filePath)
        {
            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
            _Worksheet xlWorksheet = xlWorkbook.Sheets[1];

            
            Range xlRange = xlWorksheet.UsedRange;
            try
            {
                int slNoColPos = FindColumnPosition(xlRange, "S No");
                int quantityColPos = FindColumnPosition(xlRange, "Quantity");
                int amountColPos = FindColumnPosition(xlRange, "Amount");

                int priceColPos = FindColumnPosition(xlRange, "Price");
                int strikeColPos = FindColumnPosition(xlRange, "Strike");
                int underlyingPriceColPos = FindColumnPosition(xlRange, "Underlying Price");
                int optionTypeColPos = FindColumnPosition(xlRange, "Option Type");

                double grossTotal = 0;
                double charges = 0;
                int i = 2;
                while (true)
                {
                    string slNo = xlRange[i, slNoColPos].Value?.ToString();
                    if (String.IsNullOrEmpty(slNo))
                    {
                        break;
                    }

                    var quantityStr = xlRange[i, quantityColPos].Value?.ToString();
                    if (String.IsNullOrWhiteSpace(quantityStr)) continue;

                    var rowAmountStr = xlRange[i, amountColPos].Value.ToString();
                    int quantity = int.Parse(quantityStr);
                    double price = double.Parse(xlRange[i, priceColPos].Value.ToString());
                    double strike = double.Parse(xlRange[i, strikeColPos].Value.ToString());
                    double underlyingPrice = double.Parse(xlRange[i, underlyingPriceColPos].Value.ToString());

                    string optionType = xlRange[i, optionTypeColPos].Value.ToString();

                    double premium = Math.Abs(quantity) * price;

                    //STT
                    if (quantity > 0)
                    {
                        double intrinsicValue = price - strike;
                        if (intrinsicValue > 0)
                        {
                            charges += (0.125 * intrinsicValue) / 100;
                        }
                    }
                    else
                    {
                        charges += premium * 0.0625 / 100;
                    }
                    //transaction charges

                    double transcharges = premium * 0.05 / 100;
                    double sebiCharges = 10 * premium / (10000000);

                    //stamp charges
                    if (quantity > 0)
                    {
                        charges += 0.003 * premium / 100;
                    }

                    double gst = 18 * (transcharges + sebiCharges) / 100;

                    charges += transcharges + sebiCharges + gst;

                    grossTotal += double.Parse(rowAmountStr);
                    i++;
                }


                charges = Math.Round(charges, 2);
                double netProfit = grossTotal * -1 - charges;

                Console.WriteLine($" => Gross Profit: {grossTotal * -1}, Charges: {charges}, Net Profit: {netProfit}");

                double incomeTax = 15 * netProfit / 100;

                double strategyCharges = 5 * netProfit / 100;

                double finalProfit = netProfit - incomeTax - strategyCharges;
                finalProfit = Math.Round(finalProfit, 0);
                Console.WriteLine("Profit After Income Tax and Strategy Charges: " + finalProfit);
            }
            finally
            {
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
        }

        private static int FindColumnPosition(Range xlRange, string headerName)
        {
            for (int i = 1; i < 100; i++)
            {
                if (xlRange.Cells[1, i].Value.ToString() == headerName)
                {
                    return i;
                }
            }

            return -1;
        }
        static PnL ProcessCsvFile(string filePath)
        {
            var lines = File.ReadAllLines(filePath);

            var headers = lines[0].Split(',');

            int quantityColPos = FindIndexOfString(headers, "Quantity");
            int amountColPos = FindIndexOfString(headers, "Amount");

            int priceColPos = FindIndexOfString(headers, "Price");
            int strikeColPos = FindIndexOfString(headers, "Strike");
            int underlyingPriceColPos = FindIndexOfString(headers, "Underlying Price");
            int optionTypeColPos = FindIndexOfString(headers, "Option Type");
            int instrumentColPos = FindIndexOfString(headers, "Instrument Type");
            double grossTotal = 0;

            double charges = 0;

            for (int i = 1; i < lines.Length; i++)
            {
                var rowParts = lines[i].Split(',');
                var quantityStr = rowParts[quantityColPos];
                if(String.IsNullOrWhiteSpace(quantityStr)) continue;
                
                var rowAmountStr = rowParts[amountColPos];
                int quantity = int.Parse(quantityStr);
                double price;
                string priceStr = rowParts[priceColPos];
                bool tradePending = false;
                if (!String.IsNullOrWhiteSpace(priceStr))
                {
                    price = double.Parse(priceStr);
                }
                else
                {
                    tradePending = true;
                    price = 2;
                }
                double strike;
                double.TryParse(rowParts[strikeColPos], out strike);
                
                string instrumentType = rowParts[instrumentColPos];
                
                //double underlyingPrice = double.Parse(rowParts[underlyingPriceColPos]);

                string optionType = rowParts[optionTypeColPos];

                double premium = Math.Abs(quantity) * price;

                if (instrumentType == "OPTIDX")
                {
                    Trace.Assert(strike != 0, "Strike price should not be zero for index option");
                    //STT
                    if (quantity > 0)
                    {
                        double intrinsicValue = price - strike;
                        if (intrinsicValue > 0)
                        {
                            charges += (0.125 * intrinsicValue) / 100;
                        }
                    }
                    else
                    {
                        charges += premium * 0.0625 / 100;
                    }
                    //transaction charges

                    double transcharges = premium * 0.05 / 100;
                    double sebiCharges = 10 * premium / (10000000);

                    //stamp charges
                    if (quantity > 0)
                    {
                        charges += 0.003 * premium / 10000000;
                    }

                    double gst = 18 * (transcharges + sebiCharges) / 100;

                    charges += transcharges + sebiCharges + gst;
                    if(tradePending)
                    {
                        grossTotal += quantity * price;
                    }
                    else
                    {
                        grossTotal += double.Parse(rowAmountStr);
                    }
                } 
                else if (instrumentType == "OPTCOM")
                {
                    Trace.Assert(strike != 0, "Strike price should not be zero for Commodity option");
                    //STT
                    if (quantity > 0)
                    {
                        
                    }
                    else
                    {
                        charges += premium * 0.05 / 100;
                    }
                    //transaction charges

                    double transcharges = premium * 0.05 / 100;
                    double sebiCharges = 10 * premium / (10000000);

                    //stamp charges
                    if (quantity > 0)
                    {
                        charges += 0.003 * premium / 10000000;
                    }

                    double gst = 18 * (transcharges + sebiCharges) / 100;

                    charges += transcharges + sebiCharges + gst;

                    grossTotal += double.Parse(rowAmountStr);
                }
                else if(instrumentType == "FUTCOM")
                {
                    //STT
                    if (quantity > 0)
                    {

                    }
                    else
                    {
                        charges += premium * 0.01 / 100;
                    }
                    //transaction charges

                    double transcharges = premium * 0.0026 / 100;
                    double sebiCharges = 10 * premium / (10000000);

                    //stamp charges
                    if (quantity > 0)
                    {
                        charges += 0.002 * premium / 10000000;
                    }

                    double gst = 18 * (transcharges + sebiCharges) / 100;

                    charges += transcharges + sebiCharges + gst;

                    grossTotal += double.Parse(rowAmountStr);
                }
                else
                {
                    return null;
                }
            }
            grossTotal = Math.Round(grossTotal, 0);
            charges = Math.Round(charges, 0);
            double netProfit = grossTotal * -1 - charges;
            double chargeRatio = charges / grossTotal;
            chargeRatio = Math.Round(chargeRatio, 2) * -1;

            netProfit = Math.Round(netProfit, 0);
            PnL PnL = new PnL { Charges = charges, Gross = grossTotal * -1, NetProfit = netProfit, ChargeRatio = chargeRatio };
            //Console.WriteLine($" => Gross Profit: {grossTotal * -1}, Charges: {charges} (Charge Ration: {chargeRatio}), Net Profit: {netProfit}");

            return PnL;
        }

        private static int FindIndexOfString(string[] headers, string headerName)
        {
            for (int i = 0; i < headers.Length; i++)
            {
                if (headers[i] == headerName)
                {
                    return i;
                }
            }

            return -1;
        }
    }

    internal class PnL
    {
        public string Identifier { get; set; }
        public double Gross { get; set; }

        public double Charges { get; set; }

        public double ChargeRatio { get; set; }

        public double NetProfit { get; set; }
    }
}
