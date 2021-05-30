using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using ExcelDataReader;

namespace YellowCanary.Super
{
    static class Program
    {
        record PersonQuarter(string EmployeeCode, YearQuarter YearQuarter);
        record CodePayment(string PaymentCode, Decimal Amount);
        
        
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please pass a file path.");
                return;
            }

            var filePath = args[0];
            if (!File.Exists(filePath))
            {
                Console.WriteLine("Could not find file.");
                return;
            }
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);

            var distributions = ProcessDistributions(reader);
            reader.NextResult();
            var employeePayments = ProcessPayments(reader);
            reader.NextResult();
            var oteCodes = ProcessPaymentTypes(reader);

            PrintResults(employeePayments, oteCodes, distributions);
        }

        private static void PrintResults(Dictionary<string, Dictionary<YearQuarter, List<CodePayment>>> employeePayments, List<string> oteCodes, Dictionary<PersonQuarter, decimal> distributions)
        {
            foreach (var employee in employeePayments)
            {
                var employeeCode = employee.Key;
                Console.WriteLine($"Employee: {employeeCode}");
                var codePaymentsByYearQuarter = employee
                    .Value
                    .OrderBy(x => x.Key.year)
                    .ThenBy(x => x.Key.quarter);

                foreach (var codePaymentYearQuarter in codePaymentsByYearQuarter)
                {
                    var yearQuarter = codePaymentYearQuarter.Key;
                    var totalOtcForYearQuarter = codePaymentYearQuarter.Value
                        .Where(x => oteCodes.Contains(x.PaymentCode))
                        .Sum(x => x.Amount);
                    var superPayable = totalOtcForYearQuarter * 0.095m;
                    var personYearQuarter = new PersonQuarter(employeeCode, yearQuarter);
                    Console.WriteLine($"{yearQuarter.quarter}/{yearQuarter.year}");
                    distributions.TryGetValue(personYearQuarter, out var totalDistributonForYearQuarter);
                    Console.WriteLine($"Total OTE earnings: {totalOtcForYearQuarter:C}");
                    Console.WriteLine($"Total Super Payable {superPayable:C}");
                    Console.WriteLine($"Total distributions: {totalDistributonForYearQuarter:C}");
                    Console.WriteLine($"Discrepancy: {totalDistributonForYearQuarter - superPayable:C}");
                    Console.WriteLine();
                }

                Console.WriteLine("=====================================================================");
            }
        }

        private static List<string> ProcessPaymentTypes(IExcelDataReader reader)
        {
            var otcTypes = new List<string>();
            reader.Read();

            while (reader.Read())
            {
                try
                {
                    var code = reader.GetString(0);
                    var type = reader.GetString(1);
                    if (type.Equals("ote", StringComparison.InvariantCultureIgnoreCase))
                    {
                        otcTypes.Add(code);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }

            return otcTypes;
        }

        private static Dictionary<string, Dictionary<YearQuarter, List<CodePayment>>> ProcessPayments(IExcelDataReader reader)
        {
            reader.Read();
            var employeePayments = new Dictionary<string, Dictionary<YearQuarter, List<CodePayment>>>();

            while (reader.Read())
            {
                try
                {
                    var date = reader.GetDateTime(1);
                    var yearQuarter = QuarterUtil.GetQuarter(date);
                    var employeeCode = reader.GetDouble(2).ToString(CultureInfo.InvariantCulture);
                    var paymentCode = reader.GetString(3);
                    var amount = Decimal.Parse(reader[4].ToString());

                    if (!employeePayments.TryGetValue(employeeCode, out var payments))
                    {
                        payments = new Dictionary<YearQuarter, List<CodePayment>>();
                        employeePayments.Add(employeeCode, payments);
                    }

                    CodePayment codePayment = new(paymentCode, amount);

                    if (!payments.TryAdd(yearQuarter, new List<CodePayment> {codePayment}))
                    {
                        payments[yearQuarter].Add(codePayment);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }

            return employeePayments;
        }
        
        private static Dictionary<PersonQuarter, Decimal> ProcessDistributions(IExcelDataReader reader)
        {
            reader.Read();
            var distributions = new Dictionary<PersonQuarter, Decimal>();
            while (reader.Read())
            {
                try
                {
                    var amount = Decimal.Parse(reader[0].ToString());
                    var date = DateTime.Parse(reader.GetString(1));
                    var employeeCode = reader.GetDouble(4).ToString(CultureInfo.InvariantCulture);

                    var yearQuarter = QuarterUtil.GetDistributionQuarter(date);
                    PersonQuarter personQuarterKey = new(employeeCode, yearQuarter);
                    if (!distributions.TryAdd(personQuarterKey, amount))
                    {
                        distributions[personQuarterKey] += amount;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }

            return distributions;
        }
    }
}
