using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Report report = new Report();

                var _employees = new List<Employee>();
                Random random = new Random();
                for (int i = 0; i < 10; i++)
                {
                    _employees.Add(new Employee()
                    {
                        Id = i,
                        Name = "Employee " + i,
                        DOB = new DateTime(1999, 1, 1).AddMonths(i),
                        Salary = random.Next(100, 10000)
                    });
                }

                report.CreateExcelDoc(@"D:\Report.xlsx", _employees);

            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }
    }
}
