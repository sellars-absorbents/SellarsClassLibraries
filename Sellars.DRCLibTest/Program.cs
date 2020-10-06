using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sellars.DRCLib;

namespace Sellars.DRCLibTest
{
    class Program
    {
        static void Main(string[] args)
        {
            QueuedWebSalesOrderTests qSOTests = new QueuedWebSalesOrderTests();

            qSOTests.RunAddRecordTest();
        }
    }
}
