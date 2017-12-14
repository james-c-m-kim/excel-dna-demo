using System;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Threading;
using Microsoft.Office.Interop.Excel;

namespace demo_plugin
{
    public static class DemoUdfFunctions
    {
        [ExcelFunction(Name = "hello")]
        public static object Hello(string nameOfSomeone)
        {
            if (string.IsNullOrEmpty(nameOfSomeone))
            {
                return "Hello, there!!!";
            }

            return $"Hello, {nameOfSomeone}, how's it going, eh?";
        }

        [ExcelFunction(Name = "getliveupdate")]
        public static object GetLiveUpdate(string currencyPair)
        {            
            var app = (Application)ExcelDnaUtil.Application;
            var range = app.ActiveCell;
            var reference = new ExcelReference(range.Row - 1, range.Row - 1, range.Column - 1, range.Column - 1);
            Task.Factory.StartNew(() =>
            {
                for (var i = 1; i < 30; i++)
                {
                    Thread.Sleep(1000);
                    ExcelAsyncUtil.QueueAsMacro(()=> reference.SetValue(i));   
                }

                var referenceDone = new ExcelReference(range.Row - 1, range.Row - 1, range.Column, range.Column);
                ExcelAsyncUtil.QueueAsMacro(()=> referenceDone.SetValue("All done!"));
            });

            return 0;
        }

        [ExcelFunction(Name = "dumpdata")]
        public static object DumpData()
        {
            var app = (Application)ExcelDnaUtil.Application;
            var ws = (Worksheet) app.ActiveSheet;

            var dataArray = new object[1000, 1000];
            var firstRowArray = new object[1, 999];
            var rnd = new Random(1000);

            var startCell = (Range) ws.Cells[app.ActiveCell.Row + 1, app.ActiveCell.Column];
            var endCell =
                (Range) ws.Cells[startCell.Row + dataArray.GetLength(0) - 1,
                    startCell.Column + dataArray.GetLength(1) - 1];

            var firstStartCell = (Range) ws.Cells[app.ActiveCell.Row, app.ActiveCell.Column + 1];
            var firstEndCell = (Range) ws.Cells[firstStartCell.Row, firstStartCell.Column + firstRowArray.GetLength(1) - 1];

            var firstRowRange = ws.Range[firstStartCell, firstEndCell];
            var writeRange = ws.Range[startCell, endCell];

            Task.Factory.StartNew(() =>
            {
                // special: write the first row w/ bypass of the formula cell to prevent
                // overwriting the formula itself...
                for (var firstIdx = 0; firstIdx <= firstRowArray.GetUpperBound(1); firstIdx++)
                {
                    firstRowArray[0, firstIdx] = rnd.NextDouble();
                }

                for (var rowIdx = 0; rowIdx <= dataArray.GetUpperBound(0); rowIdx++)
                {
                    for (var colIdx = 0; colIdx <= dataArray.GetUpperBound(1); colIdx++)
                    {
                        dataArray[rowIdx, colIdx] = rnd.NextDouble();
                    }
                }

                ExcelAsyncUtil.QueueAsMacro(()=>
                {
                    firstRowRange.Value2 = firstRowArray;
                    writeRange.Value2 = dataArray;
                });
            });

            return rnd.NextDouble();
        }
    }
}
