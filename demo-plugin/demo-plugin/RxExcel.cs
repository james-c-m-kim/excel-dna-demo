using System;

namespace ExcelDna.Integration.RxExcel
{
    public static class RxExcel
    {
        public static IExcelObservable ToExcelObservable<T>(this IObservable<T> observable)
        {
            return new ExcelObservable<T>(observable);
        }

        public static object Observe<T>(string functionName, object parameters, Func<IObservable<T>> observableSource)
        {
            return ExcelAsyncUtil.Observe(functionName, parameters, () => observableSource().ToExcelObservable());
        }
    }
}
