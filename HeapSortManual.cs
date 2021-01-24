using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace HeapSortManual
{
    public class HeapSortManual
    {
        private int _columns = 1;

        private void Heapify(ref long[] arr, int n, int i, ref int numberOfComparisons, ref int numberOfExchanges,
            ExcelWorksheet sheet)
        {
            var comparisons = numberOfComparisons;
            var exchanges = numberOfExchanges;

            var column = _columns;
            var largest = i;
            var l = 2 * i + 1;
            var r = 2 * i + 2;

            if (l >= n && r >= n) return;

            _columns++;

            if (l < n)
            {
                numberOfComparisons += 1;
                sheet.Cells[column, l + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[column, l + 1].Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
            }

            if (r < n)
            {
                numberOfComparisons += 1;
                sheet.Cells[column, r + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[column, r + 1].Style.Fill.BackgroundColor.SetColor(Color.Aquamarine);
            }

            sheet.Cells[column, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[column, i + 1].Style.Fill.BackgroundColor.SetColor(Color.Gold);

            for (var j = 0; j < arr.Length; j++) sheet.Cells[column, j + 1].Value = arr[j];

            if (l < n && arr[l] > arr[largest]) largest = l;

            if (r < n && arr[r] > arr[largest]) largest = r;

            sheet.Cells[column, arr.Length + 1].Value = numberOfComparisons - comparisons;
            sheet.Cells[column, arr.Length + 2].Value = numberOfExchanges - exchanges;

            if (largest != i)
            {
                numberOfExchanges++;
                sheet.Cells[column, arr.Length + 2].Value = numberOfExchanges - exchanges;
                (arr[i], arr[largest]) = (arr[largest], arr[i]);

                Heapify(ref arr, n, largest, ref numberOfComparisons, ref numberOfExchanges, sheet);
            }
        }

        public byte[] HeapSort(long[] arr, int n)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets
                .Add("HeapSort");

            int numberOfComparisons = 0, numberOfExchanges = 0;

            // Build max heap
            for (var i = n / 2 - 1; i >= 0; i--)
                Heapify(ref arr, n, i, ref numberOfComparisons, ref numberOfExchanges, sheet);

            // Heap sort
            for (var i = n - 1; i >= 0; i--)
            {
                (arr[0], arr[i]) = (arr[i], arr[0]);

                numberOfExchanges++;

                // Heapify root element to get highest element at root again
                Heapify(ref arr, i, 0, ref numberOfComparisons, ref numberOfExchanges, sheet);
            }

            return package.GetAsByteArray();
        }
    }
}