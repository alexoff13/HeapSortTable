using System.IO;
using HeapSortManual;

namespace HeapSortManual
{
    class Program
    {
        static void Main(string[] args)
        {
            var heapSort=new HeapSortManual();
            var reportFile=heapSort.HeapSort(new long[10]{3, 7, 1, 2, 9, 4, 3, 6, 5, 2},10);
            File.WriteAllBytes("Report.xlsx", reportFile);
        }
    }
}