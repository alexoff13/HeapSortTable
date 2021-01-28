using System.IO;

namespace HeapSortManual
{
    class Program
    {
        static void Main()
        {
            var heapSort=new HeapSortManual();
            var reportFile=heapSort.HeapSort(new double[]{5,4,3,2,1},5);
            File.WriteAllBytes("Report.xlsx", reportFile);
        }
    }
}