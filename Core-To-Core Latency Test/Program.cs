using System;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Drawing;

class Program
{
    [DllImport("kernel32.dll")]
    private static extern IntPtr GetCurrentThread();

    [DllImport("kernel32.dll")]
    private static extern IntPtr SetThreadAffinityMask(IntPtr hThread, IntPtr dwThreadAffinityMask);

    private const long Iterations = 10000000;
    private const string OutputFileName = "CoreToCoreLatencies.xlsx";
    private static long bounceValue;
    private static ManualResetEventSlim startSignal = new ManualResetEventSlim(false);
    private static ManualResetEventSlim endSignal1 = new ManualResetEventSlim(false);
    private static ManualResetEventSlim endSignal2 = new ManualResetEventSlim(false);

    static void Main(string[] args)
    {
        try
        {
            Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.RealTime;
            Process.GetCurrentProcess().PriorityBoostEnabled = true;

            int numCores = Environment.ProcessorCount;
            double[,] latencies = new double[numCores, numCores];
            string cpuName = GetCpuName();

            Console.WriteLine($"UXTU V3 Core-to-Core Latency Test");
            Console.WriteLine($"CPU: {cpuName}");
            Console.WriteLine($"Number of Cores/Threads: {numCores}");
            Console.WriteLine($"Your results will be exported to a spreadsheet with conditional formatting applied\n");
            Thread.Sleep(1000);
            for (int i = 0; i < numCores; i++)
            {
                for (int j = 0; j < numCores; j++)
                {
                    if (i != j)
                    {
                        latencies[i, j] = Math.Round(MeasureLatency(i, j, Iterations), 2);
                        Console.WriteLine($"Latency from core {i} to core {j}: {latencies[i, j]:F2} ns");
                    }
                    else latencies[i, j] = 0.0;
                }
            }

            SaveLatenciesToExcel(latencies, OutputFileName, cpuName);
            Console.WriteLine($"Core-to-Core Latency Matrix saved to {OutputFileName}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    static string GetCpuName()
    {
        string cpuName = string.Empty;
        try
        {
            using (var searcher = new ManagementObjectSearcher("select Name from Win32_Processor"))
            {
                foreach (var item in searcher.Get()) cpuName = item["Name"].ToString();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting CPU name: {ex.Message}");
        }
        return cpuName;
    }

    static double MeasureLatency(int core1, int core2, long iterations)
    {
        bounceValue = 0;
        double latency = 0;

        Task t1 = new Task(() => LatencyTestThread(core1, 1, 0));
        Task t2 = new Task(() => LatencyTestThread(core2, 2, 1));

        startSignal.Reset();
        endSignal1.Reset();
        endSignal2.Reset();

        t1.Start();
        t2.Start();

        Stopwatch stopwatch = Stopwatch.StartNew();
        startSignal.Set();

        Task.WaitAll(t1, t2);
        stopwatch.Stop();

        double totalElapsedNs = stopwatch.Elapsed.TotalMilliseconds * 1_000_000;
        latency = totalElapsedNs / (2 * iterations);

        return latency;
    }

    static void LatencyTestThread(int core, long startValue, long expectedValue)
    {
        SetThreadAffinity(core);
        startSignal.Wait();

        long current = startValue;

        while (current <= 2 * Iterations)
        {
            // Added workload to try and counteract OoO impacting latency in modern big.LITTLE/clustered designs 
            if (Interlocked.CompareExchange(ref bounceValue, current, expectedValue) == expectedValue)
            {
                for (int i = 0; i < 4; i++)
                {
                    double dummy = current / 2;
                    double dummy2 = current;
                    dummy = Math.Sqrt(i);
                    dummy = dummy2 * Math.Sin(dummy) + Math.Cos(i);
                    dummy += Math.Tan(i);

                    dummy2 += i;
                    dummy2 -= i;
                    dummy2 = (dummy2 + i) - i;

                    if (dummy2 % 2 == 0) dummy += 0.1;
                    else dummy -= 0.1;

                    dummy = Math.Sin(dummy) + Math.Cos(dummy2);
                }

                current += 2;
                expectedValue += 2;
            }
        }

        endSignal1.Set();
        endSignal2.Set();
    }

    static void SetThreadAffinity(int core)
    {
        IntPtr mask = new IntPtr(1 << core);
        SetThreadAffinityMask(GetCurrentThread(), mask);
    }

    static void SaveLatenciesToExcel(double[,] latencies, string fileName, string cpuName)
    {
        int numCores = latencies.GetLength(0);
        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Latencies");

            worksheet.Cells[1, 1, 1, numCores + 1].Merge = true;
            worksheet.Cells[1, 1].Value = cpuName;
            worksheet.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[1, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[1, 1].Style.Font.Bold = true;
            worksheet.Cells[1, 1].Style.Font.Size = 18;
            worksheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[1, 1].Style.Font.Color.SetColor(System.Drawing.Color.White);
            worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);

            worksheet.Cells[2, 1].Value = "(ns)";
            worksheet.Cells[2, 1].Style.Font.Bold = true;
            worksheet.Cells[2, 1].Style.Font.Color.SetColor(System.Drawing.Color.White);
            worksheet.Cells[2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells[2, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[2, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
            for (int i = 0; i < numCores; i++)
            {
                worksheet.Cells[2, i + 2].Value = $"Core {i}";
                worksheet.Cells[2, i + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[2, i + 2].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
                worksheet.Cells[2, i + 2].Style.Font.Color.SetColor(System.Drawing.Color.White);
                worksheet.Cells[2, i + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[2, i + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[2, i + 2].Style.Font.Bold = true;
            }

            for (int i = 0; i < numCores; i++)
            {
                worksheet.Cells[i + 3, 1].Value = $"Core {i}";
                worksheet.Cells[i + 3, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[i + 3, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
                worksheet.Cells[i + 3, 1].Style.Font.Color.SetColor(System.Drawing.Color.White);
                worksheet.Cells[i + 3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[i + 3, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[i + 3, 1].Style.Font.Bold = true;
                for (int j = 0; j < numCores; j++)
                {
                    if (latencies[i, j] == 0.0) worksheet.Cells[i + 3, j + 2].Value = "X";
                    else worksheet.Cells[i + 3, j + 2].Value = latencies[i, j];
                    worksheet.Cells[i + 3, j + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[i + 3, j + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }
            }

            var range = worksheet.Cells[3, 2, numCores + 2, numCores + 1];
            var rule = range.ConditionalFormatting.AddThreeColorScale();
            rule.LowValue.Color = System.Drawing.ColorTranslator.FromHtml("#63be7b");
            rule.MiddleValue.Color = System.Drawing.ColorTranslator.FromHtml("#ffe984");
            rule.HighValue.Color = System.Drawing.ColorTranslator.FromHtml("#f96a6c");

            var file = new FileInfo(fileName);
            package.SaveAs(file);
        }
    }
}