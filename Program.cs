using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using Mono.Cecil;
using Mono.Cecil.Cil;

namespace DllAnalyzer
{
    public class Program
    {
        static void Main(string[] args)
        {
            string dllPath = @"E:\o\vscode-dotnet\dhlibraries\DH.BLLCLS.dll";

            if (!File.Exists(dllPath))
            {
                Console.WriteLine($"Lỗi: Không tìm thấy file {dllPath}");
                return;
            }

            try
            {
                var analyzer = new DllAnalyzer();
                var result = analyzer.AnalyzeDll(dllPath);

                string outputPath = dllPath + ".json";
                analyzer.SaveToJson(result, outputPath);

                string summaryPath = dllPath + ".summary.json";
                analyzer.SaveSummaryToJson(result, summaryPath);

                string externalRefsPath = dllPath + ".external.json";
                analyzer.SaveExternalReferencesToJson(result, externalRefsPath);

                Console.WriteLine($"Phân tích thành công!");
                Console.WriteLine($"Kết quả đầy đủ đã được lưu vào: {outputPath}");
                Console.WriteLine($"Kết quả rút gọn đã được lưu vào: {summaryPath}");
                Console.WriteLine($"External references đã được lưu vào: {externalRefsPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi phân tích DLL: {ex.Message}");
                Console.WriteLine($"Chi tiết: {ex}");
            }
        }
    }
}