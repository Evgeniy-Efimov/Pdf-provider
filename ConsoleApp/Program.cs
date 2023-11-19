using System.Collections.Generic;
using System;
using PdfProvider;
using System.IO;

namespace ConsoleApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var document = new PdfProvider.PdfProvider("New pdf");

            document.WriteTextLine("", CustomFonts.Main, Aligment.Left);
            document.WriteTextLine("", CustomFonts.Main, Aligment.Left);
            document.WriteTextLine("", CustomFonts.Main, Aligment.Left);
            document.WriteTextLine($"Generated: {DateTime.Now:dd.MM.yyyy HH:mm}", CustomFonts.Main, Aligment.Left);

            var tableContentMargin = new Margin() { MarginTop = 7, MarginRight = 5, MarginLeft = 5, MarginBottom = 0 };
            document.DrawTableRow("Sample table", tableContentMargin, CustomFonts.Header, Aligment.Left);

            var someData = new Dictionary<string, string>();
            var random = new Random();

            for (int i = 1; i <= 10; i++)
            {
                someData.Add($"Random number {i}", random.Next(1, 100).ToString());
            }

            foreach (var key in someData.Keys)
            {
                document.DrawTableRow(2, new List<int>() { 40, 60 },
                    new List<string>() { key, someData[key] },
                    tableContentMargin, CustomFonts.Main, Aligment.Left);
            }

            File.WriteAllBytes($"Pdf sample {DateTime.Now:dd.MM.yyyy HH-mm}.pdf", document.GetDocumentData());
            Console.WriteLine("Done.");
            Console.ReadLine();
        }
    }
}
