using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Xobject;
using iText.Kernel.Utils;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExtractPDFPages
{
    class Program
    {
        static void Main(string[] args)
        {
            // Prompt for input PDF path
            Console.Write("Enter the full path to the input PDF file: ");
            string inputPdfPath = Console.ReadLine();

            // Validate the input path
            if (!File.Exists(inputPdfPath))
            {
                Console.WriteLine("The input file does not exist. Please check the path and try again.");
                return;
            }

            // Prompt for output PDF path
            Console.Write("Enter the full path to save the extracted PDF pages (output file): ");
            string outputPdfPath = Console.ReadLine();

            // Validate the output path directory
            string outputDirectory = Path.GetDirectoryName(outputPdfPath);
            if (!Directory.Exists(outputDirectory))
            {
                Console.WriteLine("The output directory does not exist. Please check the path and try again.");
                return;
            }

            // Prompt for the number of search terms
            Console.Write("Enter the number of search terms: ");
            if (!int.TryParse(Console.ReadLine(), out int numSearchTerms) || numSearchTerms <= 0)
            {
                Console.WriteLine("Please enter a valid number of search terms.");
                return;
            }

            // Collect search terms from the user
            List<string> searchTerms = new List<string>();
            for (int i = 0; i < numSearchTerms; i++)
            {
                Console.Write($"Enter search term {i + 1}: ");
                searchTerms.Add(Console.ReadLine());
            }

            // List to store pages with search terms
            List<int> pagesToExtract = new List<int>();

            // Open the PDF document
            using (PdfDocument pdfDoc = new PdfDocument(new PdfReader(inputPdfPath)))
            {
                // Loop through all the pages
                for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)  // PDF pages are 1-indexed
                {
                    string pageText = ExtractTextFromPage(pdfDoc, i);

                    // Search for terms in the page text
                    foreach (var term in searchTerms)
                    {
                        if (pageText.Contains(term, StringComparison.OrdinalIgnoreCase))
                        {
                            Console.WriteLine($"Found '{term}' on page {i}");
                            pagesToExtract.Add(i);  // Add the page to the list if the term is found
                            break;  // Stop checking other terms on this page
                        }
                    }
                }
            }

            // Extract the pages to a new PDF
            if (pagesToExtract.Count > 0)
            {
                ExtractPages(inputPdfPath, outputPdfPath, pagesToExtract);
                Console.WriteLine("Pages containing the search terms have been extracted successfully.");
            }
            else
            {
                Console.WriteLine("No matching terms were found in the document.");
            }
        }

        // Helper method to extract text from a page
        static string ExtractTextFromPage(PdfDocument pdfDoc, int pageNumber)
        {
            var strategy = new LocationTextExtractionStrategy();
            var text = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(pageNumber), strategy);
            return text;
        }

        static void ExtractPages(string sourcePdfPath, string outputPdfPath, List<int> pages)
        {
            // Ensure the output file does not already exist
            if (File.Exists(outputPdfPath))
            {
                File.Delete(outputPdfPath); // or prompt the user to choose a different path
            }

            using (PdfDocument pdfDoc = new PdfDocument(new PdfReader(sourcePdfPath)))
            {
                using (PdfDocument outputPdfDoc = new PdfDocument(new PdfWriter(outputPdfPath)))
                {
                    PdfMerger merger = new PdfMerger(outputPdfDoc);
                    foreach (int pageNum in pages)
                    {
                        merger.Merge(pdfDoc, pageNum, pageNum);
                    }
                }
            }
        }

    }
}
