using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Access2PDF.Commons;
using System.IO;

namespace Access2Pdf.Test
{
    [TestClass]
    public class TestGeneration
    {
        [TestMethod]
        public void TestPdfGeneration()
        {
            var msAccess = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,  "helloWorld.mdb");
            var outputPdf = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,  "helloWorld.pdf");

            PdfConvert.GenerateFile("Test", msAccess, outputPdf, string.Empty, true);
        }
    }
}
