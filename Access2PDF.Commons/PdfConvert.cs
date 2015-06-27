using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using MsAccess = Microsoft.Office.Interop.Access;

namespace Access2PDF.Commons
{
    /// <summary>
    /// Auxilia na conversão de relatórios do Microsoft Acess para PDF
    /// </summary>
    public static class PdfConvert
    {
        /// <summary>
        /// Transforma o relatorio do Microsoft Access em arquivo pdf
        /// </summary>
        /// <param name="reportName"></param>
        /// <param name="msAccess"></param>
        /// <param name="outputPdf"></param>
        /// <param name="filtersReport"></param>
        public static void GenerateFile(string reportName, string msAccess, string outputPdf, List<string> filtersReport)
        {
            var app = new MsAccess.Application();

            try
            {
                if (string.IsNullOrWhiteSpace(reportName))
                    throw new ArgumentException("Nome do relatório inválido", "reportName");

                if (string.IsNullOrWhiteSpace(msAccess) || !File.Exists(msAccess))
                    throw new ArgumentException("Arquivo do Microsoft Access Inválido ou não encontrado", "msAccess");

                if (string.IsNullOrWhiteSpace(outputPdf))
                    throw new ArgumentException("Arquivo de destino inválido", "msAccess");
                
                app.OpenCurrentDatabase(msAccess, false, "");
                app.Visible = false;

                if (filtersReport != null)
                {
                    filtersReport.ForEach(filterReport =>
                    {
                        app.DoCmd.OpenReport(
                            reportName,
                            MsAccess.AcView.acViewReport,
                            null,
                            filterReport,
                            MsAccess.AcWindowMode.acHidden,
                            null
                        );

                        ExportToPdf(app, reportName, outputPdf);
                    });
                    
                }
                else
                {
                    ExportToPdf(app, reportName, outputPdf);
                }
                
                app.CloseCurrentDatabase();

                app.DoCmd.Close(
                        MsAccess.AcObjectType.acReport,
                        reportName,
                        MsAccess.AcCloseSave.acSaveNo);

            }
            finally
            {
                app.Quit(MsAccess.AcQuitOption.acQuitSaveNone);
                Marshal.FinalReleaseComObject(app);

                app = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static void ExportToPdf(MsAccess.Application app, string reportName, string outputPdf)
        {
            app.DoCmd.OutputTo(
                    MsAccess.AcOutputObjectType.acOutputReport,
                    reportName,
                    "PDF Format (*.pdf)",
                    outputPdf,
                    false,
                    null,
                    null,
                    MsAccess.AcExportQuality.acExportQualityPrint);
        }
    }
}
