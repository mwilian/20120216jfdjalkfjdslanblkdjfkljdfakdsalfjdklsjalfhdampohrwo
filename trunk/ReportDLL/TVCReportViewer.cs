using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using FlexCel.Render;
using System.Threading;
using System.IO;
using FlexCel.Pdf;
using System.Drawing.Printing;
using FlexCel.Core;

using System.Diagnostics;
using System.Drawing.Drawing2D;

namespace ReportDLL
{
    public partial class TVCReportViewer : UserControl
    {
        public string _sErr = "";

        public ExcelFile ReportSource
        {
            get { return flexCelImgExport1.Workbook; }
            set
            {
                flexCelImgExport1.Workbook = value;
                try
                {
                    if (value != null)
                    {
                        cbAntiAlias.SelectedIndex = 0;
                        flexCelPreview1.InvalidatePreview();
                    }
                }
                catch { _sErr = "Report Source is not valid!"; }
            }
        }
        FlexCelPrintDocument flexCelPrintDocument1 = new FlexCelPrintDocument();
        public TVCReportViewer()
        {
            InitializeComponent();
        }

        private void btnFirst_Click(object sender, System.EventArgs e)
        {
            flexCelPreview1.StartPage = 1;
        }

        private void btnPrev_Click(object sender, System.EventArgs e)
        {
            flexCelPreview1.StartPage--;
        }

        private void btnNext_Click(object sender, System.EventArgs e)
        {
            flexCelPreview1.StartPage++;
        }

        private void btnLast_Click(object sender, System.EventArgs e)
        {
            flexCelPreview1.StartPage = flexCelPreview1.TotalPages;
        }

        private void btnZoomOut_Click(object sender, System.EventArgs e)
        {
            flexCelPreview1.Zoom -= 0.1f;
        }

        private void btnZoomIn_Click(object sender, System.EventArgs e)
        {
            flexCelPreview1.Zoom += 0.1f;
        }

        private void btnPdf_Click(object sender, System.EventArgs e)
        {
            if (flexCelImgExport1.Workbook == null)
            {
                MessageBox.Show("There is no open file");
                return;
            }
            if (PdfSaveFileDialog.ShowDialog() != DialogResult.OK) return;

            using (FlexCelPdfExport PdfExport = new FlexCelPdfExport(flexCelImgExport1.Workbook, true))
            {
                if (!DoExportToPdf(PdfExport)) return;
            }

            if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            Process.Start(PdfSaveFileDialog.FileName);
        }
        private bool DoExportToPdf(FlexCelPdfExport PdfExport)
        {
            PdfThread MyPdfThread = new PdfThread(PdfExport, PdfSaveFileDialog.FileName, true);
            Thread PdfExportThread = new Thread(new ThreadStart(MyPdfThread.ExportToPdf));
            PdfExportThread.Start();
            using (PdfProgressDialog Pg = new PdfProgressDialog())
            {
                Pg.ShowProgress(PdfExportThread, PdfExport);
                if (Pg.DialogResult != DialogResult.OK)
                {
                    PdfExport.Cancel();
                    PdfExportThread.Join(); //We could just leave the thread running until it dies, but there are 2 reasons for waiting until it finishes:
                    //1) We could dispose it before it ends. This is workaroundable.
                    //2) We might change its workbook object before it ends (by loading other file). This will surely bring issues.
                    return false;
                }

                if (MyPdfThread != null && MyPdfThread.MainException != null)
                {
                    throw MyPdfThread.MainException;
                }
            }
            return true;
        }
        private void btnRecalc_Click(object sender, System.EventArgs e)
        {
            if (flexCelImgExport1.Workbook == null)
            {
                MessageBox.Show("Please open a file before recalculating.");
                return;
            }
            flexCelImgExport1.Workbook.Recalc(true);
            flexCelPreview1.InvalidatePreview();

        }
        private void cbAntiAlias_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            switch (cbAntiAlias.SelectedIndex)
            {
                case 1:
                    flexCelPreview1.SmoothingMode = SmoothingMode.HighQuality;
                    break;
                case 2:
                    flexCelPreview1.SmoothingMode = SmoothingMode.HighSpeed;
                    break;
                default:
                    flexCelPreview1.SmoothingMode = SmoothingMode.AntiAlias;
                    break;
            }

            if (flexCelImgExport1.Workbook != null) flexCelPreview1.InvalidatePreview();
        }
        #region PdfThread
        class PdfThread
        {
            private FlexCelPdfExport PdfExport;
            private string FileName;
            private bool AllVisibleSheets;
            private Exception FMainException;

            internal PdfThread(FlexCelPdfExport aPdfExport, string aFileName, bool aAllVisibleSheets)
            {
                PdfExport = aPdfExport;
                FileName = aFileName;
                AllVisibleSheets = aAllVisibleSheets;
            }

            internal void ExportToPdf()
            {
                try
                {
                    if (AllVisibleSheets)
                    {
                        try
                        {
                            using (FileStream f = new FileStream(FileName, FileMode.Create, FileAccess.Write))
                            {
                                PdfExport.BeginExport(f);
                                PdfExport.PageLayout = TPageLayout.Outlines;
                                PdfExport.ExportAllVisibleSheets(false, System.IO.Path.GetFileNameWithoutExtension(FileName));
                                PdfExport.EndExport();
                            }
                        }
                        catch
                        {
                            try
                            {
                                File.Delete(FileName);
                            }
                            catch
                            {
                                //Not here.
                            }
                            throw;
                        }
                    }
                    else
                    {
                        PdfExport.PageLayout = TPageLayout.None;
                        PdfExport.Export(FileName);
                    }
                }
                catch (Exception ex)
                {
                    FMainException = ex;
                }
            }

            internal Exception MainException
            {
                get
                {
                    return FMainException;
                }
            }
        }
        #endregion
        private void edPage_Leave(object sender, System.EventArgs e)
        {
            ChangePages();
        }

        private void edPage_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
                ChangePages();
            if (e.KeyChar == (char)27)
                UpdatePages();
        }

        private void UpdatePages()
        {
            edPage.Text = String.Format("{0} of {1}", flexCelPreview1.StartPage, flexCelPreview1.TotalPages);
        }
        private void flexCelPreview1_StartPageChanged(object sender, System.EventArgs e)
        {

        }

        private void ChangePages()
        {
            string s = edPage.Text.Trim();
            int pos = 0;
            while (pos < s.Length && s[pos] >= '0' && s[pos] <= '9') pos++;
            if (pos > 0)
            {
                int page = flexCelPreview1.StartPage;
                try
                {
                    page = Convert.ToInt32(s.Substring(0, pos));
                }
                catch (Exception ex)
                {
                    _sErr = ex.Message;
                }

                flexCelPreview1.StartPage = page;
            }
            UpdatePages();
        }

        private void flexCelPreview1_ZoomChanged(object sender, System.EventArgs e)
        {

        }

        private void UpdateZoom()
        {
            edZoom.Text = String.Format("{0}%", (int)Math.Round(flexCelPreview1.Zoom * 100));
        }

        private void ChangeZoom()
        {
            string s = edZoom.Text.Trim();
            int pos = 0;
            while (pos < s.Length && s[pos] >= '0' && s[pos] <= '9') pos++;
            if (pos > 0)
            {
                int zoom = (int)Math.Round(flexCelPreview1.Zoom * 100);
                try
                {
                    zoom = Convert.ToInt32(s.Substring(0, pos));
                }
                catch (Exception ex)
                {
                    _sErr = ex.Message;
                }

                flexCelPreview1.Zoom = zoom / 100f;
            }
            UpdateZoom();
        }

        private void edZoom_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
                ChangeZoom();
            if (e.KeyChar == (char)27)
                UpdateZoom();
        }

        private void edZoom_Enter(object sender, System.EventArgs e)
        {
            ChangeZoom();
        }

        private bool LoadPreferences()
        {
            try
            {
                flexCelPrintDocument1.Workbook = flexCelImgExport1.Workbook;
                ExcelFile Xls = flexCelPrintDocument1.Workbook;
                Xls.PrintHeadings = false;
                Xls.PrintGridLines = false;
                //Xls.PrintPaperSize = TPaperSize.

                flexCelPrintDocument1.DefaultPageSettings.PaperSize = new PaperSize(Xls.PrintPaperDimensions.PaperName, Convert.ToInt32(Xls.PrintPaperDimensions.Width), Convert.ToInt32(Xls.PrintPaperDimensions.Height));
                //flexCelPrintDocument1.PrintPa
                flexCelPrintDocument1.DefaultPageSettings.Landscape = (Xls.PrintOptions & TPrintOptions.Orientation) == 0;
                return true;
            }
            catch (Exception ex) { throw ex; }
        }
        private bool DoSetup(PrintDocument doc)
        {
            printDialog1.Document = doc;
            printDialog1.PrinterSettings.DefaultPageSettings.PaperSize = doc.DefaultPageSettings.PaperSize;
            //printDialog1.PrinterSettings.
            bool Result = printDialog1.ShowDialog() == DialogResult.OK;
            //printDialog1.PrinterSettings.PaperSizes = doc.PrinterSettings.pga
            //Landscape.Checked = flexCelPrintDocument1.DefaultPageSettings.Landscape;
            return Result;
        }
        private void btnPrintReport_Click(object sender, EventArgs e)
        {
            try
            {
                if (!LoadPreferences()) return;
                if (!DoSetup(flexCelPrintDocument1)) return;
                flexCelPrintDocument1.Print();
            }
            catch (Exception ex)
            {
                _sErr = ex.Message;
            }
            //try
            //{
            //    frmFlexPrintPreview frmPrint = new frmFlexPrintPreview(flexCelImgExport1.Workbook);
            //    frmPrint.Show(this);
            //}
            //catch (Exception ex)
            //{
            //    lb_Err.Text = ex.Message;
            //}
        }
    }


}
