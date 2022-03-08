using ExportSample.Service;
using System;
using System.Web.UI;

namespace ExportSample
{
    public partial class _Default : Page
    {
        #region ===[ Page Events ]=================================================================

        protected void Page_Load(object sender, EventArgs e)
        {

        }

        #endregion

        #region ===[ Control Events ]==============================================================

        protected void BtnExportReportClick(object sender, EventArgs e)
        {
            //Export file.
            switch (ddlReportType.SelectedValue)
            {
                case "1":
                    ExportService.ExportExcelFile();
                    break;
                case "2":
                    ExportService.ExportCsvFile();
                    break;
                case "3":
                    ExportService.ExportTextFile();
                    break;
                case "4":
                    ExportService.ExportPdfFile();
                    break;
                case "5":
                    ExportService.ExportHtmlFile();
                    break;
            }
        }

        #endregion
    }
}