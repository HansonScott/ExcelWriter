#region Header
/*
-----------------------------------------------------------------------------
Copyright (c) 2003 ClaimRemedi, Inc. All Rights Reserved.

PROPRIETARY NOTICE: This software has been provided pursuant to a
License Agreement that contains restrictions on its use. This software
contains valuable trade secrets and proprietary information of
ClaimRemedi, Inc and is protected by Federal copyright law.  It may not 
be copied or distributed in any form or medium, disclosed to third parties, 
or used in any manner that is not provided for in said License Agreement, 
except with the prior written authorization of ClaimRemedi, Inc.

-----------------------------------------------------------------------------
$Log$
Revision 4598  2013/06/19 09:34:31  ddaine
  Set the file extension to xlsx.
  Add SetDataType to correctly set numeric values.

Revision 4557  2013/06/03 14:24:45  ddaine
  Implement GridLines and FitToPage.
  Fix AlignColumnsByDataType for short worksheets.

Revision 4555  2013/05/30 12:51:52  shanson
  lots of bug fixes, removed warnings, improved data parsing for data types.

Revision 4550  2013/05/30 10:08:32  shanson
  changed the repeat header rows and columbs per page

Revision 4541  2013/05/30 08:43:05  shanson
  re-implemented many of the formatting functions.

Revision 4133  2013/05/30 07:50:59  shanson
  updated to .Net 4.0
  removed reference to MS Excel
  recreated very basics, all formatting commented out.

Revision 3165  2013/04/05 08:13:06  shanson
  added try-catch block for saving file, still throws error, but it will quit the application.

Revision 2376  2012/12/17 15:43:03  shanson
  added try-catch around margin setting. - threw error when using it one time, not sure why.

Revision 2374  2012/09/28 10:01:31  shanson
  fixed the set margins feature to understand and translate measurements in inches.

Revision 2369  2012/09/27 21:00:13  shanson
  added highlighting of arbitrary rows.

Revision 2265  2012/09/27 20:40:01  shanson
  added set margins and set header/footer image

Revision 2254  2012/09/20 17:11:56  shanson
  bug fix - catch a bad sheet name - note there are a few more characters that need to be handled, but the list is started.

Revision 2252  2012/09/19 17:03:15  shanson
  cleaned up code, moved some defaults to parent class.

Revision 2219  2012/09/19 16:56:02  shanson
  bug fix when using custom sheet name

Revision 2218  2012/09/17 16:12:40  shanson
  continued developing more billing report customizations.  Prepping for many more reports to allow export to Excel.

Revision 1  2012/09/17 13:22:56  shanson
  created and generalized Excel writer to its own dll


-----------------------------------------------------------------------------
*/
#endregion
using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ClaimRemedi.ExcelWriter
{
    public class Writer
    {
        // a local enum to avoid needing other namespaces
        public enum SortOrder
        {
            Ascending,
            Descending
        }
        public class Keywords
        {
            public static string CurrentPage = "&P";
            public static string TotalPages = "&N";
            public static string CurrentDate = "&D";
            public static string CurrentTime = "&T";
            public static string FileName = "&F";
            public static string Picture = "&G";
            /*
                &A  Represents worksheet name  
                &"<FontName>"  Represents font name. For example: &"Arial"  
                &"<FontName>, <FontStyle>"  Represents font name with style. For example: &"Arial,Bold"  
                &<FontSize>  Represents font size. For example: &14 . If this command is followed by a plain number to be printed in the header, it will be separated from the font size with a space character. For example: &14 123  

             */

        }

        #region Data Members
        string DestinationFilePath = string.Empty;
        System.Data.DataTable Data = null;

        private OfficeOpenXml.ExcelPackage app = null;
        private OfficeOpenXml.ExcelWorksheet worksheet = null;

        public bool HaveHeaderRow = true;
        #endregion

        #region Constructor and Setup
        public Writer(): this(null, null){}
        public Writer(System.Data.DataTable data): this(data, null){}
        public Writer(System.Data.DataTable data, string path): this(data, path, "Sheet1"){}
        public Writer(System.Data.DataTable data, string path, string SheetName)
        {
            // The new version writes xlsx files. Make sure the file extension is correct.
            if (path.EndsWith(".xls"))
            {
                path = path + "x";
            }
            else if (!path.EndsWith(".xlsx"))
            {
                path = path + ".xlsx";
            }
            this.DestinationFilePath = path;

            this.Data = data;

            if (data != null && !String.IsNullOrEmpty(path))
            {
                InitFileObjects(SheetName);

                SetDefaults();
            }
        }

        private void SetDefaults()
        {
            SetFirstRowFromColumns();
            FillCellsFromDataTable(HaveHeaderRow);
            AutoFitColumnWidths();
            SetMargins(0.3);
        }

        private void InitFileObjects(string SheetName)
        {
            SheetName = CleanSheetName(SheetName);
            if (File.Exists(DestinationFilePath))
            {
                try
                {
                    File.Delete(DestinationFilePath);
                }
                catch { }
            }
            app = new ExcelPackage(new FileInfo(DestinationFilePath));
            worksheet = app.Workbook.Worksheets.Add(SheetName);
        }
        private string CleanSheetName(string SheetName)
        {
            string result = SheetName;
            if (result.Length > 30)
            {
                result = result.Substring(0, 30);
            }
            if (result.Contains(@"/"))
            {
                result.Replace("/", " ");
            }
            if (result.Contains(@"\"))
            {
                result.Replace(@"\", " ");
            }

            return result;
        }
        #endregion

        #region Public Methods
        /// <summary>
        ///  this feature not re-implemented yet.
        /// </summary>
        public void PrintGridLines()
        {
            //if (worksheet == null) { return; }

            //worksheet.PageSetup.PrintGridlines = true;
            if (worksheet != null)
                worksheet.PrinterSettings.ShowGridLines = true;
        }
        public void SetHeadRowRepeatPerPage()
        {
            if (worksheet == null || Data == null) { return; }
            worksheet.PrinterSettings.RepeatRows = new ExcelAddress("1:1");
        }
        public void SetHeaderColumnRepeatPerPage()
        {
            if (worksheet == null || Data == null) { return; }
            worksheet.PrinterSettings.RepeatColumns = new ExcelAddress("A:A");
        }
        public void FreezeHeaderRow()
        {
            if (worksheet == null) { return; }

            //worksheet.Application.ActiveWindow.SplitRow = 1;
            //worksheet.Application.ActiveWindow.FreezePanes = true;
            //worksheet.View.FreezePanes(1, 0);
        }

        /// <summary>
        ///  this feature not re-implemented yet.
        /// </summary>
        public void SetDefaultColumnWidth(int w)
        {
            //if (worksheet == null) { return; }

            //worksheet.StandardWidth = w;
        }

        /// <summary>
        ///  this feature not re-implemented yet.
        /// </summary>
        /// <param name="ColumnIndex"></param>
        /// <param name="SO"></param>
        public void Sort(string ColumnIndex, SortOrder SO)
        {
            if (worksheet == null) { return; }

            switch (SO)
            {
                //case SortOrder.Ascending:
                //    worksheet.Cells.Sort(worksheet.Columns[ColumnIndex, Type.Missing],
                //                        XlSortOrder.xlAscending, Type.Missing, Type.Missing,
                //                        XlSortOrder.xlAscending, Type.Missing,
                //                        XlSortOrder.xlAscending,
                //                        XlYesNoGuess.xlYes, Type.Missing, Type.Missing,
                //                        XlSortOrientation.xlSortColumns, XlSortMethod.xlPinYin,
                //                        XlSortDataOption.xlSortNormal,
                //                        XlSortDataOption.xlSortNormal,
                //                        XlSortDataOption.xlSortNormal);
                //    break;
                //case SortOrder.Descending:
                //    worksheet.Cells.Sort(worksheet.Columns[ColumnIndex, Type.Missing],
                //                        XlSortOrder.xlDescending, Type.Missing, Type.Missing,
                //                        XlSortOrder.xlAscending, Type.Missing,
                //                        XlSortOrder.xlAscending,
                //                        XlYesNoGuess.xlYes, Type.Missing, Type.Missing,
                //                        XlSortOrientation.xlSortColumns, XlSortMethod.xlPinYin,
                //                        XlSortDataOption.xlSortNormal,
                //                        XlSortDataOption.xlSortNormal,
                //                        XlSortDataOption.xlSortNormal);
                //    break;
                default:
                    break;
            }
        }

        /// <summary>
        ///  this feature not re-implemented yet.
        /// </summary>
        public void SetGridLinesForCellsWithData(int BorderWidth)
        {
            //if (worksheet == null) { return; }

            //for (int i = 0; i <= Data.Rows.Count; i++)
            //{
            //    for (int j = 0; j < Data.Columns.Count; j++)
            //    {
            //        ((Range)worksheet.Cells[i + 1, j + 1]).Borders.LineStyle = XlLineStyle.xlContinuous;
            //        ((Range)worksheet.Cells[i + 1, j + 1]).Borders.Weight = BorderWidth;
            //    }
            //}
        }

        public void SetHeaderRowBold()
        {
            SetRowBold(1);
        }
        public void SetLastRowBold()
        {
            if (Data == null) { return; }
            int offset = 0;
            if (HaveHeaderRow) { offset = 1; }

            SetRowBold(Data.Rows.Count + offset);
        }
        public void SetRowBold(int RowIndex)
        {
            if (worksheet == null) { return; }

            ExcelRange range = worksheet.Cells[RowIndex, 1, RowIndex, Data.Columns.Count];
            range.Style.Font.Bold = true;
        }
        public void ColorHeaderRow(Color C)
        {
            ColorRow(0, C);
        }
        public void ColorLastRow(Color C)
        {
            if (Data == null) { return; }
            ColorRow(Data.Rows.Count, C);
        }
        public void ColorRow(int RowIndex, Color C)
        {
            if (worksheet == null) { return; }

            int offset = 0;
            if (HaveHeaderRow) { offset = 1; }

            ExcelRange range = worksheet.Cells[RowIndex + offset, 1, RowIndex + offset, Data.Columns.Count];

            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(C);


            // FONT COLOR
            //range.Style.Font.Color.SetColor(C);
        }

        public void SetColumnWidth(string ColumnIndex, int Width)
        {
            if (worksheet == null) { return; }
            int nColumnIndex = 1;
            if (Int32.TryParse(ColumnIndex, out nColumnIndex))
            {
                worksheet.Column(nColumnIndex).Width = Width;
            }
        }

        public void SetColumnFormat(int ColumnIndex, string Format)
        {
            if(worksheet == null){return;}
            worksheet.Column(ColumnIndex).Style.Numberformat.Format = Format;
        }
        private void SetColumnAlignment(int ColumnIndex, ExcelHorizontalAlignment a)
        {
            if (worksheet == null) { return; }
            worksheet.Column(ColumnIndex).Style.HorizontalAlignment = a;
        }

        private void SetFormat(ExcelRange r, string f)
        {
            r.Style.Numberformat.Format = f;
        }
        public void SetCenterHeader(string HeaderText)
        {
            if (worksheet == null) { return; }
            worksheet.HeaderFooter.FirstHeader.CenteredText = HeaderText;
            worksheet.HeaderFooter.EvenHeader.CenteredText = HeaderText;
            worksheet.HeaderFooter.OddHeader.CenteredText = HeaderText;
        }
        public void SetCenterFooter(string FooterText)
        {
            if (worksheet == null) { return; }
            worksheet.HeaderFooter.FirstFooter.CenteredText = FooterText;
            worksheet.HeaderFooter.EvenFooter.CenteredText = FooterText;
            worksheet.HeaderFooter.OddFooter.CenteredText = FooterText;
        }
        /// <summary>
        ///  this feature not re-implemented yet.
        /// </summary>
        public void SetCenterHeaderPic(string PicURL)
        {
            //if (worksheet == null) { return; }
            //worksheet.PageSetup.CenterHeaderPicture.Filename = PicURL;
            //SetCenterHeader(Keywords.Picture);
        }
        /// <summary>
        ///  this feature not re-implemented yet.
        /// </summary>
        public void SetCenterFooterPic(string PicURL)
        {
            //if (worksheet == null) { return; }
            //worksheet.PageSetup.CenterFooterPicture.Filename = PicURL;
            //SetCenterFooter(Keywords.Picture);
        }

        /// <summary>
        ///  this feature not re-implemented yet.
        /// </summary>
        public void FitToPagesWide(int p)
        {
            //worksheet.PageSetup.Zoom = false;
            //worksheet.PageSetup.FitToPagesWide = p;
            //worksheet.PageSetup.FitToPagesTall = 999;
            worksheet.PrinterSettings.FitToPage = true;
            worksheet.PrinterSettings.FitToWidth = 1;
            worksheet.PrinterSettings.FitToHeight = 999;
        }
        public void AutoFitColumnWidths()
        {
            if (worksheet == null) { return; }
            worksheet.Cells.AutoFitColumns(0);
        }
        /// <summary>
        /// size in inches
        /// </summary>
        public void SetMargins(double All)
        {
            SetMargins(All * 2, All, All * 2, All, All, All);
        }
        /// <summary>
        /// size in inches
        /// </summary>
        public void SetMargins(double top, double header, double bottom, double footer, double left, double right)
        {
            try
            {
                app.Workbook.Worksheets[0].PrinterSettings.HeaderMargin = (decimal)header;
                app.Workbook.Worksheets[0].PrinterSettings.TopMargin = (decimal)top;
                app.Workbook.Worksheets[0].PrinterSettings.FooterMargin = (decimal)footer;
                app.Workbook.Worksheets[0].PrinterSettings.BottomMargin = (decimal)bottom;
                app.Workbook.Worksheets[0].PrinterSettings.LeftMargin = (decimal)left;
                app.Workbook.Worksheets[0].PrinterSettings.RightMargin = (decimal)right;
            }
            catch{}
        }
        public void FillCellsFromDataTable(bool HaveHeaderRow)
        {
            if (worksheet == null || Data == null) { return; }

            worksheet.Cells.LoadFromDataTable(Data, true);

            AlignColumnsByDataType();

            SetDataType();
        }

        private void AlignColumnsByDataType()
        {
            // make sure each column has the dataType properly classified
            decimal d;
            int i;
            for (int j = 0; j < Data.Columns.Count; j++)
            {
                string testValue = string.Empty;
                if (Data.Rows.Count > 2) { testValue = Data.Rows[1][j].ToString(); }
                else if (Data.Rows.Count == 1 ) { testValue = Data.Rows[0][j].ToString(); }
                else { return; } // no row, don't bother aligning them.

                int nRowsToCheck = Math.Min(7, Data.Rows.Count);
                int nRow = 2;
                for (int k = 2; k < nRowsToCheck; k++)
                {
                    if (!String.IsNullOrEmpty(Data.Rows[k][j].ToString()))
                    {
                        testValue = Data.Rows[k][j].ToString();
                        nRow = k;
                        break;
                    }
                }

                if (Int32.TryParse(testValue, out i))
                {
                    SetColumnFormat(j + 1, "0");
                    SetColumnAlignment(j + 1, ExcelHorizontalAlignment.Right);
                }
                else if (Decimal.TryParse(testValue, out d))
                {
                    SetColumnFormat(j + 1, "#,###.##");
                    SetColumnAlignment(j + 1, ExcelHorizontalAlignment.Right);
                }
                else
                {
                    SetColumnFormat(j + 1, "@");
                    SetColumnAlignment(j + 1, ExcelHorizontalAlignment.Left);
                }
            }
        }
        private void SetDataType()
        {
            for (int i = 0; i < Data.Rows.Count; i++)
            {
                for (int j = 0; j < Data.Columns.Count; j++)
                {
                    int iVal;
                    if (Int32.TryParse(Data.Rows[i][j].ToString(), out iVal))
                    {
                        worksheet.Cells[i + 2, j + 1].Value = iVal;
                    }

                }
            }
        }
        public void SetFirstRowFromColumns()
        {
            if (worksheet == null || Data == null) { return; }

            HaveHeaderRow = true;

            for (int i = 1; i < Data.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i].Value = Data.Columns[i - 1].ColumnName;
            }
        }
        public void SaveAndCloseFile()
        {
            if (app == null) { return; }
            if (worksheet == null) { return; }

            // save the application
            try
            {
                app.Save();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        #endregion
    }
}
