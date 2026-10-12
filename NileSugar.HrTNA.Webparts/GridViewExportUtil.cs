using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace NileSugar.HrTNA.Webparts
{
    public static class GridViewExportUtil
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="gv"></param>
        public static void Export(string fileName, GridView gv)
        {
            //HttpContext.Current.Response.Clear();
            //HttpContext.Current.Response.AddHeader(PrepareControlForExport
            //    "content-disposition", string.Format("attachment; filename={0}", fileName));
            //HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";

            using (StringWriter sw = new StringWriter())
            {
                using (HtmlTextWriter htw = new HtmlTextWriter(sw))
                {
                    //  Create a table to contain the grid
                    Table table = new Table();

                    //  include the gridline settings
                    table.GridLines = gv.GridLines;

                    //  add the header row to the table
                    if (gv.HeaderRow != null)
                    {
                        GridViewExportUtil.PrepareControlForExport(gv.HeaderRow);
                        table.Rows.Add(gv.HeaderRow);
                    }

                    //  add each of the data rows to the table
                    foreach (GridViewRow row in gv.Rows)
                    {
                        GridViewExportUtil.PrepareControlForExport(row);
                        table.Rows.Add(row);
                    }

                    //  add the footer row to the table
                    if (gv.FooterRow != null)
                    {
                        GridViewExportUtil.PrepareControlForExport(gv.FooterRow);
                        table.Rows.Add(gv.FooterRow);
                    }

                    //  render the table into the htmlwriter
                    table.RenderControl(htw);

                    //  render the htmlwriter into the response
                    HttpContext.Current.Response.Write(sw.ToString());
                    HttpContext.Current.Response.End();
                }
            }
        }


        public static Table ExportGridtoTable(GridView gv)
        {
            Table table = new Table();
            using (StringWriter sw = new StringWriter())
            {
                using (HtmlTextWriter htw = new HtmlTextWriter(sw))
                {
                    //  Create a table to contain the grid  


                    //  include the gridline settings
                    table.GridLines = gv.GridLines;

                    //  add the header row to the table
                    if (gv.HeaderRow != null)
                    {
                        GridViewExportUtil.PrepareControlForExport(gv.HeaderRow);
                        table.Rows.Add(gv.HeaderRow);
                   
                    }

                    //  add each of the data rows to the table
                   
                    foreach (GridViewRow row in gv.Rows)
                    {
                        row.Attributes.CssStyle["text-align"] = "center";
                       table.Rows.Add(row);
                    
                    }

                   
                    
                    //  add the footer row to the table
                    if (gv.FooterRow != null)
                    {
                        GridViewExportUtil.PrepareControlForExport(gv.FooterRow);
                        table.Rows.Add(gv.FooterRow);
                    }
                   // table.Style(gv.Style);

                }
                return table;
            }
        }

        public static Table ExportGridtoTable(GridView gv,string status)
        {
            Table table = new Table();
            using (StringWriter sw = new StringWriter())
            {
                using (HtmlTextWriter htw = new HtmlTextWriter(sw))
                {
                    //  Create a table to contain the grid  


                    //  include the gridline settings
                    table.GridLines = gv.GridLines;

                    //  add the header row to the table
                    if (gv.HeaderRow != null)
                    {
                        GridViewExportUtil.PrepareControlForExport(gv.HeaderRow);
                        table.Rows.Add(gv.HeaderRow);
                    }

                    //  add each of the data rows to the table

                    foreach (GridViewRow row in gv.Rows)
                    {
                        TableRow tb_row = new TableRow();
                        for (int i = 0; i < gv.Columns.Count; i++)
                        {
                            TableCell tb_cell = new TableCell();
                            
                                tb_cell.Text = row.Cells[i].Text;
                                tb_row.Cells.Add(tb_cell);

                            
                        }
                        table.Rows.Add(tb_row);

                        /*Mary
                        table.Rows.Add(row);*/

                    }



                    //  add the footer row to the table
                    if (gv.FooterRow != null)
                    {
                        GridViewExportUtil.PrepareControlForExport(gv.FooterRow);
                        table.Rows.Add(gv.FooterRow);
                    }

                }
                return table;
            }
        }
        /// <summary>
        /// Replace any of the contained controls with literals
        /// </summary>
        /// <param name="control"></param>
        public static void PrepareControlForExport(Control control)
        {
            for (int i = 0; i < control.Controls.Count; i++)
            {
                Control current = control.Controls[i];
                if (current is LinkButton)
                {
                    control.Controls.Remove(current);
                    control.Controls.AddAt(i, new LiteralControl((current as LinkButton).Text));
                }
                else if (current is ImageButton)
                {
                    control.Controls.Remove(current);
                    control.Controls.AddAt(i, new LiteralControl((current as ImageButton).AlternateText));
                }
                else if (current is HyperLink)
                {
                    control.Controls.Remove(current);
                    control.Controls.AddAt(i, new LiteralControl((current as HyperLink).Text));
                }
                else if (current is DropDownList)
                {
                    control.Controls.Remove(current);
                    control.Controls.AddAt(i, new LiteralControl((current as DropDownList).SelectedValue));
                }
                else if (current is CheckBox)
                {
                    control.Controls.Remove(current);
                    control.Controls.AddAt(i, new LiteralControl((current as CheckBox).Checked ? "True" : "False"));
                }

                if (current.HasControls())
                {
                    GridViewExportUtil.PrepareControlForExport(current);
                }
            }
        }

        public static string AddExcelStyling()
        {

            StringBuilder sb = new StringBuilder();

            sb.Append("<html xmlns:o='urn:schemas-microsoft-com:office:office'\n" +

            "xmlns:x='urn:schemas-microsoft-com:office:excel'\n" +

            "xmlns='http://www.w3.org/TR/REC-html40'>\n" +

            "<head>\n");

            sb.Append("<style>\n");

            sb.Append("@page");

            sb.Append("mso-page-orientation:landscape;\n");
            sb.Append(@" td { mso-number-format:'\@';} ");
            sb.Append("</style>\n");

            sb.Append("<!--[if gte mso 9]><xml>\n");

            sb.Append("<x:ExcelWorkbook>\n");

            sb.Append("<x:ExcelWorksheets>\n");

            sb.Append("<x:ExcelWorksheet>\n");

            sb.Append("<x:Name>Sheet Name</x:Name>\n");

            sb.Append("<x:WorksheetOptions>\n");

            sb.Append("<x:Print>\n");

            sb.Append("<x:HorizontalResolution>600</x:HorizontalResolution\n");

            sb.Append("<x:VerticalResolution>600</x:VerticalResolution\n");

            sb.Append("</x:Print>\n");

            sb.Append("<x:Selected/>\n");

            sb.Append("<x:DisplayRightToLeft/>\n");



            sb.Append("</x:WorksheetOptions>\n");

            sb.Append("</x:ExcelWorksheet>\n");

            sb.Append("</x:ExcelWorksheets>\n");

            sb.Append("</x:ExcelWorkbook>\n");

            sb.Append("</xml><![endif]-->\n");

            sb.Append("</head>\n");

            sb.Append("<body>\n");

            return sb.ToString();
           


        }

        public static void HideColumn(this Table table, int index)
        {
            foreach (TableRow row in table.Rows)
            {
                if (row.Cells.Count - 1 >= index)
                {
                    row.Cells[index].Visible = false;
                }
            }
        }

        public static void HideColumn(this Table table, string id)
        {
            int index = 0;
            bool columnFound = false;

            if (table.Rows.Count > 1)
            {
                TableHeaderRow headerRow = table.Rows[0] as TableHeaderRow;
                if (headerRow != null)
                {
                    foreach (TableHeaderCell cell in headerRow.Cells)
                    {
                        if (cell.ID.ToLower() == id.ToLower())
                        {
                            columnFound = true;
                            break;
                        }

                        index++;
                    }
                }
            }

            if (columnFound)
                HideColumn(table, index);
        }

    }
}
