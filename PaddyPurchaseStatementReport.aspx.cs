using BAL;
using BAL.eKharid;
using BOL.eKharid;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Text;
using System.Globalization;

namespace eKharidNew_UI.MMS
{
    public partial class PaddyPurchaseStatementReport : System.Web.UI.Page
    {
        public string SortDirection
        {
            get { return ViewState["SortDirection"] != null ? ViewState["SortDirection"].ToString() : "ASC"; }
            set { ViewState["SortDirection"] = value; }
        }
        static AuctionManagementBLL _objAuctionBLL = new AuctionManagementBLL();
        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["usr"] == null)
            {
                Response.Redirect("login");
                sessionId.Value = Session["cBy"].ToString();
            }
            else
            {
            }
            if (!Page.IsPostBack)
            {
                CommonFunction.AntiFixationInit();
                CommonFunction.AntiHijackInit();
                Session["LicenceId"] = Session["cBy"].ToString();
               
                DataSet _DataSet = _objAuctionBLL.GetRole(Session["cBy"].ToString());
                string Role = string.Empty;
                if (_DataSet != null)
                {
                    Role = _DataSet.Tables[0].Rows[0][0].ToString();
                }

                if (Role == "SMC")
                {
                    Session["Type"] = "SMC";
                    
                    DataSet _ds = _objAuctionBLL.GetDistrict(Session["cBy"].ToString());
                    string DistrictId = "0";
                    if (_ds != null)
                    {
                        DistrictId = _ds.Tables[0].Rows[0][0].ToString();
                       
                    }
                }

                if (Role == "AdminMB")
                {
                    Session["Type"] = "AMB";
              
                }

                if (Role == "Super Admin")
                {
                    Session["Type"] = "SA";
                    
                    string DistrictId = "0";
                
                }

                if (Role == "AdminHF")
                {
                    Session["Type"] = "HAFED";
                    
                }

                if (Role == "AdminWC")
                {
                    Session["Type"] = "HSWC";
                   
                }

                if (Role == "AdminFCI")
                {
                    Session["Type"] = "FCI";
                   
                }

                if (Role == "AdminFS")
                {
                    Session["Type"] = "FOOD AND SUPPLY";
                    
                }

                if (Role == "DFSC")
                {
                    Session["Type"] = "FOOD AND SUPPLY";
                   
                }
                if (Role == "DMFCI")
                {
                    Session["Type"] = "FCI";
                   
                }
                if (Role == "DMWC")
                {
                    Session["Type"] = "HSWC";
                   
                }
                if (Role == "DMHF")
                {
                    Session["Type"] = "HAFED";
                    
                }
            }
        }
       

        public static DataTable GetReportData(string Type, string Auction_date)
        {
            DataTable oTable = new DataTable();
            try
            {
                DataSet _dsList = null;
                PaddyPurchaseBLL _objBLL = new PaddyPurchaseBLL();
                _dsList = _objBLL.PaddyPurchaseReport(Type, Auction_date);
                if (_dsList.Tables[0].Rows.Count > 0)
                {
                    oTable = _dsList.Tables[0];
                }
                //oTable.Columns.Remove("Agro");
                //oTable.Columns.RemoveAt(9);
                return oTable;
            }
            catch (Exception ex)
            {
                return oTable;
            }
        }

        private string DateFormat()
        {
            string dateString = txtDatepicker.Text.Trim();
            string formattedDateTime = "";
            DateTime parsedDate;
            if (DateTime.TryParseExact(dateString, new string[] {

               "yyyy-MM-dd",      // Format: 2023-09-25
    "dd/MM/yyyy",      // Format: 25/09/2023
    "MM/dd/yyyy",      // Format: 09/25/2023
    "yyyy/MM/dd",      // Format: 2023/09/25
    "dd-MM-yyyy",      // Format: 25-09-2023
    "MM-dd-yyyy",      // Format: 09-25-2023
    "yyyy.MM.dd",      // Format: 2023.09.25
    "dd.MM.yyyy",      // Format: 25.09.2023
    "MM.dd.yyyy",      // Format: 09.25.2023
    "yyyy/MM/dd",      // Format: 2023/09/25
    "dd MMM yyyy",     // Format: 25 Sep 2023
    "dd MMMM yyyy",    // Format: 25 September 2023
    "MMMM dd, yyyy",   // Format: September 25, 2023
    "yyyy/MMM/dd",     // Format: 2023/Sep/25
    "yyyy/MMMM/dd",    // Format: 2023/September/25
    "MMM dd, yyyy",    // Format: Sep 25, 2023
            
            
            
            },
                                       CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
            {
                // Successfully parsed, now format it
                formattedDateTime = parsedDate.ToString("dd/MM/yyyy");
                    
                    //+" " + DateTime.Now .ToString ("hh:mm tt");
              
            }
            else
            {
                formattedDateTime = DateTime.Now.ToString("dd/MM/yyyy");
                    //+ " " + DateTime.Now.ToString("hh:mm tt");
            }
            return formattedDateTime;
        }

        protected void grd_RowDataBound(object sender, GridViewRowEventArgs e)
        {
           


            if (e.Row.RowType == DataControlRowType.Header)
            {
                GridViewRow headerRow1 = new GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert);

                GridViewRow headerRow2 = new GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert);
           
                GridViewRow headerRow3 = new GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert);
                GridViewRow headerRow4 = new GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert);

                TableCell Heading = new TableCell
                {
                    Text = "FOOD, CIVIL SUPPLIES & CONSUMER AFFAIRS DEPARTMENT",
                    ColumnSpan = 14,
                    CssClass = "header-parent",
                    HorizontalAlign = HorizontalAlign.Center ,// Center-align the text,
                    BackColor = System.Drawing.Color.LightGray,
                  

                };
                Heading.Style.Add("font-weight", "bold");
                headerRow1.Cells.Add(Heading);


                
                TableCell Heading1 = new TableCell
                {
                    Text = "PROGRESSIVE ARRIVAL/PROCUREMENT OF PADDY DURING KHARIF 2023-24 AS ON " + DateFormat(),
                    ColumnSpan = 11,
                    CssClass = "header-parent",
                    HorizontalAlign = HorizontalAlign.Center, // Center-align the text
                    BackColor = System.Drawing.Color.AliceBlue

                };
                Heading1.Style.Add("font-weight", "bold");
                headerRow2.Cells.Add(Heading1);

                TableCell Heading2 = new TableCell
                {
                    Text = "(All Fig in MT.)",
                    ColumnSpan = 3,
                    CssClass = "header-parent",
                    HorizontalAlign = HorizontalAlign.Right, // RIGHT-align the text
                    BackColor = System.Drawing.Color.AliceBlue

                };
                Heading1.Style.Add("font-weight", "bold");
                headerRow2.Cells.Add(Heading2);


                // Placeholder cells for Sr No and Name of District
                TableCell srHeader = new TableCell
                {
                    Text = "Sr No.",

                    CssClass = "header-parent",
                   RowSpan =2,
                    HorizontalAlign = HorizontalAlign.Center

                };

                headerRow3.Cells.Add(srHeader);
                TableCell srDistrict = new TableCell
                {
                    Text = "Name Of District",
                    CssClass = "header-parent",
                    RowSpan = 2,
                    HorizontalAlign = HorizontalAlign.Center

                };

                headerRow3.Cells.Add(srDistrict);

                // Main header for Arrival
                TableCell arrivalHeader = new TableCell
                {
                    Text = "PROGRESSIVE MANDI ARRIVAL 2023-24(as per entry gate pass)",
                    ColumnSpan = 4,
                    CssClass = "header-parent",
                    HorizontalAlign = HorizontalAlign.Center
                };
                headerRow3.Cells.Add(arrivalHeader);

                // Main header for Purchased
                TableCell purchasedHeader = new TableCell
                {
                    Text = "PROGRESSIVE PADDY PURCHASED BY AGENCIES(as per Auction)",
                    ColumnSpan = 4,
                    CssClass = "header-parent",
                    HorizontalAlign = HorizontalAlign.Center
                };
                headerRow3.Cells.Add(purchasedHeader);


                

                TableCell daysHeader = new TableCell
                {
                    Text = "Days Procurement By All Agencies",
                    CssClass = "header-parent",
                    RowSpan = 2,
                    HorizontalAlign = HorizontalAlign.Center

                };
                headerRow3.Cells.Add(daysHeader);
                TableCell PurchaseHeader = new TableCell
                {
                    Text = "Total Purchase in JForm",
                    CssClass = "header-parent",
                    RowSpan = 2,
                    HorizontalAlign = HorizontalAlign.Center

                };
                headerRow3.Cells.Add(PurchaseHeader);

                //TableCell LastYearHeader = new TableCell
                //{
                //    Text = "Last Year on Same Date",
                //    CssClass = "header-parent",
                //    RowSpan = 2

                //};
                //headerRow3.Cells.Add(LastYearHeader);

                TableCell LiftHeader = new TableCell
                {
                    Text = "Lifting (Till Date)",
                    CssClass = "header-parent",
                    RowSpan = 2,
                    HorizontalAlign = HorizontalAlign.Center

                };
                headerRow3.Cells.Add(LiftHeader);

                TableCell FarmerBenefitHeader = new TableCell
                {
                    Text = "Total No. Of Farmers Benefited Till Date",
                    CssClass = "header-parent",
                    RowSpan = 2,
                    HorizontalAlign = HorizontalAlign.Center

                };
                headerRow3.Cells.Add(FarmerBenefitHeader);



                //TableCell LstYearPurHeader = new TableCell
                //{
                //    Text = "Last Year Total Purchase by All Agencies)",
                //    CssClass = "header-parent",
                //    RowSpan = 2

                //};
                //headerRow3.Cells.Add(LstYearPurHeader);

                ///


                // Placeholder cells for Sr No and Name of District
                TableCell HeaderCommon = new TableCell
                {
                    Text = "Common",

                    CssClass = "header-parent",
                    HorizontalAlign = HorizontalAlign.Center


                };

                headerRow4.Cells.Add(HeaderCommon);
                TableCell srGradeA = new TableCell
                {
                    Text = "Grade A",
                    CssClass = "header-parent",
                    HorizontalAlign = HorizontalAlign.Center


                };

                headerRow4.Cells.Add(srGradeA);

                // Main header for Arrival
                TableCell leviHeader = new TableCell
                {
                    Text = "Total Arrival",
                    
                    CssClass = "header-parent",
                    HorizontalAlign = HorizontalAlign.Center
                };
                headerRow4.Cells.Add(leviHeader);

                // Main header for Purchased
                TableCell DayHeader = new TableCell
                {
                    Text = "Day's Arrival",
                    
                    CssClass = "header-parent",
                    HorizontalAlign = HorizontalAlign.Center
                };
                headerRow4.Cells.Add(DayHeader);

                TableCell FoodHeader = new TableCell
                {
                    Text = "Food",
                    CssClass = "header-parent",
                    HorizontalAlign = HorizontalAlign.Center


                };
                headerRow4.Cells.Add(FoodHeader);

                TableCell hafedHeader = new TableCell
                {
                    Text = "Hafed",
                    CssClass = "header-parent",
                    HorizontalAlign = HorizontalAlign.Center

                };
                headerRow4.Cells.Add(hafedHeader);

                TableCell HWCHeader = new TableCell
                {
                    Text = "HWC",
                    CssClass = "header-parent",
                    HorizontalAlign = HorizontalAlign.Center

                };
                headerRow4.Cells.Add(HWCHeader);
                //TableCell AgroHeader = new TableCell
                //{
                //    Text = "AGRO",
                //    CssClass = "header-parent",
                //    HorizontalAlign = HorizontalAlign.Center

                //};
                //headerRow4.Cells.Add(AgroHeader);


                TableCell AgencyHeader = new TableCell
                {
                    Text = "Agency Total",
                    CssClass = "header-parent",
                    HorizontalAlign = HorizontalAlign.Center

                };
                headerRow4.Cells.Add(AgencyHeader);

                ((GridView)sender).Controls[0].Controls.AddAt(0, headerRow1);
                ((GridView)sender).Controls[0].Controls.AddAt(1, headerRow2);
            
                ((GridView)sender).Controls[0].Controls.AddAt(2, headerRow3);
                ((GridView)sender).Controls[0].Controls.AddAt(3, headerRow4);

            }

            // Apply the child header class to the existing GridView header
            if (e.Row.RowType == DataControlRowType.Header)
            {
                foreach (TableCell cell in e.Row.Cells)
                {
                    cell.CssClass = "header-child";
                }
            }

        }
        public DataTable BindGrid(string Type, string Auction_date, string SortExpression = null)
        {
            DataTable dt = GetReportData( Type ,Auction_date);

            decimal totalCommon = 0;
            decimal totalGradeA = 0;
            decimal totalTotalLeviable = 0;
            decimal totalDayArrival= 0;
            decimal  totalFood = 0;
            decimal totalHafed = 0;
            decimal totalHWC = 0;
            decimal totalPurchase= 0;
            decimal totalAgency = 0;
            decimal totalProcurement = 0;
            decimal totalLastYearSameDay = 0;
            decimal totalLifting= 0;
            decimal totalLastYearPurchase = 0;
            int totalFarmerBenefited = 0;
            // Calculate totals for other columns as needed

                foreach (DataRow row in dt.Rows)
                {
                    totalCommon += Convert.ToDecimal(row["Common"]);
                    totalGradeA += Convert.ToDecimal(row["GradeA"]);
                    totalTotalLeviable += Convert.ToDecimal(row["Total"]);
                totalDayArrival += Convert.ToDecimal(row["DayArrival"]);
                totalFood += Convert.ToDecimal(row["Food"]);
                totalHafed += Convert.ToDecimal(row["Hafed"]);
                totalHWC += Convert.ToDecimal(row["HWC"]);
               
                totalAgency += Convert.ToDecimal(row["AgencyTotal"]);
                totalProcurement += Convert.ToDecimal(row["Days"]);
                //totalLastYearSameDay += Convert.ToDecimal(row["LastYear"]);
                totalLifting += Convert.ToDecimal(row["Lifiting"]);
                totalPurchase += Convert.ToDecimal(row["TotalPurchase"]);
                totalFarmerBenefited += Convert.ToInt32(row["FarmerBenefitedtotal"]); 
                //totalLastYearPurchase += Convert.ToDecimal(row["LastYearTotal"]);
                // Calculate totals for other columns and update respective total variables
            }

            

            
                 DataRow drow = dt.NewRow();
            drow["District"] = "Total";



            drow["Common"] = totalCommon;



            drow["GradeA"] = totalGradeA;



            drow["Total"] = totalTotalLeviable;



            drow["DayArrival"] = totalDayArrival;



            drow["Food"] = totalFood;



            drow["Hafed"] = totalHafed;



            drow["HWC"] = totalHWC;



           



            drow["AgencyTotal"] = totalAgency;



            drow["Days"] = totalProcurement;



            //drow["LastYear"] = totalLastYearSameDay;



            drow["Lifiting"] = totalLifting;
            drow["totalPurchase"] = totalPurchase;
            drow["FarmerBenefitedtotal"] = totalFarmerBenefited;


            //drow["LastYearTotal"] = totalLastYearPurchase;
            dt.Rows.Add(drow);




            return dt;
        }

        protected void txtDatepicker_TextChanged(object sender, EventArgs e)
        {
            
           

            grd.DataSource = BindGrid("", txtDatepicker.Text.Trim());
            grd.DataBind();
        }

        protected void btnPdf_Click(object sender, EventArgs e)
        {


            DateTime date = DateTime.Parse(txtDatepicker.Text); // Parse the input string to a DateTime object
            string fileName = "PaddyPurchaseReport_" + date.ToString("dd-MM-yyyy") + ".pdf";

            DataTable dt = BindGrid("", txtDatepicker.Text.Trim());

            // Create a memory stream to store the PDF content
            MemoryStream memoryStream = new MemoryStream();
            Document document = new Document(PageSize.A4.Rotate(), 10f, 10f, 10f, 10f);
            PdfWriter writer = PdfWriter.GetInstance(document, memoryStream);

            document.Open();

            // Create a PDF table
            PdfPTable pdfTable = new PdfPTable(dt.Columns.Count);
            pdfTable.WidthPercentage = 100;

            // Set cell padding and border
            pdfTable.DefaultCell.Padding = 3;
            pdfTable.DefaultCell.BorderWidth = 1;

            // Set cell alignment
            pdfTable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;

            // Add parent header row
            pdfTable.AddCell(new PdfPCell(new Phrase("FOOD, CIVIL SUPPLIES & CONSUMER AFFAIRS DEPARTMENT", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { Colspan = 14, HorizontalAlignment = Element.ALIGN_CENTER,BackgroundColor =BaseColor.LIGHT_GRAY });
            pdfTable.CompleteRow();

            pdfTable.AddCell(new PdfPCell(new Phrase("PROGRESSIVE ARRIVAL/ PROCUREMENT OF PADDY DURING KHARIF 2023 - 24 AS ON " + DateFormat(), new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { Colspan = 11, HorizontalAlignment = Element.ALIGN_CENTER ,BackgroundColor = new BaseColor(240, 248, 255) });
            pdfTable.AddCell(new PdfPCell(new Phrase("(All Fig in MT.)", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { Colspan = 3, HorizontalAlignment = Element.ALIGN_RIGHT, BackgroundColor = new BaseColor(240, 248, 255) });
            pdfTable.CompleteRow();

            // Add parent header row
            pdfTable.AddCell(new PdfPCell(new Phrase("Sr No.", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) {Rowspan =2, HorizontalAlignment = Element.ALIGN_CENTER });
            pdfTable.AddCell(new PdfPCell(new Phrase("Name Of District", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { Rowspan = 2, HorizontalAlignment = Element.ALIGN_CENTER });

    
            pdfTable.AddCell(new PdfPCell(new Phrase("PROGRESSIVE MANDI ARRIVAL 2023-24(as per entry gate pass)", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { Colspan = 4, HorizontalAlignment = Element.ALIGN_CENTER });
            pdfTable.AddCell(new PdfPCell(new Phrase("PROGRESSIVE PADDY PURCHASED BY AGENCIES(as per Auction)", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { Colspan = 4, HorizontalAlignment = Element.ALIGN_CENTER });
          
            pdfTable.AddCell(new PdfPCell(new Phrase("Days Procurement By All Agencies", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { Rowspan =2, HorizontalAlignment = Element.ALIGN_CENTER });
            pdfTable.AddCell(new PdfPCell(new Phrase("Total Purchase in JForm", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { Rowspan = 2, HorizontalAlignment = Element.ALIGN_CENTER });
            //pdfTable.AddCell(new PdfPCell(new Phrase("Last Year on Same Date", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) {Rowspan=2, HorizontalAlignment = Element.ALIGN_CENTER });
            pdfTable.AddCell(new PdfPCell(new Phrase("Lifting (Till Date)", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { Rowspan=2, HorizontalAlignment = Element.ALIGN_CENTER });
            pdfTable.AddCell(new PdfPCell(new Phrase("Total No. Of Farmers Benefited Till Date", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { Rowspan = 2, HorizontalAlignment = Element.ALIGN_CENTER });
            //pdfTable.AddCell(new PdfPCell(new Phrase("Last Year Total Purchase by All Agencies)", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { Rowspan=2, HorizontalAlignment = Element.ALIGN_CENTER });
            pdfTable.CompleteRow();
           

            pdfTable.AddCell(new PdfPCell(new Phrase("Common", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) {  HorizontalAlignment = Element.ALIGN_CENTER });
            pdfTable.AddCell(new PdfPCell(new Phrase("Grade A", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) {  HorizontalAlignment = Element.ALIGN_CENTER });

            pdfTable.AddCell(new PdfPCell(new Phrase("Total Arrival", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { HorizontalAlignment = Element.ALIGN_CENTER });
            pdfTable.AddCell(new PdfPCell(new Phrase("Day's Arrival", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { HorizontalAlignment = Element.ALIGN_CENTER });
            pdfTable.AddCell(new PdfPCell(new Phrase("Food", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { HorizontalAlignment = Element.ALIGN_CENTER });
            pdfTable.AddCell(new PdfPCell(new Phrase("Hafed", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { HorizontalAlignment = Element.ALIGN_CENTER });
            pdfTable.AddCell(new PdfPCell(new Phrase("HWC", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { HorizontalAlignment = Element.ALIGN_CENTER });

            //pdfTable.AddCell(new PdfPCell(new Phrase("AGRO", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { HorizontalAlignment = Element.ALIGN_CENTER });
            pdfTable.AddCell(new PdfPCell(new Phrase("Agency Total", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { HorizontalAlignment = Element.ALIGN_CENTER });
            
            pdfTable.CompleteRow();

            // Add data rows
            //foreach (DataRow dr in dt.Rows)
            //{
            //    pdfTable.AddCell(new PdfPCell(new Phrase(dr["SrNo"].ToString(), new Font(Font.FontFamily.HELVETICA, 10))) { HorizontalAlignment = Element.ALIGN_CENTER });
            //    pdfTable.AddCell(new PdfPCell(new Phrase(dr["District"].ToString(), new Font(Font.FontFamily.HELVETICA, 10))) { HorizontalAlignment = Element.ALIGN_CENTER });
            //    pdfTable.AddCell(new PdfPCell(new Phrase(dr["Common"].ToString(), new Font(Font.FontFamily.HELVETICA, 10))) { HorizontalAlignment = Element.ALIGN_CENTER });
            //    pdfTable.AddCell(new PdfPCell(new Phrase(dr["GradeA"].ToString(), new Font(Font.FontFamily.HELVETICA, 10))) { HorizontalAlignment = Element.ALIGN_CENTER });
            //    pdfTable.AddCell(new PdfPCell(new Phrase(dr["Total"].ToString(), new Font(Font.FontFamily.HELVETICA, 10))) { HorizontalAlignment = Element.ALIGN_CENTER });
            //    pdfTable.AddCell(new PdfPCell(new Phrase(dr["DayArrival"].ToString(), new Font(Font.FontFamily.HELVETICA, 10))) { HorizontalAlignment = Element.ALIGN_CENTER });
            //    pdfTable.AddCell(new PdfPCell(new Phrase(dr["Food"].ToString(), new Font(Font.FontFamily.HELVETICA, 10))) { HorizontalAlignment = Element.ALIGN_CENTER });
            //    pdfTable.AddCell(new PdfPCell(new Phrase(dr["Hafed"].ToString(), new Font(Font.FontFamily.HELVETICA, 10))) { HorizontalAlignment = Element.ALIGN_CENTER });
            //    pdfTable.AddCell(new PdfPCell(new Phrase(dr["HWC"].ToString(), new Font(Font.FontFamily.HELVETICA, 10))) { HorizontalAlignment = Element.ALIGN_CENTER });
            //    pdfTable.AddCell(new PdfPCell(new Phrase(dr["Agro"].ToString(), new Font(Font.FontFamily.HELVETICA, 10))) { HorizontalAlignment = Element.ALIGN_CENTER });
            //    pdfTable.AddCell(new PdfPCell(new Phrase(dr["AgencyTotal"].ToString(), new Font(Font.FontFamily.HELVETICA, 10))) { HorizontalAlignment = Element.ALIGN_CENTER });
            //    pdfTable.AddCell(new PdfPCell(new Phrase(dr["Days"].ToString(), new Font(Font.FontFamily.HELVETICA, 10))) { HorizontalAlignment = Element.ALIGN_CENTER });
            //    //pdfTable.AddCell(new PdfPCell(new Phrase(dr["LastYear"].ToString(), new Font(Font.FontFamily.HELVETICA, 10))) { HorizontalAlignment = Element.ALIGN_CENTER });
            //    pdfTable.AddCell(new PdfPCell(new Phrase(dr["Lifiting"].ToString(), new Font(Font.FontFamily.HELVETICA, 10))) { HorizontalAlignment = Element.ALIGN_CENTER });
            //    //pdfTable.AddCell(new PdfPCell(new Phrase(dr["LastYearTotal"].ToString(), new Font(Font.FontFamily.HELVETICA, 10))) { HorizontalAlignment = Element.ALIGN_CENTER });

            //    pdfTable.CompleteRow();
            //}
            for (int rowIndex = 0; rowIndex < dt.Rows.Count; rowIndex++)
            {
                DataRow dr = dt.Rows[rowIndex];

                PdfPCell cell;

                for (int columnIndex = 0; columnIndex < dt.Columns.Count; columnIndex++)
                {
                    string cellText = dr[columnIndex].ToString();
                    Font cellFont = new Font(Font.FontFamily.HELVETICA, 10);

                    // Check if it's the last row and set the font to bold
                    if (rowIndex == dt.Rows.Count - 1)
                    {
                        cellFont.SetStyle(Font.BOLD);
                    }

                    if (columnIndex < 2)
                    {
                        // Center-align text for the first and second columns
                        cell = new PdfPCell(new Phrase(cellText, cellFont)) { HorizontalAlignment = Element.ALIGN_CENTER };
                    }
                    else
                    {
                        // Right-align text for other columns
                        cell = new PdfPCell(new Phrase(cellText, cellFont)) { HorizontalAlignment = Element.ALIGN_RIGHT };
                    }


                    
                    pdfTable.AddCell(cell);
                }

                pdfTable.CompleteRow();
            }
            //// Add footer row
            //pdfTable.AddCell(new PdfPCell(new Phrase("This Report is generated through ekharid portal of FOOD, CIVIL SUPPLIES & CONSUMER AFFAIRS DEPARTMENT ", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { Colspan = 14, HorizontalAlignment = Element.ALIGN_CENTER});
            //pdfTable.CompleteRow();
            //pdfTable.AddCell(new PdfPCell(new Phrase("Date:" + DateTime.Now.ToString ("dd/MM/yyyy") , new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { Colspan = 14, HorizontalAlignment = Element.ALIGN_RIGHT });
            //pdfTable.CompleteRow();
            //pdfTable.AddCell(new PdfPCell(new Phrase("Time:" + DateTime.Now.ToString("hh:mm tt"), new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD))) { Colspan = 14, HorizontalAlignment = Element.ALIGN_RIGHT });
            //pdfTable.CompleteRow();

            document.Add(pdfTable);
            // Add the footer text outside the table
            Paragraph footerText = new Paragraph("This Report is generated through ekharid portal of FOOD, CIVIL SUPPLIES & CONSUMER AFFAIRS DEPARTMENT", new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD));
            footerText.Alignment = Element.ALIGN_LEFT;
            document.Add(footerText);

            Paragraph dateText = new Paragraph("Date: " + DateTime.Now.ToString("dd/MM/yyyy"), new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD));
            dateText.Alignment = Element.ALIGN_LEFT;
            document.Add(dateText);

            Paragraph timeText = new Paragraph("Time: " + DateTime.Now.ToString("hh:mm tt"), new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD));
            timeText.Alignment = Element.ALIGN_LEFT;
            document.Add(timeText);


            document.Close();

            // Save the PDF to a file
            Response.Clear();
            Response.AddHeader("content-disposition", "attachment;filename="+ fileName);
            Response.ContentType = "application/pdf";
            Response.BinaryWrite(memoryStream.ToArray());
            Response.End();
        }

      protected void btnexcel_Click(object sender, EventArgs e)
{

           

            DateTime date = DateTime.Parse(txtDatepicker.Text); // Parse the input string to a DateTime object
             string fileName = "PaddyPurchaseReport_" + date.ToString("dd-MM-yyyy") + ".xls";

            DataTable dt = BindGrid("", txtDatepicker.Text.Trim());
            Response.Clear();
            Response.AddHeader("content-disposition", "attachment;filename="+ fileName);
            Response.Charset = "";
            Response.ContentType = "application/ms-excel";

            // Use HTML table to structure your data
            Response.Write("<table border='1'>");


            Response.Write("<tr>");
            Response.Write("<th colspan='14' style='text-align: center;white-space: normal;background-color:LightGray;'>FOOD, CIVIL SUPPLIES & CONSUMER AFFAIRS DEPARTMENT</th>");

            Response.Write("</tr>");
            Response.Write("<tr>");
            Response.Write("<th colspan='11' style='text-align: center;white-space: normal;background-color: aliceBlue;' >PROGRESSIVE ARRIVAL/ PROCUREMENT OF PADDY DURING KHARIF 2023 - 24 AS ON "+ DateFormat() + "</th>");
            Response.Write("<th colspan='3'  style='text-align: right;white-space: normal;background-color: aliceBlue;'>(All Fig in MT.)</th>");
            Response.Write("</tr>");

          


            // Add parent header row
            Response.Write("<tr>");
            Response.Write("<th  style='text-align: center;white-space: normal;' rowspan='2' >Sr No.</th>");
            Response.Write("<th  style='text-align: center;white-space: normal;' rowspan='2' >Name Of District</th>");
            Response.Write("<th  colspan='4'  style='text-align: center;white-space: normal;'>PROGRESSIVE MANDI ARRIVAL 2023-24(as per entry gate pass)</th>");
            Response.Write("<th  colspan='4'  style='text-align: center;white-space: normal;'>PROGRESSIVE PADDY PURCHASED BY AGENCIES(as per Auction)</th>");
        
            Response.Write("<th  rowspan='2'  style='text-align: center;white-space: normal;'>Days Procurement By All Agencies</th>");
            Response.Write("<th  rowspan='2'  style='text-align: center;white-space: normal;'>Total Purchase in JForm</th>");
            //Response.Write("<th  rowspan='2'  style='text-align: center;white-space: normal;'>Last Year on Same Date</th>");
            Response.Write("<th  rowspan='2'  style='text-align: center;white-space: normal;'>Lifting (Till Date)</th>");
            Response.Write("<th  rowspan='2'  style='text-align: center;white-space: normal;'>Total No. Of Farmers Benefited Till Date</th>");
            //Response.Write("<th  rowspan='2'  style='text-align: center;white-space: normal;'>Last Year Total Purchase by All Agencies)</th>");
            Response.Write("</tr>");

            // Add child header row
            Response.Write("<tr>");
          
            Response.Write("<th style='text-align: center;white-space: normal;'>Common</th>");
            Response.Write("<th  style='text-align: center;white-space: normal;'>Grade A</th>");
            Response.Write("<th  style='text-align: center;white-space: normal;'>Total Arrival</th>");
            Response.Write("<th  style='text-align: center;white-space: normal;'>Day's Arrival</th>");
            Response.Write("<th  style='text-align: center;white-space: normal;'>Food</th>");
            Response.Write("<th  style='text-align: center;white-space: normal;'>Hafed</th>");
            Response.Write("<th  style='text-align: center;white-space: normal;'>HWC</th>");
            //Response.Write("<th  style='text-align: center;white-space: normal;'>Agro</th>");
            Response.Write("<th  style='text-align: center;white-space: normal;'>Agency Total</th>");
       
           
            Response.Write("</tr>");

            // Add data rows
            //foreach (DataRow dr in dt.Rows)
            //{
            //    Response.Write("<tr>");
            //    foreach (DataColumn column in dt.Columns)
            //    {
            //        Response.Write("<td>" + HttpUtility.HtmlEncode(dr[column].ToString()) + "</td>");
            //    }
            //    Response.Write("</tr>");
            //}
           
            foreach (DataRow dr in dt.Rows)
            {
                
                Response.Write("<tr" + (dt.Rows.IndexOf(dr) == dt.Rows.Count - 1 ? " style='font-weight: bold;'" : "") + ">");
                foreach (DataColumn column in dt.Columns)
                {
                    Response.Write("<td style='text-align:" +(column.Ordinal<2 ? "center" :"right") +";'>"   + HttpUtility.HtmlEncode(dr[column].ToString()) + "</td>");
                }
                Response.Write("</tr>");
            }

            Response.Write("</table>");
            Response.End();


        }

    }
}