using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using ClosedXML.Excel;
using System.Configuration;
using System.Web;
using DocumentFormat.OpenXml.Drawing;
using ErrorEmp;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using Microsoft.AspNetCore.Http;

namespace EmployeeClass
{
    public class Class1
    {

        public DataSet SelectData(string str,string p)
        {
            string q = @"Data Source=DESKTOP-GI21EV1\SQLEXPRESS; Initial Catalog=Employee_ExportImport; Integrated Security=true;";
            SqlConnection con=new SqlConnection(q);
            con.Open();
            SqlCommand cmd = new SqlCommand(str, con);
            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            adapter.Fill(ds);        

            using (System.Data.DataTable dt = new System.Data.DataTable())
            {
                adapter.Fill(dt);
                using (XLWorkbook wb = new XLWorkbook())
                {
                    wb.Worksheets.Add(dt,"Emp");
                    string temp = p+".xlsx";
                    wb.SaveAs(temp);
                    //wb.SaveAs("D:\\intern\\kan.xlsx");
                }
            }
            con.Close();
            return ds;
        }

        public void InsertData(string p1,string p2)
        {
            string q = @"Data Source=DESKTOP-GI21EV1\SQLEXPRESS; Initial Catalog=Employee_ExportImport; Integrated Security=true;";
            SqlConnection con = new SqlConnection(q);
            con.Open();

            //using (XLWorkbook workBook = new XLWorkbook("D:\\intern\\kan.xlsx"))//+".xlsx"p1
            using (XLWorkbook workBook = new XLWorkbook(p1))
            {
                IXLWorksheet workSheet = workBook.Worksheet("Emp");
                DocumentFormat.OpenXml.Spreadsheet.Worksheet excelSheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet();

                System.Data.DataTable dt = new System.Data.DataTable();
                System.Data.DataTable dataTable = new System.Data.DataTable();

                dataTable.Columns.Add("Emp_Id");
                dataTable.Columns.Add("Emp_Name");
                dataTable.Columns.Add("Emp_Email");
                dataTable.Columns.Add("Mobile_No");
                dataTable.Columns.Add("Emp_Age");
               
                dt.Columns.Add("Emp_Id");
                dt.Columns.Add("Emp_Name");
                dt.Columns.Add("Emp_Email");
                dt.Columns.Add("Mobile_No");
                dt.Columns.Add("Emp_Age");
                dt.Columns.Add("Error");
                
                List<ErrorEmployee> l = new List<ErrorEmployee>();

                int flag = 0;
                int i = 0;
                foreach(IXLRow r in workSheet.Rows())
                {
                    DataRow dr = dataTable.NewRow();
                    dr["Emp_Id"] = r.Cell("A").Value;
                    //dr["Emp_Id"] = r.Cell("A").Value;
                    dr["Emp_Name"] = r.Cell("B").Value;
                    dr["Emp_Email"] = r.Cell("C").Value;
                    dr["Mobile_No"] = r.Cell("D").Value;
                    dr["Emp_Age"] =r.Cell("E").Value;
                    //dr["Emp_Age"] = r.Cell("E").Value;
                    dataTable.Rows.Add(dr);
                }
                int f = 0;
                foreach (IXLRow row in workSheet.Rows())
                {
                    if(f==0)
                    {
                        f = 1;
                    }
                    
                    else
                    {
                        //int age = 0, id = 0;
                        Int64  id = 0;
                        string age = "";
                        //string age = "", id = "";
                        string name = "", email = "", mobile = "";
                        id = Int64.Parse(dataTable.Rows[i][0].ToString());
                        //id = dataTable.Rows[i][0].ToString();
                        name = dataTable.Rows[i][1].ToString();
                        email = dataTable.Rows[i][2].ToString();
                        mobile = dataTable.Rows[i][3].ToString();
                        //age=Int32.Parse(dataTable.Rows[i][4].ToString());
                        age = dataTable.Rows[i][4].ToString();

                        if (String.IsNullOrWhiteSpace(name) || String.IsNullOrWhiteSpace(email) || String.IsNullOrWhiteSpace(mobile) || String.IsNullOrWhiteSpace(dataTable.Rows[i][0].ToString()) || String.IsNullOrWhiteSpace(dataTable.Rows[i][4].ToString()))
                        {
                            flag = 1;
                            DataRow dr = dt.NewRow();
                            ///////Error Table 
                            dr["Emp_Id"] = id;
                            dr["Emp_Name"] = name;
                            dr["Emp_Email"] = email;
                            dr["Mobile_No"] = mobile;
                            dr["Emp_Age"] = age;
                            dr["Error"] = "Error";
                            dt.Rows.Add(dr);
                        }
                        else
                        {
                            //int Eid = Convert.ToUInt16(id);
                            //int Eid = int.TryParse(dataTable.Rows[i][0].ToString(), out );
                            //int Eid=Int32.Parse(id);
                            int Eage=Int32.Parse(age);
                            string str = "insert into Employee (Emp_Id,Emp_Name,Emp_Email,Mobile_No,Emp_Age) values('" + id + "','" + name + "','" + email + "','" + mobile + "','" + Eage + "')";
                            SqlCommand cmd = new SqlCommand(str, con);
                            SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                            adapter.SelectCommand.ExecuteNonQuery();
                        }
                    }
                    
                    i++;
                }
                if(flag==1)
                {
                    using (System.Data.DataTable dt1 = new System.Data.DataTable())
                    {
                        using (XLWorkbook wb = new XLWorkbook())
                        {
                            wb.Worksheets.Add(dt, "EmpError");
                            //Environment.GetFolderPath(Environment.SpecialFolder.Desktop//);
                            wb.SaveAs("C:\\Users\\kanda\\Downloads\\kanError.xlsx");//p2//p1D:\\intern\\kanError.xlsx                  
                        }
                    }
                }
                

                

                ////////this is valid code but not proper//////////////////////////

                //foreach (IXLRow row in workSheet.Rows())
                //{
                //    if (firstRow)
                //    {
                //        firstRow = false;
                //    }
                //    else
                //    {
                //        i = 0;
                //        id = 0; age=0;
                //        name = ""; email = ""; mobile="";
                //        flag = 0;
                //        foreach (IXLCell cell in row.Cells())
                //        {
                //            i++;
                //            if (String.IsNullOrWhiteSpace(cell.Value.ToString()))
                //            {
                //                flag = 1;
                //                check = 1;
                //            }
                //            if (i==1)
                //            {
                //                if (String.IsNullOrWhiteSpace(cell.Value.ToString()))
                //                {
                //                    dr["Error"] = "id";
                //                    id = -1;
                //                }
                //                else
                //                {
                //                    id = Int32.Parse(cell.Value.ToString());
                //                }                             
                //            }
                //            else if(i==2)
                //            {
                //                if (String.IsNullOrWhiteSpace(cell.Value.ToString()))
                //                {
                //                    dr["Error"] = "name";
                //                    name = "Error";
                //                }
                //                else
                //                {
                //                    name = cell.Value.ToString();
                //                }
                                
                //            }
                //            else if (i == 3)
                //            {
                //                if (String.IsNullOrWhiteSpace(cell.Value.ToString()))
                //                {
                //                    dr["Error"] = "Email";
                //                    email = "Error";
                //                }
                //                else
                //                {
                //                    email = cell.Value.ToString();
                //                }
                //            }
                //            else if (i == 4)
                //            {
                //                if (String.IsNullOrWhiteSpace(cell.Value.ToString()))
                //                {
                //                    dr["Error"] = "mobile";
                //                    mobile = "Error";
                //                }
                //                else
                //                {
                //                    mobile = cell.Value.ToString();
                //                }
                //            }
                //            else if(i==5)
                //            {

                //                if (String.IsNullOrWhiteSpace(cell.Value.ToString()))
                //                {
                //                    dr["Error"] = "age";
                //                    age = -1;
                //                }
                //                else
                //                {
                //                    age = Int32.Parse(cell.Value.ToString()); 
                //                }
                                
                //            }
                //        }
                //        if (flag == 1)
                //        {
                //            //if (temp == 0)
                //            //{
                //            //    dt.Columns.Add("Emp_Id");
                //            //    dt.Columns.Add("Emp_Name");
                //            //    dt.Columns.Add("Emp_Email");
                //            //    dt.Columns.Add("Mobile_No");
                //            //    dt.Columns.Add("Emp_Age");
                //            //    temp = 1;
                //            //}
                //           // DataRow dr = dt.NewRow();
                //            dr["Emp_Id"] = id;
                //            dr["Emp_Name"] = name;
                //            dr["Emp_Email"] = email;
                //            dr["Mobile_No"] = mobile;
                //            dr["Emp_Age"] = age;
                //            dt.Rows.Add(dr);
                //        }
                //        else
                //        {
                                
                //        }      
                        
                //    }
                //}
                //if (check == 1)
                //{
                //    using (DataTable dt1 = new DataTable())
                //    {
                //        using (XLWorkbook wb = new XLWorkbook())
                //        {
                //            wb.Worksheets.Add(dt, "Emp");
                //            wb.SaveAs("D:\\intern\\kanError.xlsx");//p2
                //        }
                //    }
                //}
            }
            con.Close();
        }
    }
}
