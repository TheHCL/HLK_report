using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Windows.Kits.Hardware.ObjectModel;
using Microsoft.Windows.Kits.Hardware.ObjectModel.DBConnection;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.XSSF.Util;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {

        public static ProjectManager ConnectToController()
        {

            XmlDocument doc = new XmlDocument();
            doc.Load("C:\\HLK\\JobsWorkingDir\\ConfigManager\\Log\\MachineId.xml");
            string controllerName = doc.DocumentElement.ChildNodes[0].InnerText;
            ProjectManager manager = null;
            manager = new DatabaseProjectManager(controllerName);
            return manager;
        }

        public static IEnumerable<Test> Test_list(String pro_name)
        {
            ProjectManager manager = ConnectToController();
            Project proj = manager.GetProject(pro_name);
            IEnumerable<Test> tests = proj.GetTests().OrderBy(f => f.Name);
            return tests;
        }

        public static XSSFCellStyle def_style(IWorkbook workbook,int type)
        { 
            
            XSSFCellStyle my_style = (XSSFCellStyle)workbook.CreateCellStyle();
            XSSFFont font1 = (XSSFFont)workbook.CreateFont();
            
            my_style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            my_style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            my_style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            my_style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            font1.FontName = "Calibri";
            font1.FontHeightInPoints = 10;
            my_style.SetFont(font1);

            var pass_cor = new XSSFColor(new byte[] { 0, 176, 80 });
            var fail_cor = new XSSFColor(new byte[] { 255, 0, 0 });
            var not_cor = new XSSFColor(new byte[] { 166, 166, 166 });

            if (type == 0) //pass
            {
                my_style.SetFillForegroundColor(pass_cor);
                my_style.FillPattern = FillPattern.SolidForeground;
                my_style.VerticalAlignment = VerticalAlignment.Center;
                my_style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

            }
            else if (type == 1) //fail
            {
                my_style.SetFillForegroundColor(fail_cor);
                my_style.FillPattern = FillPattern.SolidForeground;
                my_style.VerticalAlignment = VerticalAlignment.Center;
                my_style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            }
            else if (type ==2) //not run
            {
                my_style.SetFillForegroundColor(not_cor);
                my_style.FillPattern = FillPattern.SolidForeground;
                my_style.VerticalAlignment = VerticalAlignment.Center;
                my_style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            }
            else // default
            {
                return my_style;
            }
            return my_style;
            


        }

        public Form1()
        {
            

            ProjectManager manager = ConnectToController();
            
           

            InitializeComponent();
            label1.Visible=false;
            label2.Visible=false;
            button1.Visible=false;
            Array projects = manager.GetProjectNames().ToArray();
            
            foreach (string str in projects)
            {
                listBox1.Items.Add(str);
            }
            //listBox1.Items.Add(controllerName);
            
        }

       
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            label1.Visible = true;
            label2.Visible = true;
            button1.Visible = true;
            string text = listBox1.GetItemText(listBox1.SelectedItem); //Get the selected project name
           
            IEnumerable<Test> my_res = Test_list(text);
            int count_pass = 0;
            int count_fail = 0;
            int count_not = 0;
            int count_total = 0;
            string numbers_title;
            string numbers;
            foreach (Test test in my_res)
            {
                // here we can get test results list (Name)(Status)(AreFiltersApplied)
                
                if (test.Status.ToString() == "Passed")
                {
                    count_pass += 1;
                }
                else if (test.Status.ToString() == "Failed")
                {
                    count_fail += 1;
                }
                else
                {
                    count_not += 1;
                }
                
            }
            count_total = count_pass+count_fail+count_not;
            numbers_title = String.Format("{0,-7}", "Pass")+"\n";
            numbers = String.Format("{0,3}", count_pass.ToString()) + "\n";
            numbers_title += String.Format("{0,-9}", "Fail") + "\n";
            numbers += String.Format("{0,3}", count_fail.ToString()) + "\n";
            numbers_title += String.Format("{0,-9}", "NotRun") + "\n";
            numbers += String.Format("{0,3}", count_not.ToString()) + "\n";
            numbers_title += String.Format("{0,-9}", "Total");
            numbers += String.Format("{0,3}", count_total.ToString()) ;
            label1.Text = numbers_title;
            label2.Text = numbers;


        }

        private void button1_Click(object sender, EventArgs e) //click button to convert the results to Excel
        {
            string text = listBox1.GetItemText(listBox1.SelectedItem);
            SaveFileDialog save_diag = new SaveFileDialog();
            save_diag.Filter = "Excel File(.xlsx)|*.xlsx";
            save_diag.Title = "Save to xlsx";
            save_diag.ShowDialog();

            
            if (save_diag.FileName != "")
            {
                IWorkbook workbook = new XSSFWorkbook();
                XSSFCellStyle cs_test_name = def_style(workbook,3);
                XSSFCellStyle cs_pass = def_style(workbook,0);
                XSSFCellStyle cs_fail = def_style(workbook,1);
                XSSFCellStyle cs_not = def_style(workbook,2);
                
                
                
                

                

                

                ISheet sheet = workbook.CreateSheet("Test Summary");
                sheet.CreateRow(0);
                sheet.GetRow(0).CreateCell(1).SetCellValue("Units");
                sheet.GetRow(0).CreateCell(2).SetCellValue("my machine");
                sheet.CreateRow(1);
                sheet.GetRow(1).CreateCell(0).SetCellValue("No.");
                sheet.GetRow(1).CreateCell(1).SetCellValue("Test items");
                sheet.GetRow(1).CreateCell(2).SetCellValue("Results");
                sheet.GetRow(1).CreateCell(3).SetCellValue("Filter ID");
                int my_row = 2;
                int count = 1;
                sheet.SetColumnWidth(1, (int)(59.22*256));
                sheet.SetColumnWidth(2, (int)(19.22*256));
                
                IEnumerable<Test> t = Test_list(text);

                foreach (Test a in t)
                {
                    sheet.CreateRow(my_row);
                    sheet.GetRow(my_row).CreateCell(0).SetCellValue(count);
                    sheet.GetRow(my_row).GetCell(0).CellStyle = cs_test_name;
                    sheet.GetRow(my_row).CreateCell(1).SetCellValue(a.Name.ToString());
                    sheet.GetRow(my_row).GetCell(1).CellStyle = cs_test_name;
                    if (a.Status.ToString() == "Passed")
                    {
                        sheet.GetRow(my_row).CreateCell(2).SetCellValue(a.Status.ToString());
                        sheet.GetRow(my_row).GetCell(2).CellStyle = cs_pass;
                    }
                    else if (a.Status.ToString() == "Failed")
                    {
                        sheet.GetRow(my_row).CreateCell(2).SetCellValue(a.Status.ToString());
                        sheet.GetRow(my_row).GetCell(2).CellStyle = cs_fail;
                    }
                    else 
                    {
                        sheet.GetRow(my_row).CreateCell(2).SetCellValue(a.Status.ToString());
                        sheet.GetRow(my_row).GetCell(2).CellStyle = cs_not;
                    }

                    sheet.GetRow(my_row).CreateCell(3);
                    sheet.GetRow(my_row).GetCell(3).CellStyle = cs_test_name;
                    my_row++;
                    count++;
                }
                
                using (FileStream stream = new FileStream(save_diag.FileName, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(stream);
                }
                MessageBox.Show("Done","Info");
            
               


            }
        }
    }
}
