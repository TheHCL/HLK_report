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

     
       
        public Form1()
        {
            

            ProjectManager manager = ConnectToController();
            
           

            InitializeComponent();
            
            Array projects = manager.GetProjectNames().ToArray();
            
            foreach (string str in projects)
            {
                listBox1.Items.Add(str);
            }
            //listBox1.Items.Add(controllerName);
            
        }

       
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string text = listBox1.GetItemText(listBox1.SelectedItem); //Get the selected project name
            ProjectManager manager = ConnectToController();
            Console.WriteLine(text);
            Project proj = manager.GetProject(text);
            IEnumerable<Test> my_res = proj.GetTests().OrderBy(f => f.Name);
            foreach (Test test in my_res)
            {
                Console.WriteLine(test.Name); // here we can get test results list (Name)(Status)(AreFiltersApplied)
                /*foreach (TestResult testResult in test.GetTestResults())
                {
                    //Console.WriteLine(testResult.);
                    //MessageBox.Show(testResult.Machine.Name.ToString(),"fuck");
                    //MessageBox.Show(testResult.Test.Name.ToString(), "title");
                    my_res.Append
                }*/
            }


        }


    }
}
