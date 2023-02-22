using COVID_Risks.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using WinFormsApp1.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace COVID_Risks
{
    public partial class Form1 : Form
       
    {
        public String _filePath = "";
        public Form1()
        {
            InitializeComponent();
        }
           
        

        private void Form1_Load(object sender, EventArgs e)
        {
            
        
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel files| *.xlsx";
            ofd.DefaultExt = ".xlsx";
            Nullable<bool> dialogOK = Convert.ToBoolean(ofd.ShowDialog());
            if (dialogOK == true)
            {
                _filePath = ofd.FileName;
            }


            Person.personlist = ExcelHelper.GetPersonData(_filePath);
            
            foreach (Person person in Person.personlist)
            {
                PersonList.Items.Add(person.Fio);
            }
            
            
        }

        private void splitContainer2_SplitterMoved(object sender, SplitterEventArgs e)
        {

        }

        private void PersonList_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
          PersonView1.Items.Clear();
            PersonView1.Items.Add(Person.personlist[);
           

        }

        private void PersonView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void PersonList_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}