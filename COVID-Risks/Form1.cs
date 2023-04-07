using COVID_Risks.Models;
using COVID_Risks.Properties;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.InteropServices;
using WinFormsApp1.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace COVID_Risks
{
    public partial class Form1 : Form

    {


        static String _filePath = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void OpenButton_Click(object sender, EventArgs e)
        {
            PersonList.Items.Clear();
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel files| *.xlsx";
            ofd.DefaultExt = ".xlsx";
            Nullable<bool> dialogOK = Convert.ToBoolean(ofd.ShowDialog());
            if (dialogOK == true)
            {
                _filePath = ofd.FileName;
            }


            Person.personlist = ExcelHelper.GetPersonData(_filePath);
            PersonListRefresh();
        }



        public void PersonListRefresh()
        {
            PersonList.Items.Clear();

            foreach (Person person in Person.personlist)
            {
                PersonList.Items.Add(person.Fio);
            }
        }


        private void PersonList_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (PersonList.SelectedItems.Count == 0)
            {
                return;
            }

            PersonView1.Items.Clear();
            int selectedIndex = PersonList.SelectedIndices[0];

            ListViewItem Fio = new ListViewItem("���");
            Fio.SubItems.Add(Person.personlist[selectedIndex].Fio);

            ListViewItem Age = new ListViewItem("�������");
            Age.SubItems.Add(Convert.ToString(Person.personlist[selectedIndex].Age));

            ListViewItem group = new ListViewItem("������");
            group.SubItems.Add(Person.personlist[selectedIndex].Group);

            ListViewItem Akanam = new ListViewItem("���������� �������");
            Akanam.Tag = Person.personlist[selectedIndex].Akanam;
            switch (Akanam.Tag)
            {
                case "0":
                    Akanam.SubItems.Add("�� ��������");
                    break;
                case "1":
                    Akanam.SubItems.Add("�������� 1 �����������");
                    break;
                case "2":
                    Akanam.SubItems.Add("������� ����������� ���������");
                    break;
                default:
                    Akanam.SubItems.Add("is it works?");
                    break;
            }

            ListViewItem Somanam = new ListViewItem("������������ �������");
            Somanam.Tag = Person.personlist[selectedIndex].Somanam;
            switch (Somanam.Tag)
            {
                case "0":
                    Somanam.SubItems.Add("�� ��������");
                    break;
                case "1":
                    Somanam.SubItems.Add("�������� 1 �����������");
                    break;
                case "2":
                    Somanam.SubItems.Add("������� ����������� ���������");
                    break;
            }


            ListViewItem PregnancythreatStatus = new ListViewItem("������ ���������� ������������");
            PregnancythreatStatus.Tag = Person.personlist[selectedIndex].PregnancythreatStatus;
            switch (PregnancythreatStatus.Tag)
            {
                case "0":
                    PregnancythreatStatus.SubItems.Add("���");
                    break;
                case "1":
                    PregnancythreatStatus.SubItems.Add("������������� �������");
                    break;
                case "2":
                    PregnancythreatStatus.SubItems.Add("��������� ������������� ������������");
                    break;
            }


            ListViewItem EdemaStatus = new ListViewItem("�����, ��������� ������������� ���������� ��� ��������� ������������� �����������");
            EdemaStatus.Tag = Person.personlist[selectedIndex].EdemaStatus;
            switch (EdemaStatus.Tag)
            {
                case "0":
                    EdemaStatus.SubItems.Add("���");
                    break;
                case "1":
                    EdemaStatus.SubItems.Add("���� ������ ����,��������� �������������");
                    break;
                case "2":
                    EdemaStatus.SubItems.Add("����� � ��������� � ������������� � ��� ������������");
                    break;
            }


            ListViewItem PreeclampsiaStatus = new ListViewItem("������������");
            PreeclampsiaStatus.Tag = Person.personlist[selectedIndex].PreeclampsiaStatus;
            switch (PreeclampsiaStatus.Tag)
            {
                case "0":
                    PreeclampsiaStatus.SubItems.Add("���");
                    break;
                case "1":
                    PreeclampsiaStatus.SubItems.Add("���������");
                    break;
                case "2":
                    PreeclampsiaStatus.SubItems.Add("������");
                    break;
            }


            ListViewItem AnemiaStatus = new ListViewItem("������ �� ����� ������������");
            AnemiaStatus.Tag = Person.personlist[selectedIndex].AnemiaStatus;
            switch (AnemiaStatus.Tag)
            {
                case "0":
                    AnemiaStatus.SubItems.Add("���");
                    break;
                case "1":
                    AnemiaStatus.SubItems.Add("1 �������");
                    break;
                case "2":
                    AnemiaStatus.SubItems.Add("2 � ����� �������");
                    break;
            }


            ListViewItem InfectionStatus = new ListViewItem("�������� �/��� TORCH-��������");
            InfectionStatus.Tag = Person.personlist[selectedIndex].InfectionStatus;
            switch (InfectionStatus.Tag)
            {
                case "0":
                    InfectionStatus.SubItems.Add("���");
                    break;
                case "1":
                    InfectionStatus.SubItems.Add("���� 1 � ���������� ����");
                    break;
                case "2":
                    InfectionStatus.SubItems.Add("� �������� ����");
                    break;
            }


            ListViewItem GravidogramStatus = new ListViewItem("�������������");
            GravidogramStatus.Tag = Person.personlist[selectedIndex].GravidogramStatus;
            switch (GravidogramStatus.Tag)
            {
                case "0":
                    GravidogramStatus.SubItems.Add("������������� ����� ������������");
                    break;
                case "1":
                    GravidogramStatus.SubItems.Add("������� �� ����� ������������ �� 1-2 ������");
                    break;
                case "2":
                    GravidogramStatus.SubItems.Add("������� �� ����� ������������ ����� ��� �� 2 ������");
                    break;
            }


            ListViewItem BloodflowStatus = new ListViewItem("��������� �������-������������� ���������");
            BloodflowStatus.Tag = Person.personlist[selectedIndex].BloodflowStatus;
            switch (BloodflowStatus.Tag)
            {
                case "0":
                    BloodflowStatus.SubItems.Add("���");
                    break;
                case "1":
                    BloodflowStatus.SubItems.Add("����������������");
                    break;
                case "2":
                    BloodflowStatus.SubItems.Add("������������������");
                    break;
            }


            ListViewItem FetusStatus = new ListViewItem("�������������� ��������� ��������� ����� � ������������� ��������, ���������� �� ���");
            FetusStatus.Tag = Person.personlist[selectedIndex].FetusStatus;
            switch (FetusStatus.Tag)
            {
                case "0":
                    FetusStatus.SubItems.Add("���");
                    break;
                case "1":
                    FetusStatus.SubItems.Add("����������������");
                    break;
                case "2":
                    FetusStatus.SubItems.Add("������������������");
                    break;
            }


            ListViewItem FeverStatus = new ListViewItem("���������");
            FeverStatus.Tag = Person.personlist[selectedIndex].FeverStatus;
            switch (FeverStatus.Tag)
            {
                case "0":
                    FeverStatus.SubItems.Add("���");
                    break;
                case "1":
                    FeverStatus.SubItems.Add("�� 3-� �����");
                    break;
                case "2":
                    FeverStatus.SubItems.Add("3 � ����� �����");
                    break;
            }


            ListViewItem CoughStatus = new ListViewItem("������");
            CoughStatus.Tag = Person.personlist[selectedIndex].CoughStatus;
            switch (CoughStatus.Tag)
            {
                case "0":
                    CoughStatus.SubItems.Add("���");
                    break;
                case "1":
                    CoughStatus.SubItems.Add("�� 5 �����");
                    break;
                case "2":
                    CoughStatus.SubItems.Add("���������� ������� ������ � ����������� �������");
                    break;
            }


            ListViewItem DyspneaStatus = new ListViewItem("������");
            DyspneaStatus.Tag = Person.personlist[selectedIndex].DyspneaStatus;
            switch (DyspneaStatus.Tag)
            {
                case "0":
                    DyspneaStatus.SubItems.Add("���");
                    break;
                case "1":
                    DyspneaStatus.SubItems.Add("��� ���������� ��������");
                    break;
                case "2":
                    DyspneaStatus.SubItems.Add("� �����");
                    break;
            }


            ListViewItem HemodynamicsStatus = new ListViewItem("��������� ������������");
            HemodynamicsStatus.Tag = Person.personlist[selectedIndex].HemodynamicsStatus;
            switch (HemodynamicsStatus.Tag)
            {
                case "0":
                    HemodynamicsStatus.SubItems.Add("���");
                    break;
                case "1":
                    HemodynamicsStatus.SubItems.Add("��� ���������� ��������");
                    break;
                case "2":
                    HemodynamicsStatus.SubItems.Add("� �����");
                    break;
            }


            ListViewItem NauseaStatus = new ListViewItem("�������, �����, ������");
            NauseaStatus.Tag = Person.personlist[selectedIndex].NauseaStatus;
            switch (NauseaStatus.Tag)
            {
                case "0":
                    NauseaStatus.SubItems.Add("���");
                    break;
                case "1":
                    NauseaStatus.SubItems.Add("1-2 ���� � �����");
                    break;
                case "2":
                    NauseaStatus.SubItems.Add("����� 2 ��� � �����");
                    break;
            }


            ListViewItem SaturationStatus = new ListViewItem("���������");
            SaturationStatus.Tag = Person.personlist[selectedIndex].SaturationStatus;
            switch (SaturationStatus.Tag)
            {
                case "0":
                    SaturationStatus.SubItems.Add("���� 95");
                    break;
                case "1":
                    SaturationStatus.SubItems.Add("92-95");
                    break;
                case "2":
                    SaturationStatus.SubItems.Add("����� 92");
                    break;
            }


            ListViewItem LungtissueStatus = new ListViewItem("��������� �������� ����� ��� ���������� ��");
            LungtissueStatus.Tag = Person.personlist[selectedIndex].LungtissueStatus;
            switch (LungtissueStatus.Tag)
            {
                case "0":
                    LungtissueStatus.SubItems.Add("���");
                    break;
                case "1":
                    LungtissueStatus.SubItems.Add("�� 1");
                    break;
                case "2":
                    LungtissueStatus.SubItems.Add("�� 2 � ����");
                    break;
            }

            PersonView1.Items.Add(Fio);
            PersonView1.Items.Add(Age);
            PersonView1.Items.Add(group);
            PersonView1.Items.Add(Akanam);
            PersonView1.Items.Add(Somanam);
            PersonView1.Items.Add(PregnancythreatStatus);
            PersonView1.Items.Add(EdemaStatus);
            PersonView1.Items.Add(PreeclampsiaStatus);
            PersonView1.Items.Add(AnemiaStatus);
            PersonView1.Items.Add(InfectionStatus);
            PersonView1.Items.Add(GravidogramStatus);
            PersonView1.Items.Add(BloodflowStatus);
            PersonView1.Items.Add(FetusStatus);
            PersonView1.Items.Add(FeverStatus);
            PersonView1.Items.Add(CoughStatus);
            PersonView1.Items.Add(DyspneaStatus);
            PersonView1.Items.Add(HemodynamicsStatus);
            PersonView1.Items.Add(NauseaStatus);
            PersonView1.Items.Add(SaturationStatus);

        }

        public void AddButton_Click(object sender, EventArgs e)
        {



            AddPersonForm addPersonForm = new AddPersonForm();

            addPersonForm.ShowDialog();

            PersonListRefresh();

        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            int selectedPerson = PersonList.SelectedIndices[0];
            DialogResult result = MessageBox.Show("�� �������, ��� ������ ������� ������?", "������������� ��������", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                PersonList.Items[selectedPerson].Remove();
                Person.personlist.RemoveAt(selectedPerson);
                PersonView1.Items.Clear();
            }

        }

        private void SaveAsMenuButton_Click(object sender, EventArgs e)
        {
            ExcelHelper.SaveAs();
        }

        private void SaveMenuButton_Click(object sender, EventArgs e)
        {
            ExcelHelper.Save(_filePath);
        }
    }



}