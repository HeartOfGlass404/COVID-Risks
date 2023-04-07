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

            ListViewItem Fio = new ListViewItem("ФИО");
            Fio.SubItems.Add(Person.personlist[selectedIndex].Fio);

            ListViewItem Age = new ListViewItem("Возраст");
            Age.SubItems.Add(Convert.ToString(Person.personlist[selectedIndex].Age));

            ListViewItem group = new ListViewItem("Группа");
            group.SubItems.Add(Person.personlist[selectedIndex].Group);

            ListViewItem Akanam = new ListViewItem("Акушерский анамнез");
            Akanam.Tag = Person.personlist[selectedIndex].Akanam;
            switch (Akanam.Tag)
            {
                case "0":
                    Akanam.SubItems.Add("не отягощен");
                    break;
                case "1":
                    Akanam.SubItems.Add("отягощен 1 осложнением");
                    break;
                case "2":
                    Akanam.SubItems.Add("имеются коморбидные состояния");
                    break;
                default:
                    Akanam.SubItems.Add("is it works?");
                    break;
            }

            ListViewItem Somanam = new ListViewItem("Соматический анамнез");
            Somanam.Tag = Person.personlist[selectedIndex].Somanam;
            switch (Somanam.Tag)
            {
                case "0":
                    Somanam.SubItems.Add("не отягощен");
                    break;
                case "1":
                    Somanam.SubItems.Add("отягощен 1 осложнением");
                    break;
                case "2":
                    Somanam.SubItems.Add("имеются коморбидные состояния");
                    break;
            }


            ListViewItem PregnancythreatStatus = new ListViewItem("Угроза прерывания беременности");
            PregnancythreatStatus.Tag = Person.personlist[selectedIndex].PregnancythreatStatus;
            switch (PregnancythreatStatus.Tag)
            {
                case "0":
                    PregnancythreatStatus.SubItems.Add("нет");
                    break;
                case "1":
                    PregnancythreatStatus.SubItems.Add("спорадический выкидыш");
                    break;
                case "2":
                    PregnancythreatStatus.SubItems.Add("привычное невынашивание беременности");
                    break;
            }


            ListViewItem EdemaStatus = new ListViewItem("Отеки, вызванный беременностью протеинури или вызванная беременностью гипертензия");
            EdemaStatus.Tag = Person.personlist[selectedIndex].EdemaStatus;
            switch (EdemaStatus.Tag)
            {
                case "0":
                    EdemaStatus.SubItems.Add("нет");
                    break;
                case "1":
                    EdemaStatus.SubItems.Add("есть только отёки,вызванные беременностью");
                    break;
                case "2":
                    EdemaStatus.SubItems.Add("отеки в сочетании с спротеинурией и или гипертензией");
                    break;
            }


            ListViewItem PreeclampsiaStatus = new ListViewItem("Преэклампсия");
            PreeclampsiaStatus.Tag = Person.personlist[selectedIndex].PreeclampsiaStatus;
            switch (PreeclampsiaStatus.Tag)
            {
                case "0":
                    PreeclampsiaStatus.SubItems.Add("нет");
                    break;
                case "1":
                    PreeclampsiaStatus.SubItems.Add("умеренная");
                    break;
                case "2":
                    PreeclampsiaStatus.SubItems.Add("тяжёлая");
                    break;
            }


            ListViewItem AnemiaStatus = new ListViewItem("Анемия во время беременности");
            AnemiaStatus.Tag = Person.personlist[selectedIndex].AnemiaStatus;
            switch (AnemiaStatus.Tag)
            {
                case "0":
                    AnemiaStatus.SubItems.Add("нет");
                    break;
                case "1":
                    AnemiaStatus.SubItems.Add("1 степень");
                    break;
                case "2":
                    AnemiaStatus.SubItems.Add("2 и более степень");
                    break;
            }


            ListViewItem InfectionStatus = new ListViewItem("Вирусные и/или TORCH-инфекция");
            InfectionStatus.Tag = Person.personlist[selectedIndex].InfectionStatus;
            switch (InfectionStatus.Tag)
            {
                case "0":
                    InfectionStatus.SubItems.Add("нет");
                    break;
                case "1":
                    InfectionStatus.SubItems.Add("есть 1 в неактивной фазе");
                    break;
                case "2":
                    InfectionStatus.SubItems.Add("в активной фазе");
                    break;
            }


            ListViewItem GravidogramStatus = new ListViewItem("Гравидограмма");
            GravidogramStatus.Tag = Person.personlist[selectedIndex].GravidogramStatus;
            switch (GravidogramStatus.Tag)
            {
                case "0":
                    GravidogramStatus.SubItems.Add("соответствует сроку беременности");
                    break;
                case "1":
                    GravidogramStatus.SubItems.Add("отстает от срока беременности на 1-2 недели");
                    break;
                case "2":
                    GravidogramStatus.SubItems.Add("отстает от срока беременности более чем на 2 недели");
                    break;
            }


            ListViewItem BloodflowStatus = new ListViewItem("Нарушения маточно-плацентарного кровотока");
            BloodflowStatus.Tag = Person.personlist[selectedIndex].BloodflowStatus;
            switch (BloodflowStatus.Tag)
            {
                case "0":
                    BloodflowStatus.SubItems.Add("нет");
                    break;
                case "1":
                    BloodflowStatus.SubItems.Add("компенсированные");
                    break;
                case "2":
                    BloodflowStatus.SubItems.Add("декомпенсированные");
                    break;
            }


            ListViewItem FetusStatus = new ListViewItem("Патологические изменения состояния плода и парафетальных структур, выявленные на УЗИ");
            FetusStatus.Tag = Person.personlist[selectedIndex].FetusStatus;
            switch (FetusStatus.Tag)
            {
                case "0":
                    FetusStatus.SubItems.Add("нет");
                    break;
                case "1":
                    FetusStatus.SubItems.Add("компенсированные");
                    break;
                case "2":
                    FetusStatus.SubItems.Add("декомпенсированные");
                    break;
            }


            ListViewItem FeverStatus = new ListViewItem("Лихорадка");
            FeverStatus.Tag = Person.personlist[selectedIndex].FeverStatus;
            switch (FeverStatus.Tag)
            {
                case "0":
                    FeverStatus.SubItems.Add("нет");
                    break;
                case "1":
                    FeverStatus.SubItems.Add("до 3-х суток");
                    break;
                case "2":
                    FeverStatus.SubItems.Add("3 и более суток");
                    break;
            }


            ListViewItem CoughStatus = new ListViewItem("Кашель");
            CoughStatus.Tag = Person.personlist[selectedIndex].CoughStatus;
            switch (CoughStatus.Tag)
            {
                case "0":
                    CoughStatus.SubItems.Add("нет");
                    break;
                case "1":
                    CoughStatus.SubItems.Add("до 5 суток");
                    break;
                case "2":
                    CoughStatus.SubItems.Add("постоянный сильный кашель с отхождением мокроты");
                    break;
            }


            ListViewItem DyspneaStatus = new ListViewItem("Одышка");
            DyspneaStatus.Tag = Person.personlist[selectedIndex].DyspneaStatus;
            switch (DyspneaStatus.Tag)
            {
                case "0":
                    DyspneaStatus.SubItems.Add("нет");
                    break;
                case "1":
                    DyspneaStatus.SubItems.Add("при физической нагрузке");
                    break;
                case "2":
                    DyspneaStatus.SubItems.Add("в покое");
                    break;
            }


            ListViewItem HemodynamicsStatus = new ListViewItem("Изменения гемодинамики");
            HemodynamicsStatus.Tag = Person.personlist[selectedIndex].HemodynamicsStatus;
            switch (HemodynamicsStatus.Tag)
            {
                case "0":
                    HemodynamicsStatus.SubItems.Add("нет");
                    break;
                case "1":
                    HemodynamicsStatus.SubItems.Add("при физической нагрузке");
                    break;
                case "2":
                    HemodynamicsStatus.SubItems.Add("в покое");
                    break;
            }


            ListViewItem NauseaStatus = new ListViewItem("Тошнота, рвота, диарея");
            NauseaStatus.Tag = Person.personlist[selectedIndex].NauseaStatus;
            switch (NauseaStatus.Tag)
            {
                case "0":
                    NauseaStatus.SubItems.Add("нет");
                    break;
                case "1":
                    NauseaStatus.SubItems.Add("1-2 раза в сутки");
                    break;
                case "2":
                    NauseaStatus.SubItems.Add("более 2 раз в сутки");
                    break;
            }


            ListViewItem SaturationStatus = new ListViewItem("Сатурация");
            SaturationStatus.Tag = Person.personlist[selectedIndex].SaturationStatus;
            switch (SaturationStatus.Tag)
            {
                case "0":
                    SaturationStatus.SubItems.Add("выше 95");
                    break;
                case "1":
                    SaturationStatus.SubItems.Add("92-95");
                    break;
                case "2":
                    SaturationStatus.SubItems.Add("менее 92");
                    break;
            }


            ListViewItem LungtissueStatus = new ListViewItem("Изменения легочной ткани при выполнении КТ");
            LungtissueStatus.Tag = Person.personlist[selectedIndex].LungtissueStatus;
            switch (LungtissueStatus.Tag)
            {
                case "0":
                    LungtissueStatus.SubItems.Add("нет");
                    break;
                case "1":
                    LungtissueStatus.SubItems.Add("КТ 1");
                    break;
                case "2":
                    LungtissueStatus.SubItems.Add("КТ 2 и выше");
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
            DialogResult result = MessageBox.Show("Вы уверены, что хотите удалить запись?", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
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