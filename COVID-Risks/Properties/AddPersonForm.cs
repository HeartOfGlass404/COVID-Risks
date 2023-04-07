using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WinFormsApp1.Models;

namespace COVID_Risks.Properties
{
    public partial class AddPersonForm : Form
    {
        string akanam = "";
        string somanam = "";
        string pregnancyThreatStatus = "";
        string edemaStatus = "";
        string preeclampsiaStatus = "";
        string anemiaStatus = "";
        string infectionStatus = "";
        string gravidogramStatus = "";
        string bloodflowStatus = "";
        string fetusStatus = "";
        string feverStatus = "";
        string coughStatus = "";
        string dyspneaStatus = "";
        string hemodynamicsStatus = "";
        string saturationStatus = "";
        string nauseaStatus = "";
        string lungtissueStatus = "";
        public AddPersonForm()
        {
            InitializeComponent();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы уверены, что хотите отменить добавление записи?",
                "Подтверждение", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                this.Close();
            }

        }


        public void OkButton_Click(object sender, EventArgs e)
        {

            string fio = fioBox.Text;
            int age = Convert.ToInt16(ageBox.Text);
            string group = "";
            float pregCoef = (Convert.ToByte(pregnancyThreatStatus) + Convert.ToByte(edemaStatus) +
                Convert.ToByte(preeclampsiaStatus) + Convert.ToByte(anemiaStatus) + Convert.ToByte(infectionStatus) +
                Convert.ToByte(gravidogramStatus) + Convert.ToByte(bloodflowStatus) + Convert.ToByte(fetusStatus)) / 8;
            float sarsCov2 = (Convert.ToByte(feverStatus) + Convert.ToByte(coughStatus) + Convert.ToByte(dyspneaStatus) +
                Convert.ToByte(hemodynamicsStatus) + Convert.ToByte(nauseaStatus) + Convert.ToByte(saturationStatus) +
                Convert.ToByte(lungtissueStatus)) / 7;
            float fkrpo = Convert.ToByte(akanam) + Convert.ToByte(somanam) + pregCoef + sarsCov2;
            if (fkrpo <= 3.0)
            {
                group = "низкий риск перинататльных осложнений";
            }
            else if ((fkrpo >= 3.1) && (fkrpo >= 5.0))
            {
                group = "средний риск перинатальных осложнений";
            }
            else if (fkrpo >= 5.1)
            {
                group = "высокий риск перинатальных осложнений";
            }




            Person tmpPerson = new Person(fio, age, group, akanam, somanam, pregnancyThreatStatus, edemaStatus,
             preeclampsiaStatus, anemiaStatus, infectionStatus, gravidogramStatus,
             bloodflowStatus, fetusStatus, feverStatus, coughStatus, dyspneaStatus, hemodynamicsStatus, saturationStatus,
             nauseaStatus, lungtissueStatus);

            bool exists = Person.personlist.Any(p => p.Fio == fio);
            int index = Person.personlist.FindIndex(x => x.Fio == fio);
            if (exists)
            {
                Person.personlist.RemoveAt(index);
                Person.personlist[index] = tmpPerson;
            }
            int c = Person.personlist.Count;
            Debug.WriteLine(c);
            Person.personlist.Add(tmpPerson);
            this.Close();
            c = Person.personlist.Count;
            var li = Person.personlist.Last();
            Debug.WriteLine(li.Fio);
            Debug.WriteLine(c);

        }


        private void ageBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        private void akanamBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            akanam = akanamBox.GetItemText(akanamBox.SelectedItem);
            switch (akanam)
            {
                case "не отягощен":
                    akanam = "0";
                    break;
                case "отягощен 1 осложнением":
                    akanam = "1";
                    break;
                case "имеются коморбидные состояния":
                    akanam = "2";
                    break;
            }
        }

        private void somanamBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            somanam = somanamBox.GetItemText(somanamBox.SelectedItem);
            switch (somanam)
            {
                case "не отягощен":
                    somanam = "0";
                    break;
                case "отягощен 1 осложнением":
                    somanam = "1";
                    break;
                case "имеются коморбидные состояния":
                    somanam = "2";
                    break;
                default:
                    somanam = "неизвестно";
                    break;
            }
        }

        private void pregnancyThreatBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            pregnancyThreatStatus = pregnancyThreatBox.GetItemText(pregnancyThreatBox.SelectedItem);
            switch (pregnancyThreatStatus)
            {
                case "нет":
                    pregnancyThreatStatus = "0";
                    break;
                case "спорадический выкидыш":
                    pregnancyThreatStatus = "1";
                    break;
                case "привычное невынашивание беременности":
                    pregnancyThreatStatus = "2";
                    break;
                default:
                    pregnancyThreatStatus = "неизвестно";
                    break;
            }
        }

        private void edemaBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            edemaStatus = edemaBox.GetItemText(edemaBox.SelectedItem);
            switch (edemaStatus)
            {
                case "нет":
                    edemaStatus = "0";
                    break;
                case "есть только отеки, вызванные беременностью":
                    edemaStatus = "1";
                    break;
                case "отеки в сочетании с спротеинурией и\\или гипертензией":
                    edemaStatus = "2";
                    break;
            }
        }

        private void preeclampsiaBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            preeclampsiaStatus = preeclampsiaBox.GetItemText(preeclampsiaBox.SelectedItem);
            switch (preeclampsiaStatus)
            {
                case "нет":
                    preeclampsiaStatus = "0";
                    break;
                case "умеренная":
                    preeclampsiaStatus = "1";
                    break;
                case "тяжелая":
                    preeclampsiaStatus = "2";
                    break;
            }
        }

        private void anemiaBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            anemiaStatus = anemiaBox.GetItemText(anemiaBox.SelectedItem);
            switch (anemiaStatus)
            {
                case "нет":
                    anemiaStatus = "0";
                    break;
                case "1 степень":
                    anemiaStatus = "1";
                    break;
                case "2 и более степень":
                    anemiaStatus = "2";
                    break;
            }
        }

        private void infectionBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            infectionStatus = infectionBox.GetItemText(infectionBox.SelectedItem);
            switch (infectionStatus)
            {
                case "нет":
                    infectionStatus = "0";
                    break;
                case "есть 1 в неактивной фазе":
                    infectionStatus = "1";
                    break;
                case "в активной фазе":
                    infectionStatus = "2";
                    break;
            }
        }

        private void gravidogrammBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            gravidogramStatus = gravidogrammBox.GetItemText(gravidogrammBox.SelectedItem);
            switch (gravidogramStatus)
            {
                case "соответствует сроку беременности":
                    gravidogramStatus = "0";
                    break;
                case "отстает от срока беременности на 1-2 недели":
                    gravidogramStatus = "1";
                    break;
                case "отстает от срока беременности более чем на 2 недели":
                    gravidogramStatus = "2";
                    break;
            }
        }

        private void bloodflowBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            bloodflowStatus = bloodflowBox.GetItemText(bloodflowBox.SelectedItem);
            switch (bloodflowStatus)
            {
                case "нет":
                    bloodflowStatus = "0";
                    break;
                case "компенсированные":
                    bloodflowStatus = "1";
                    break;
                case "декомпенсированные":
                    bloodflowStatus = "2";
                    break;
                default:
                    bloodflowStatus = "неизвестно";
                    break;
            }
        }

        private void fetusBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            fetusStatus = fetusBox.GetItemText(fetusBox.SelectedItem);
            switch (fetusStatus)
            {
                case "нет":
                    fetusStatus = "0";
                    break;
                case "компенсированные":
                    fetusStatus = "1";
                    break;
                case "декомпенсированные":
                    fetusStatus = "2";
                    break;
            }
        }

        private void feverBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            feverStatus = feverBox.GetItemText(feverBox.SelectedItem);
            switch (feverStatus)
            {
                case "нет":
                    feverStatus = "0";
                    break;
                case "до 3-х суток":
                    feverStatus = "1";
                    break;
                case "3 и более суток":
                    feverStatus = "2";
                    break;
            }
        }

        private void coughBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            coughStatus = coughBox.GetItemText(coughBox.SelectedItem);
            switch (coughStatus)
            {
                case "нет":
                    coughStatus = "0";
                    break;
                case "до 5 суток":
                    coughStatus = "1";
                    break;
                case "постоянный сильный кашель с отхождением мокроты":
                    coughStatus = "2";
                    break;
            }
        }

        private void dyspneaBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            dyspneaStatus = dyspneaBox.GetItemText(dyspneaBox.SelectedItem);
            switch (dyspneaStatus)
            {
                case "нет":
                    dyspneaStatus = "0";
                    break;
                case "при физической нагрузке":
                    dyspneaStatus = "1";
                    break;
                case " в покое":
                    dyspneaStatus = "2";
                    break;
            }
        }

        private void hemodynamicsBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            hemodynamicsStatus = hemodynamicsBox.GetItemText(hemodynamicsBox.SelectedItem);
            switch (hemodynamicsStatus)
            {
                case "нет":
                    hemodynamicsStatus = "0";
                    break;
                case "при физической нагрузке":
                    hemodynamicsStatus = "1";
                    break;
                case "в покое":
                    hemodynamicsStatus = "2";
                    break;
            }
        }

        private void nauseaBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            nauseaStatus = nauseaBox.GetItemText(nauseaBox.SelectedItem);
            switch (nauseaStatus)
            {
                case "нет":
                    nauseaStatus = "0";
                    break;
                case "1-2 раза в сутки":
                    nauseaStatus = "1";
                    break;
                case "более 2 раз в сутки":
                    nauseaStatus = "2";
                    break;
            }
        }

        private void saturationBx_SelectionChangeCommitted(object sender, EventArgs e)
        {
            saturationStatus = saturationBx.GetItemText(saturationBx.SelectedItem);
            switch (saturationStatus)
            {
                case "выше 95":
                    saturationStatus = "0";
                    break;
                case "92-95":
                    saturationStatus = "1";
                    break;
                case "менее 92":
                    saturationStatus = "2";
                    break;
            }
        }

        private void lungTissueBox_SelectionChangeCommitted(object sender, EventArgs e)
        {
            lungtissueStatus = lungTissueBox.GetItemText(lungTissueBox.SelectedItem);
            switch (lungtissueStatus)
            {
                case "нет":
                    lungtissueStatus = "0";
                    break;
                case "КТ 1":
                    lungtissueStatus = "1";
                    break;
                case "КТ 2 и выше":
                    lungtissueStatus = "2";
                    break;
            }
        }

        private void ageBox_Enter(object sender, EventArgs e)
        {
            if (ageBox.Text == "Введите возраст")
            {
                ageBox.Text = "";
            }
        }

        private void ageBox_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(ageBox.Text))
            {
                ageBox.Text = "Введите возраст";
            }
        }
    }
}
