using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using WinFormsApp1.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace COVID_Risks.Models
{
    class ExcelHelper
    {
        static string? _filePath;


        public static void SaveAs()
        {
            Excel.Application excel = null;
            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;
            Excel.Sheets worksheets = null;
            Excel.Worksheet worksheet = null;


            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel files|*.xlsx";
            Nullable<bool> dialogOK = Convert.ToBoolean(saveFileDialog.ShowDialog());
            if (dialogOK == true)
            {
                _filePath = saveFileDialog.FileName;
            }

            try
            {
                excel = new Excel.Application();
                
                workbooks = excel.Workbooks;
                workbook = workbooks.Add();
                worksheets = workbook.Worksheets;
                worksheet = worksheets.Item[1];

                Excel.Range range = excel.get_Range("A1", "BB3");
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;


                excel.Visible = true;
                excel.DisplayAlerts = false;
                excel.Range["D1"].Value = "Акушерский анамнез";
                excel.Range["D1", "F1"].Merge();
                excel.Range["G1"].Value = "Соматический анамнез";
                excel.Range["G1", "I1"].Merge();
                excel.Range["J1"].Value = "Течение беременности";
                excel.Range["J1", "AG1"].Merge();
                excel.Range["AH1"].Value = "Течение SARS-CoV-2";
                excel.Range["AH1", "BB1"].Merge();
                excel.Range["A2"].Value = "Пациент";
                excel.Range["B2"].Value = "Наименование группы";
                excel.Range["C2"].Value = "Возраст";
                excel.Range["D2"].Value = "не отягощен";
                excel.Range["E2"].Value = "отягощен 1 осложнением";
                excel.Range["F2"].Value = "имеются коморбидные состояния";
                excel.Range["G2"].Value = "не отягощен";
                excel.Range["H2"].Value = "отягощен 1 осложнением";
                excel.Range["I2"].Value = "имеются коморбидные состояния";
                excel.Range["J2"].Value = "Угроза прерывания беременности";
                excel.Range["J2", "L2"].Merge();
                excel.Range["M2"].Value = "Отеки, вызванная беременностью протеинурия или вызванная беременностью гипертензия";
                excel.Range["M2", "O2"].Merge();
                excel.Range["P2"].Value = "Преэклампсия";
                excel.Range["P2", "R2"].Merge();
                excel.Range["S2"].Value = "Анемия во время беременности";
                excel.Range["S2", "U2"].Merge();
                excel.Range["V2"].Value = "Вирусные и/или TORCH-инфекция";
                excel.Range["V2", "X2"].Merge();
                excel.Range["Y2"].Value = "Гравидограмма";
                excel.Range["Y2", "AA2"].Merge();
                excel.Range["AB2"].Value = "Нарушения маточно-плацентарного кровотока";
                excel.Range["AB2", "AD2"].Merge();
                excel.Range["AE2"].Value = "Патологические изменения состояния плода и парафетальных структур, выявленные на УЗИ";
                excel.Range["AE2", "AG2"].Merge();
                excel.Range["AH2"].Value = "Лихорадка";
                excel.Range["AH2", "AJ2"].Merge();
                excel.Range["AK2"].Value = "Кашель";
                excel.Range["AK2", "AM2"].Merge();
                excel.Range["AN2"].Value = "Одышка";
                excel.Range["AN2", "AP2"].Merge();
                excel.Range["AQ2"].Value = "Изменения гемодинамики";
                excel.Range["AQ2", "AS2"].Merge();
                excel.Range["AT2"].Value = "Тошнота, рвота, диарея";
                excel.Range["AT2", "AV2"].Merge();
                excel.Range["AW2"].Value = "Сатурация";
                excel.Range["AW2", "AY2"].Merge();
                excel.Range["AZ2"].Value = "Изменения легочной ткани при выполнении КТ";
                excel.Range["AZ2", "BB2"].Merge();
                excel.Range["J3"].Value = "нет";
                excel.Range["K3"].Value = "спорадический выкидыш";
                excel.Range["L3"].Value = "привычное невынашивание беременности";
                excel.Range["M3"].Value = "нет";
                excel.Range["N3"].Value = "есть только отеки, вызванные беременностью";
                excel.Range["O3"].Value = "отеки в сочетании с спротеинурией и\\или гипертензией";
                excel.Range["P3"].Value = "нет";
                excel.Range["Q3"].Value = "умеренная";
                excel.Range["R3"].Value = "тяжелая";
                excel.Range["S3"].Value = "нет";
                excel.Range["T3"].Value = "1 степень";
                excel.Range["U3"].Value = "2 и более степень";
                excel.Range["V3"].Value = "нет";
                excel.Range["W3"].Value = "есть 1 в неактивной фазе";
                excel.Range["X3"].Value = "в активной фазе";
                excel.Range["Y3"].Value = "соответствует сроку беременности";
                excel.Range["Z3"].Value = "отстает от срока беременности на 1-2 недели";
                excel.Range["AA3"].Value = "отстает от срока беременности более чем на 2 недели";
                excel.Range["AB3"].Value = "нет";
                excel.Range["AC3"].Value = "компенсированные";
                excel.Range["AD3"].Value = "декомпенсированные";
                excel.Range["AE3"].Value = "нет";
                excel.Range["AF3"].Value = "компенсированные";
                excel.Range["AG3"].Value = "декомпенсированные";
                excel.Range["AH3"].Value = "нет";
                excel.Range["AI3"].Value = "до 3-х суток";
                excel.Range["AJ3"].Value = "3 и более суток";
                excel.Range["AK3"].Value = "нет";
                excel.Range["AL3"].Value = "до 5 суток";
                excel.Range["AM3"].Value = "постоянный сильный кашель с отхождением мокроты";
                excel.Range["AN3"].Value = "нет";
                excel.Range["AO3"].Value = "при физической нагрузке";
                excel.Range["AP3"].Value = "в покое";
                excel.Range["AQ3"].Value = "нет";
                excel.Range["AR3"].Value = "при физической нагрузке";
                excel.Range["AS3"].Value = "в покое";
                excel.Range["AT3"].Value = "нет";
                excel.Range["AU3"].Value = "1-2 раза в сутки";
                excel.Range["AV3"].Value = "более 2 раз в сутки";
                excel.Range["AW3"].Value = "выше 95";
                excel.Range["AX3"].Value = "92-95";
                excel.Range["AY3"].Value = "менее 92";
                excel.Range["AZ3"].Value = "нет";
                excel.Range["BA3"].Value = "КТ 1";
                excel.Range["BB3"].Value = "КТ 2 и выше";

               
                
                

                range = excel.get_Range("D1", "F1");
                range.Interior.Color = System.Drawing.Color.FromArgb(198, 239, 206);
                range = excel.get_Range("G1", "I1");
                range.Interior.Color = System.Drawing.Color.FromArgb(255, 235, 156);
                range = excel.get_Range("J1", "AG1");
                range.Interior.Color = System.Drawing.Color.FromArgb(180, 198, 231);
                range = excel.get_Range("AH1", "BB1");
                range.Interior.Color = System.Drawing.Color.FromArgb(255, 199, 206);


                if (Person.personlist.Count > 0)
                {
                    foreach (Person person in Person.personlist)
                    {
                        for (int i = 4; i < Person.personlist.Count;)
                        {
                            for (int j = 1; j < Person.personlist.Count; j++)
                            {
                                excel.Range["A" + i].Value = Person.personlist[j].Fio;
                                Debug.WriteLine(person.Fio);
                                excel.Range["B" + i].Value = Person.personlist[j].Group;
                                excel.Range["C" + i].Value = Person.personlist[j].Age;
                                switch (Person.personlist[j].Akanam)
                                {
                                    case "0":
                                        excel.Range["D" + i].Value = 0;
                                        break;
                                    case "1":
                                        excel.Range["E" + i].Value = 1;
                                        break;
                                    case "2":
                                        excel.Range["F" + i].Value = 2;
                                        break;
                                }
                                switch (Person.personlist[j].Somanam)
                                {
                                    case "0":
                                        excel.Range["G" + i].Value = 0;
                                        break;
                                    case "1":
                                        excel.Range["H" + i].Value = 1;
                                        break;
                                    case "2":
                                        excel.Range["I" + i].Value = 2;
                                        break;
                                }
                                switch (Person.personlist[j].PregnancythreatStatus)
                                {
                                    case "0":
                                        excel.Range["J" + i].Value = 0;
                                        break;
                                    case "1":
                                        excel.Range["K" + i].Value = 1;
                                        break;
                                    case "2":
                                        excel.Range["L" + i].Value = 2;
                                        break;
                                }
                                switch (Person.personlist[j].EdemaStatus)
                                {
                                    case "0":
                                        excel.Range["M" + i].Value = 0;
                                        break;
                                    case "1":
                                        excel.Range["N" + 1].Value = 1;
                                        break;
                                    case "2":
                                        excel.Range["O" + 1].Value = 2;
                                        break;
                                }

                                i++;
                            }
                        }
                    }
                }
                else
                {

                }
                worksheet.UsedRange.Columns.AutoFit();
                worksheet.UsedRange.Rows.AutoFit();
            }
            finally
            {
                if (excel != null)
                {
                    workbook.SaveAs(_filePath);
                    
                    Marshal.ReleaseComObject(worksheet);
                    Marshal.ReleaseComObject(worksheets);
                    Marshal.ReleaseComObject(workbook);
                    Marshal.ReleaseComObject(workbooks);
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);
                }
            }
        }

        public static void Save(String _filePath)
        {
            Excel.Application excel = null;
            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;
            Excel.Sheets worksheets = null;
            Excel.Worksheet worksheet = null;

            try
            {
                
                excel = new Excel.Application();
                workbooks = excel.Workbooks;
                workbook = workbooks.Open(_filePath);

                if (Person.personlist.Count > 0)
                {
                    foreach (Person person in Person.personlist)
                    {
                        for (int i = 4; i < Person.personlist.Count;)
                        {
                            for (int j = 0; j < Person.personlist.Count - 1; j++)
                            {
                                excel.Range["A" + i].Value = Person.personlist[j].Fio;
                                excel.Range["B" + i].Value = Person.personlist[j].Group;
                                excel.Range["C" + i].Value = Person.personlist[j].Age;
                                switch (Person.personlist[j].Akanam)
                                {
                                    case "0":
                                        excel.Range["D" + i].Value = 0;
                                        break;
                                    case "1":
                                        excel.Range["E" + i].Value = 1;
                                        break;
                                    case "2":
                                        excel.Range["F" + i].Value = 2;
                                        break;
                                }
                                switch (Person.personlist[j].Somanam)
                                {
                                    case "0":
                                        excel.Range["G" + i].Value = 0;
                                        break;
                                    case "1":
                                        excel.Range["H" + i].Value = 1;
                                        break;
                                    case "2":
                                        excel.Range["I" + i].Value = 2;
                                        break;
                                }
                                switch (Person.personlist[j].PregnancythreatStatus)
                                {
                                    case "0":
                                        excel.Range["J" + i].Value = 0;
                                        break;
                                    case "1":
                                        excel.Range["K" + i].Value = 1;
                                        break;
                                    case "2":
                                        excel.Range["L" + i].Value = 2;
                                        break;
                                }
                                switch (Person.personlist[j].EdemaStatus)
                                {
                                    case "0":
                                        excel.Range["M" + i].Value = 0;
                                        break;
                                    case "1":
                                        excel.Range["N" + 1].Value = 1;
                                        break;
                                    case "2":
                                        excel.Range["O" + 1].Value = 2;
                                        break;
                                }

                                i++;
                            }
                        }
                    }
                }

            }
            finally
            {
                if (excel != null)
                {
                    workbook.Save();

                   
                    Marshal.ReleaseComObject(workbook);
                    Marshal.ReleaseComObject(workbooks);
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);
                }
            }
        }





        public static List<Person> GetPersonData(String _filePath)
        {
            Excel.Application excel = null;
            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;
            Excel.Sheets worksheets = null;
            Excel.Worksheet worksheet = null;

            excel = new Excel.Application();
            excel.DisplayAlerts = false;
            if (string.IsNullOrEmpty(_filePath))
            {
                excel.DisplayAlerts = true;
                MessageBox.Show("Вы не выбрали файл", "Внимание!",
                     MessageBoxButtons.OK, MessageBoxIcon.Error);
                List<Person> list = new List<Person>();
                return list;
            }
            workbooks = excel.Workbooks;
            workbook = workbooks.Open(_filePath);
            worksheets = workbook.Worksheets;
            worksheet = worksheets.Item[1];

            int rows = worksheet.UsedRange.Rows.Count;
            int cols = worksheet.UsedRange.Columns.Count;

            for (int i = 4; i <= rows; i++)
            {
                string? fio = (string)(worksheet.Cells[i, 1] as Excel.Range).Value2;

                int age = (int)(worksheet.Cells[i, 3] as Excel.Range).Value2;

                string? akanam =
                    (worksheet.Cells[i, 4].Value != null) ? (worksheet.Cells[i, 4] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 5].Value != null) ? (worksheet.Cells[i, 5] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 6].Value != null) ? (worksheet.Cells[i, 6] as Excel.Range).Value.ToString() : "";


                string? somanam =
                   (worksheet.Cells[i, 7].Value != null) ? (worksheet.Cells[i, 7] as Excel.Range).Value.ToString() :
                   (worksheet.Cells[i, 8].Value != null) ? (worksheet.Cells[i, 8] as Excel.Range).Value.ToString() :
                   (worksheet.Cells[i, 9].Value != null) ? (worksheet.Cells[i, 9] as Excel.Range).Value.ToString() : "";



                string? pregnancythreatStatus =
                    (worksheet.Cells[i, 10].Value != null) ? (worksheet.Cells[i, 10] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 11].Value != null) ? (worksheet.Cells[i, 11] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 12].Value != null) ? (worksheet.Cells[i, 12] as Excel.Range).Value.ToString() : "";



                string? edemaStatus =
                    (worksheet.Cells[i, 13].Value != null) ? (worksheet.Cells[i, 13] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 14].Value != null) ? (worksheet.Cells[i, 14] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 15].Value != null) ? (worksheet.Cells[i, 15] as Excel.Range).Value.ToString() : "";


                string? preeclampsiaStatus =
                    (worksheet.Cells[i, 16].Value != null) ? (worksheet.Cells[i, 16] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 17].Value != null) ? (worksheet.Cells[i, 17] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 18].Value != null) ? (worksheet.Cells[i, 18] as Excel.Range).Value.ToString() : "";


                string? anemiaStatus =
                    (worksheet.Cells[i, 19].Value != null) ? (worksheet.Cells[i, 19] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 20].Value != null) ? (worksheet.Cells[i, 20] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 21].Value != null) ? (worksheet.Cells[i, 21] as Excel.Range).Value.ToString() : "";


                string? infectionStatus =
                    (worksheet.Cells[i, 22].Value != null) ? (worksheet.Cells[i, 22] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 23].Value != null) ? (worksheet.Cells[i, 23] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 24].Value != null) ? (worksheet.Cells[i, 24] as Excel.Range).Value.ToString() : "";


                string? gravidogramStatus =
                    (worksheet.Cells[i, 25].Value != null) ? (worksheet.Cells[i, 25] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 26].Value != null) ? (worksheet.Cells[i, 26] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 27].Value != null) ? (worksheet.Cells[i, 27] as Excel.Range).Value.ToString() : "";


                string? bloodflowStatus =
                    (worksheet.Cells[i, 28].Value != null) ? (worksheet.Cells[i, 28] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 29].Value != null) ? (worksheet.Cells[i, 29] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 30].Value != null) ? (worksheet.Cells[i, 30] as Excel.Range).Value.ToString() : "";


                string? fetusStatus =
                    (worksheet.Cells[i, 31].Value != null) ? (worksheet.Cells[i, 31] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 32].Value != null) ? (worksheet.Cells[i, 32] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 33].Value != null) ? (worksheet.Cells[i, 33] as Excel.Range).Value.ToString() : "";


                string? feverStatus =
                    (worksheet.Cells[i, 34].Value != null) ? (worksheet.Cells[i, 34] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 35].Value != null) ? (worksheet.Cells[i, 35] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 36].Value != null) ? (worksheet.Cells[i, 36] as Excel.Range).Value.ToString() : "";


                string? coughStatus =
                    (worksheet.Cells[i, 37].Value != null) ? (worksheet.Cells[i, 37] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 38].Value != null) ? (worksheet.Cells[i, 38] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 39].Value != null) ? (worksheet.Cells[i, 39] as Excel.Range).Value.ToString() : "";


                string? dyspneaStatus =
                    (worksheet.Cells[i, 40].Value != null) ? (worksheet.Cells[i, 40] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 41].Value != null) ? (worksheet.Cells[i, 41] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 42].Value != null) ? (worksheet.Cells[i, 42] as Excel.Range).Value.ToString() : "";


                string? hemodynamicsStatus =
                    (worksheet.Cells[i, 43].Value != null) ? (worksheet.Cells[i, 43] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 44].Value != null) ? (worksheet.Cells[i, 44] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 45].Value != null) ? (worksheet.Cells[i, 45] as Excel.Range).Value.ToString() : "";


                string? saturationStatus =
                    (worksheet.Cells[i, 46].Value != null) ? (worksheet.Cells[i, 46] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 47].Value != null) ? (worksheet.Cells[i, 47] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 48].Value != null) ? (worksheet.Cells[i, 48] as Excel.Range).Value.ToString() : "";


                string? nauseaStatus =
                    (worksheet.Cells[i, 49].Value != null) ? (worksheet.Cells[i, 49] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 50].Value != null) ? (worksheet.Cells[i, 50] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 51].Value != null) ? (worksheet.Cells[i, 51] as Excel.Range).Value.ToString() : "";


                string? lungtissueStatus =
                    (worksheet.Cells[i, 52].Value != null) ? (worksheet.Cells[i, 52] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 53].Value != null) ? (worksheet.Cells[i, 53] as Excel.Range).Value.ToString() :
                    (worksheet.Cells[i, 54].Value != null) ? (worksheet.Cells[i, 54] as Excel.Range).Value.ToString() : "";

                string? group = (worksheet.Cells[i, 2] != null) ? ((worksheet.Cells[i, 2] as Excel.Range).Value != null ?
                    worksheet.Cells[i, 2].Value.ToString() : "") : "";
                if (group == "")
                {
                    float pregCoef = (Convert.ToByte(pregnancythreatStatus) + Convert.ToByte(edemaStatus) +
                Convert.ToByte(preeclampsiaStatus) + Convert.ToByte(anemiaStatus) + Convert.ToByte(infectionStatus) +
                Convert.ToByte(gravidogramStatus) + Convert.ToByte(bloodflowStatus) + Convert.ToByte(fetusStatus)) / 8f;
                    float sarsCov2 = (Convert.ToByte(feverStatus) + Convert.ToByte(coughStatus) + Convert.ToByte(dyspneaStatus) +
                        Convert.ToByte(hemodynamicsStatus) + Convert.ToByte(nauseaStatus) + Convert.ToByte(saturationStatus) +
                        Convert.ToByte(lungtissueStatus)) / 7f;
                    float fkrpo = Convert.ToByte(akanam) + Convert.ToByte(somanam) + pregCoef + sarsCov2;
                    
                    Debug.WriteLine($"fkrpo = {fkrpo}");

                    if (Math.Round(fkrpo, 1) <= 3.0)
                    {
                        group = "низкий риск перинататльных осложнений";
                    }
                    else if ((Math.Round(fkrpo, 1) >= 3.1) && (Math.Round(fkrpo, 1) <= 5.0))
                    {
                        group = "средний риск перинатальных осложнений";
                    }
                    else if ((Math.Round(fkrpo, 1) >= 5.1) && (Math.Round(fkrpo, 1) <= 8.0))
                    {
                        group = "высокий риск перинатальных осложнений";
                    }
                }

                Debug.WriteLine(group);
                var tmpPerson = new Person(fio, age, group, akanam, somanam, pregnancythreatStatus, edemaStatus, preeclampsiaStatus,
                        anemiaStatus, infectionStatus, gravidogramStatus, bloodflowStatus, fetusStatus, feverStatus, coughStatus, dyspneaStatus,
                        hemodynamicsStatus, saturationStatus, nauseaStatus, lungtissueStatus);
                Person.personlist.Add(tmpPerson);



            }

           
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(worksheets);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(workbooks);
            excel.Quit();
            Marshal.ReleaseComObject(excel);
            GC.Collect();

            return Person.personlist;


        }

    }
}
