using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1.Models
{
    public class Person
    {
        public Person( string? fio, int age, string? group, string? akanam, string? somanam, 
            string? pregnancythreatStatus, string? edemaStatus, string? preeclampsiaStatus, 
            string? anemiaStatus, string? infectionStatus, string? gravidogramStatus, 
            string? bloodflowStatus, string? fetusStatus, string? feverStatus, string? coughStatus, 
            string? dyspneaStatus, string? hemodynamicsStatus, string? saturationStatus, 
            string? nauseaStatus, string? lungtissueStatus)
        {
            Age = age;
            Fio = fio;
            Group = group;
            Akanam = akanam;
            Somanam = somanam;
            PregnancythreatStatus = pregnancythreatStatus;
            EdemaStatus = edemaStatus;
            PreeclampsiaStatus = preeclampsiaStatus;
            AnemiaStatus = anemiaStatus;
            InfectionStatus = infectionStatus;
            GravidogramStatus = gravidogramStatus;
            BloodflowStatus = bloodflowStatus;
            FetusStatus = fetusStatus;
            FeverStatus = feverStatus;
            CoughStatus = coughStatus;
            DyspneaStatus = dyspneaStatus;
            HemodynamicsStatus = hemodynamicsStatus;
            SaturationStatus = saturationStatus;
            NauseaStatus = nauseaStatus;
            LungtissueStatus = lungtissueStatus;
        }

        public int Age { get; set; }
        public string? Fio { get; set; }
        public string? Group { get; set; }
        public string? Akanam { get; set; }
        public string? Somanam { get; set; }
        public string? PregnancythreatStatus { get; set; }
        public string? EdemaStatus { get; set; }
        public string? PreeclampsiaStatus { get; set; }
        public string? AnemiaStatus { get; set; }
        public string? InfectionStatus { get; set; }
        public string? GravidogramStatus { get; set; }
        public string? BloodflowStatus { get; set; }
        public string? FetusStatus { get; set; }
        public string? FeverStatus { get; set; }
        public string? CoughStatus { get; set; }
        public string? DyspneaStatus { get; set; }
        public string? HemodynamicsStatus { get; set; }
        public string? SaturationStatus { get; set; }
        public string? NauseaStatus { get; set; }
        public string? LungtissueStatus { get; set; }

        public static List<Person> personlist= new List<Person>();
    }
}
