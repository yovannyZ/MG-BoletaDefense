using System;

namespace MG_BoletaDefense
{
    public class IBoleta
    {
        public string CodigoTrabajador { get; set; }
        public string DesTrabajador { get; set; }
        public string CodCentroCosto { get; set; }
        public string CodBase { get; set; }
        public string Direccion { get; set; }
        public string CargoTrabajador { get; set; }
        public DateTime FeIngreso { get; set; }
        public DateTime FeCese { get; set; }
        public DateTime FeSalidaVac { get; set; }
        public DateTime FeIngresoVac { get; set; }
        public decimal SueldoBasico { get; set; }
        public decimal DiasTrabajados { get; set; }
        public decimal HorasTrabajadas { get; set; }
        public string NumeroDoc { get; set; }
        public string Ips { get; set; }
        public string Afp { get; set; }
        public int Edad { get; set; }
        public string Situacion { get; set; }
        public string TipoPlanilla { get; set; }
        public string   DesEmpleado { get; set; }
        public DateTime FeAl { get; set; }
        public DateTime FeDel { get; set; }
        public string NuSecuencia { get; set; }
        public string FeAño { get; set; }
        public string FeMes { get; set; }
    }
}
