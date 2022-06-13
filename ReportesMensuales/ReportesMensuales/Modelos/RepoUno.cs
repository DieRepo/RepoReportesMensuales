using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReportesMensuales.Modelos
{
    public class RepoUno
    {
        private string nsjPenal;
        private string penal;
        private string laboral;
        private string adolescentes;
        private string penalAcu;
        private string celebradas;
        private string noCelebradas;
        private string ejecucion;
        private List<JuicioMateriaTotales> jui = new List<JuicioMateriaTotales>();
        
        public string NsjPenal { get => nsjPenal; set => nsjPenal = value; }
        public string Penal { get => penal; set => penal = value; }
        public string Laboral { get => laboral; set => laboral = value; }
        public List<JuicioMateriaTotales> Jui { get => jui; set => jui = value; }
        public string PenalAcu { get => penalAcu; set => penalAcu = value; }
        public string Celebradas { get => celebradas; set => celebradas = value; }
        public string NoCelebradas { get => noCelebradas; set => noCelebradas = value; }
        public string Ejecucion { get => ejecucion; set => ejecucion = value; }
    }
}