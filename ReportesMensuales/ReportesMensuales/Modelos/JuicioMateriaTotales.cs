using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReportesMensuales.Modelos
{
    public class JuicioMateriaTotales
    {
        private string civilOral;
        private string civilTradicional;
        private string familiarlOral;
        private string familiarTradicional;
        private string MercantilOral;
        private string MercantilTradicional;
        private string tipoJuzgado;
        private string civil;
        private string familiar;
        private string mercantil;
        private string celebradas;
        private string noCelebradas;
        private string control;
        private string tribunal;
        private string total;


        public string CivilOral { get => civilOral; set => civilOral = value; }
        public string CivilTradicional { get => civilTradicional; set => civilTradicional = value; }
        public string FamiliarlOral { get => familiarlOral; set => familiarlOral = value; }
        public string FamiliarTradicional { get => familiarTradicional; set => familiarTradicional = value; }
        public string TipoJuzgado { get => tipoJuzgado; set => tipoJuzgado = value; }
        public string MercantilOral1 { get => MercantilOral; set => MercantilOral = value; }
        public string MercantilTradicional1 { get => MercantilTradicional; set => MercantilTradicional = value; }
        public string Civil { get => civil; set => civil = value; }
        public string Familiar { get => familiar; set => familiar = value; }
        public string Mercantil { get => mercantil; set => mercantil = value; }
        public string Celebradas { get => celebradas; set => celebradas = value; }
        public string NoCelebradas { get => noCelebradas; set => noCelebradas = value; }
        public string Control { get => control; set => control = value; }
        public string Tribunal { get => tribunal; set => tribunal = value; }
        public string Total { get => total; set => total = value; }
    }
}