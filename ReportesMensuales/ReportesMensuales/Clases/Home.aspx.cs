using MySql.Data.MySqlClient;
using OfficeOpenXml;
using ReportesMensuales.Modelos;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace ReportesMensuales.Clases
{
    public partial class Home : System.Web.UI.Page
    {
        string query;
        string horaIni = "00:00:00";
        string horaFin = "23:59:59";


        protected void Page_Load(object sender, EventArgs e)
        {

        }

        public void GeneraReporteUno(object sender, EventArgs e)
        {
            //PrimerReporteCuadro1();
            PrimerReporteCuadro2();
            PrimerReporteCuadro3();
            PrimerReporteCuadro4();
            PrimerReporteCuadro5();
            PrimerReporteCuadro6();
            PrimerReporteCuadro7();
            PrimerReporteCuadro8();
            PrimerReporteCuadro9();
            PrimerReporteCuadro10();
            PrimerReporteCuadro11();
            PrimerReporteCuadro12();
            PrimerReporteCuadro13();
            //mascara.Visible = false;


        }

        public void PrimerReporteCuadro1()
        {
            string fichero = "C:\\Reportes/ReporteUno.xlsx";
            ExcelPackage excel = new ExcelPackage(new FileInfo(fichero));
            ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];

            RepoUno re = new RepoUno();

            try
            {

                MySqlConnection con = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con.Open();

                query = "Create Temporary Table IF NOT EXISTS estadistica.util_tblcatalogos " +
                        "(cveCatalogo char(7) COLLATE utf8_general_ci, " +
                        "descripcion varchar(200) CHARSET utf8 COLLATE utf8_general_ci, " +
                        "tipo char(2), " +
                        "INDEX cveCatalogo_util_idx(cveCatalogo), " +
                        "INDEX cveCatalogo_descripcion_idx(descripcion), " +
                        "INDEX cveCatalogo_tipo_idx(tipo)) " +
                        "(Select concat(codigo, rango) cveCatalogo, descripcion, tipo " +
                        "From estadistica.tblcatalogos c); " +
                           "select t.matjuzgado tipoJuzgado, " +
                            "ifnull(sum(if (t.matJuicio = 'CIVIL ORAL', t.total,null)),0) 'CO', " +
                            "ifnull(sum(if (t.matJuicio = 'CIVIL ESCRITO', t.total,null)),0) 'CT', " +
                            "ifnull(sum(if (t.matJuicio = 'FAMILIAR ORAL', t.total,null)),0) 'FO', " +
                            "ifnull(sum(if (t.matJuicio = 'FAMILIAR ESCRITO', t.total,null)),0) 'FT', " +
                            "ifnull(sum(if (t.matJuicio = 'MERCANTIL ORAL', t.total,null)),0) 'MO', " +
                            "ifnull(sum(if (t.matJuicio = 'MERCANTIL ESCRITO', t.total,null)),0) 'MT' " +
                            "from " +
                            "(select " +
                            "case " +
                            "When j.NomJuzgado like '%Civil%' or j.NomJuzgado like '%Menor%' or j.NomJuzgado like '%Mixto%'  or j.NomJuzgado like '%C.M.%'or j.NomJuzgado like '%C. M.%' then 'JUZGADOS CIVILES' " +
                            "When j.NomJuzgado like '%Mercantil%' then 'JUZGADOS MERCANTILES' " +
                            "When j.NomJuzgado like '%Familiar%' or j.nomJuzgado like '%Adopcion%' then 'JUZGADOS FAMILIARES' " +
                            "When j.NomJuzgado like '%Usucapion%' then 'JUZGADOS USUCAPION' " +
                            "Else j.NomJuzgado  end  matJuzgado, " +
                            "case " +
                            "when r.CveJuicioDelito = 'A102242' then 'CIVIL ORAL' " +
                            "when r.CveJuicioDelito = 'A103024' or((c.tipo = 'C' or c.tipo like 'C%') and r.CveJuicioDelito != 'A102242') then 'CIVIL ESCRITO' " +
                            "when((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121', 'A103067') and(c.descripcion not like '%Sucesorio%' and c.descripcion not like '%Tramitacion%' " +
                            "and c.descripcion not like '%testament%' and c.descripcion not like '%here%')) then 'FAMILIAR ORAL' " +
                            "when(((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121')  and(c.descripcion like '%Sucesorio%' or c.descripcion like '%Tramitacion%' " +
                            "or c.descripcion like '%testament%' or c.descripcion like '%here%') or r.CveJuicioDelito = 'A103067')) then 'FAMILIAR ESCRITO' " +
                            "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion like '%oral%') then 'MERCANTIL ORAL' " +
                            "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion not like '%oral%') or r.cveJuicioDelito = 'A103121' then 'MERCANTIL ESCRITO' " +
                            "else if (c.descripcion = 'NO EXISTE EXPEDIENTE', " +
                            "(case when j.NomJuzgado like '%Civil%' or j.NomJuzgado like '%Menor%' or j.NomJuzgado like '%Mixto%'  or j.NomJuzgado like '%C.M.%' or j.NomJuzgado like '%Usucapion%'  then 'CIVIL ESCRITO' " +
                            "when j.NomJuzgado like '%Familiar%' or j.nomJuzgado like '%Adopcion%' then 'FAMILIAR ORAL' " +
                            "when j.NomJuzgado like '%Mercantil%' then 'MERCANTIL ESCRITO' " +
                            "else j.NomJuzgado end), c.descripcion) " +
                            "end matJuicio, " +
                            "count(distinct i.cveini) total " +
                            "from estadistica.tblinijuzgados i " +
                            "inner join estadistica.tblrepjuidel r on i.cveini = r.cveini " +
                            "inner join estadistica.util_tblcatalogos c on c.cveCatalogo = r.CveJuicioDelito " +
                            "inner join estadistica.tbljuzgados j on j.cveAdscripcion = i.cveJuzgado " +
                            "where i.fechaRad between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                            "and  i.Estado = 1 " +
                            "and((i.Observaciones not like '%migrado%' or i.Observaciones is null) and i.idIniAnterior = 0 and(i.cveJuzProcedencia = 0  or i.cveJuzProcedencia is null)) " +
                            "group by matJuzgado, matJuicio) as t " +
                            "group by t.matjuzgado " +
                            "order by field(t.matjuzgado,'JUZGADOS CIVILES','JUZGADOS USUCAPION','JUZGADOS MERCANTILES','JUZGADOS FAMILIARES')";

                MySqlCommand cmd = new MySqlCommand(query, con);
                cmd.CommandTimeout = 1800;
                MySqlDataReader r = cmd.ExecuteReader();


                while (r.Read())
                {
                    JuicioMateriaTotales j = new JuicioMateriaTotales();
                    j.TipoJuzgado = r.GetString("tipoJuzgado");
                    j.CivilOral = r.GetString("CO");
                    j.CivilTradicional = r.GetString("CT");
                    j.FamiliarlOral = r.GetString("FO");
                    j.FamiliarTradicional = r.GetString("FT");
                    j.MercantilOral1 = r.GetString("MO");
                    j.MercantilTradicional1 = r.GetString("MT");
                    re.Jui.Add(j);

                }
                con.Close();

                /*Civil*/
                List<JuicioMateriaTotales> civiles = new List<JuicioMateriaTotales>();
                civiles = re.Jui;
                int col = 7;

                foreach (JuicioMateriaTotales c in civiles)
                {
                    worksheet.Cells[col, 5].Value = c.CivilOral;
                    col++;
                }

                List<JuicioMateriaTotales> civilesT = new List<JuicioMateriaTotales>();
                civilesT = re.Jui;
                int col1 = 7;

                foreach (JuicioMateriaTotales c in civilesT)
                {
                    worksheet.Cells[col1, 6].Value = c.CivilTradicional;
                    col1++;
                }

                /*Familiar*/
                List<JuicioMateriaTotales> familiar = new List<JuicioMateriaTotales>();
                familiar = re.Jui;
                int col2 = 7;

                foreach (JuicioMateriaTotales c in familiar)
                {
                    worksheet.Cells[col2, 7].Value = c.FamiliarlOral;
                    col2++;
                }

                List<JuicioMateriaTotales> familiarT = new List<JuicioMateriaTotales>();
                familiarT = re.Jui;
                int col3 = 7;

                foreach (JuicioMateriaTotales c in familiarT)
                {
                    worksheet.Cells[col3, 8].Value = c.FamiliarTradicional;
                    col3++;
                }

                /*Mercantil*/
                List<JuicioMateriaTotales> mercantil = new List<JuicioMateriaTotales>();
                mercantil = re.Jui;
                int col4 = 7;

                foreach (JuicioMateriaTotales c in mercantil)
                {
                    worksheet.Cells[col4, 12].Value = c.MercantilOral1;
                    col4++;
                }

                List<JuicioMateriaTotales> mercantilT = new List<JuicioMateriaTotales>();
                mercantilT = re.Jui;
                int col5 = 7;

                foreach (JuicioMateriaTotales c in mercantilT)
                {
                    worksheet.Cells[col5, 13].Value = c.MercantilTradicional1;
                    col5++;
                }

                /*Laboral*/
                MySqlConnection con1 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql"]);
                con1.Open();

                string q = "select count(distinct cj.idCarpetaJudicial) total " +
                            "from htsj_laboral.tblcarpetasjudiciales cj " +
                            "left " +
                            "join htsj_laboral.tbljuzgados as jz on cj.cveJuzgado = jz.cveJuzgado " +
                            "where cj.fechaRadicacion between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                            "and cvetipocarpeta = 1 " +
                            "AND jz.cveJuzgado not in (9, 10, 11) " +
                            "AND cj.activo = 'S' " +
                            "AND jz.activo = 'S' ";

                MySqlCommand cmd1 = new MySqlCommand(q, con1);
                MySqlDataReader r1 = cmd1.ExecuteReader();


                while (r1.Read())
                {
                    re.Laboral = r1.GetString("total");

                }
                con1.Close();

                /*Penal Escrito*/
                MySqlConnection con2 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con2.Open();

                string qu = "select " +
                            "count(distinct i.cveini) total " +
                            "from estadistica.tblinijuzpen i " +
                            "where i.fechaRad between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                            "and(i.CveJuzPro = 0  or i.CveJuzPro is null) and(i.CveExpJuzPro = 0  or i.CveExpJuzPro is null) and(i.AnioExpJuzPro = 0  or i.AnioExpJuzPro is null) " +
                            "and i.Estado = 1; ";

                MySqlCommand cmd2 = new MySqlCommand(qu, con2);
                MySqlDataReader r2 = cmd2.ExecuteReader();


                while (r2.Read())
                {
                    re.Penal = r2.GetString("total");

                }
                con2.Close();

                /*Penal acusatorio*/
                MySqlConnection con3 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["sigejupe"]);
                con3.Open();


                string q3 = "SELECT " +
                            "count(distinct idCarpetajudicial) total " +
                            "FROM htsj_sigejupe.tblcarpetasjudiciales " +
                            "where " +
                            "fecharadicacion between '" + fecIni.Text + " " + " " + horaIni + " ' and '" + fecFin.Text + " " + " " + horaFin + " '  " +
                            "and cveTipoCarpeta in (2, 3, 4) and cvejuzgado != 11353  and activo = 'S' ";

                MySqlCommand cmd3 = new MySqlCommand(q3, con3);
                MySqlDataReader r3 = cmd3.ExecuteReader();


                while (r3.Read())
                {
                    re.PenalAcu = r3.GetString("total");

                }
                con3.Close();

                /*Ejecucion*/
                MySqlConnection con4 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con4.Open();


                string q4 = "select " + 
                            "count(distinct idIniciado) total " +
                            "from htsj_ejecucion_tradicional.tbliniciados " +
                            "where fecha_iniciado between '" + fecIni.Text + "' and '" + fecFin.Text + "' ";

                MySqlCommand cmd4 = new MySqlCommand(q4, con4);
                MySqlDataReader r4 = cmd4.ExecuteReader();


                while (r4.Read())
                {
                    re.Ejecucion = r4.GetString("total");

                }
                con4.Close();

                worksheet.Cells[12, 4].Value = re.Penal;
                worksheet.Cells[14, 9].Value = re.Laboral;
                worksheet.Cells[11, 3].Value = re.PenalAcu;
                worksheet.Cells[13, 3].Value = re.Ejecucion;

                excel.Save();
                excel.Dispose();
                Debug.WriteLine("Se termino el proceso 1 ...");
            }
            catch (Exception i)
            {
                Debug.WriteLine("No se establecio la conexion :" + i);

            }
        }

        public void PrimerReporteCuadro2()
        {
            string fichero = "C:\\Reportes/ReporteUno.xlsx";
            ExcelPackage excel = new ExcelPackage(new FileInfo(fichero));
            ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];

            RepoUno re = new RepoUno();

            try
            {

                MySqlConnection con = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con.Open();

                query = "select t.matjuzgado tipoJuzgado, " +
                            "ifnull(sum(if (t.matJuicio = 'CIVIL ORAL', t.total,null)),0) 'CO', " +
                            "ifnull(sum(if (t.matJuicio = 'CIVIL ESCRITO', t.total,null)),0) 'CT', " +
                            "ifnull(sum(if (t.matJuicio = 'FAMILIAR ORAL', t.total,null)),0) 'FO', " +
                            "ifnull(sum(if (t.matJuicio = 'FAMILIAR ESCRITO', t.total,null)),0) 'FT', " +
                            "ifnull(sum(if (t.matJuicio = 'MERCANTIL ORAL', t.total,null)),0) 'MO', " +
                            "ifnull(sum(if (t.matJuicio = 'MERCANTIL ESCRITO', t.total,null)),0) 'MT' " +
                            "from " +
                            "(select " +
                            "case " +
                            "When j.NomJuzgado like '%Civil%' or j.NomJuzgado like '%Menor%' or j.NomJuzgado like '%Mixto%'  or j.NomJuzgado like '%C.M.%'or j.NomJuzgado like '%C. M.%' then 'JUZGADOS CIVILES' " +
                            "When j.NomJuzgado like '%Mercantil%' then 'JUZGADOS MERCANTILES' " +
                            "When j.NomJuzgado like '%Familiar%' or j.nomJuzgado like '%Adopcion%' then 'JUZGADOS FAMILIARES' " +
                            "When j.NomJuzgado like '%Usucapion%' then 'JUZGADOS USUCAPION' " +
                            "Else j.NomJuzgado  end  matJuzgado, " +
                            "case " +
                            "when r.CveJuicioDelito = 'A102242' then 'CIVIL ORAL' " +
                            "when r.CveJuicioDelito = 'A103024' or((c.tipo = 'C' or c.tipo like 'C%') and r.CveJuicioDelito != 'A102242') then 'CIVIL ESCRITO' " +
                            "when((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121', 'A103067') and(c.descripcion not like '%Sucesorio%' and c.descripcion not like '%Tramitacion%' " +
                            "and c.descripcion not like '%testament%' and c.descripcion not like '%here%')) then 'FAMILIAR ORAL' " +
                            "when(((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121')  and(c.descripcion like '%Sucesorio%' or c.descripcion like '%Tramitacion%' " +
                            "or c.descripcion like '%testament%' or c.descripcion like '%here%') or r.CveJuicioDelito = 'A103067')) then 'FAMILIAR ESCRITO' " +
                            "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion like '%oral%') then 'MERCANTIL ORAL' " +
                            "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion not like '%oral%') or r.cveJuicioDelito = 'A103121' then 'MERCANTIL ESCRITO' " +
                            "else if (c.descripcion = 'NO EXISTE EXPEDIENTE', " +
                            "(case when j.NomJuzgado like '%Civil%' or j.NomJuzgado like '%Menor%' or j.NomJuzgado like '%Mixto%'  or j.NomJuzgado like '%C.M.%' or j.NomJuzgado like '%Usucapion%'  then 'CIVIL ESCRITO' " +
                            "when j.NomJuzgado like '%Familiar%' or j.nomJuzgado like '%Adopcion%' then 'FAMILIAR ORAL' " +
                            "when j.NomJuzgado like '%Mercantil%' then 'MERCANTIL ESCRITO' " +
                            "else j.NomJuzgado end), c.descripcion) " +
                            "end matJuicio, " +
                            "count(distinct t.cveter) total " +
                            "from estadistica.tblterjuzgados t " +
                            "inner join estadistica.tblinijuzgados i on i.cveini = t.cveini " +
                            "inner join estadistica.tblrepjuidel r on i.cveini = r.cveini " +
                            "inner join estadistica.util_tblcatalogos c on c.cveCatalogo = r.CveJuicioDelito " +
                            "inner join estadistica.tbljuzgados j on j.cveAdscripcion = i.cveJuzgado " +
                            "where t.fechater between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                            "and  i.Estado = 1 " +
                            "group by matJuzgado, matJuicio) as t " +
                            "group by t.matjuzgado " +
                            "order by field(t.matjuzgado,'JUZGADOS CIVILES','JUZGADOS USUCAPION','JUZGADOS MERCANTILES','JUZGADOS FAMILIARES') ";

                MySqlCommand cmd = new MySqlCommand(query, con);
                cmd.CommandTimeout = 1800;
                MySqlDataReader r = cmd.ExecuteReader();


                while (r.Read())
                {
                    JuicioMateriaTotales j = new JuicioMateriaTotales();
                    j.TipoJuzgado = r.GetString("tipoJuzgado");
                    j.CivilOral = r.GetString("CO");
                    j.CivilTradicional = r.GetString("CT");
                    j.FamiliarlOral = r.GetString("FO");
                    j.FamiliarTradicional = r.GetString("FT");
                    j.MercantilOral1 = r.GetString("MO");
                    j.MercantilTradicional1 = r.GetString("MT");
                    re.Jui.Add(j);

                }
                con.Close();

                /*Civil*/
                List<JuicioMateriaTotales> civiles = new List<JuicioMateriaTotales>();
                civiles = re.Jui;
                int col = 22;

                foreach (JuicioMateriaTotales c in civiles)
                {
                    worksheet.Cells[col, 5].Value = c.CivilOral;
                    col++;
                }

                List<JuicioMateriaTotales> civilesT = new List<JuicioMateriaTotales>();
                civilesT = re.Jui;
                int col1 = 22;

                foreach (JuicioMateriaTotales c in civilesT)
                {
                    worksheet.Cells[col1, 6].Value = c.CivilTradicional;
                    col1++;
                }

                /*Familiar*/
                List<JuicioMateriaTotales> familiar = new List<JuicioMateriaTotales>();
                familiar = re.Jui;
                int col2 = 22;

                foreach (JuicioMateriaTotales c in familiar)
                {
                    worksheet.Cells[col2, 7].Value = c.FamiliarlOral;
                    col2++;
                }

                List<JuicioMateriaTotales> familiarT = new List<JuicioMateriaTotales>();
                familiarT = re.Jui;
                int col3 = 22;

                foreach (JuicioMateriaTotales c in familiarT)
                {
                    worksheet.Cells[col3, 8].Value = c.FamiliarTradicional;
                    col3++;
                }

                /*Mercantil*/
                List<JuicioMateriaTotales> mercantil = new List<JuicioMateriaTotales>();
                mercantil = re.Jui;
                int col4 = 22;

                foreach (JuicioMateriaTotales c in mercantil)
                {
                    worksheet.Cells[col4, 12].Value = c.MercantilOral1;
                    col4++;
                }

                List<JuicioMateriaTotales> mercantilT = new List<JuicioMateriaTotales>();
                mercantilT = re.Jui;
                int col5 = 22;

                foreach (JuicioMateriaTotales c in mercantilT)
                {
                    worksheet.Cells[col5, 13].Value = c.MercantilTradicional1;
                    col5++;
                }

                /*Laboral*/
                MySqlConnection con1 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql"]);
                con1.Open();

                string q = "select sum(t.total) total from (SELECT " +
                            "count(distinct s.idSentencia) total " +
                            "FROM htsj_laboral.tblsentencias s " +
                            "inner join htsj_laboral.tblactuaciones a on a.idActuacion = s.idActuacion " +
                            "inner join htsj_laboral.tblcarpetasjudiciales c on c.idCarpetaJudicial = a.idCarpetaJudicial " +
                            "left join htsj_laboral.tbljuicios ju on ju.cveJuicio = c.cveJuicio " +
                            "left join htsj_laboral.tbljuzgados j on j.cveJuzgado = a.cveJuzgado " +
                             "where s.cveTipoSentencia = 2 and s.fechaSentencia between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                            "union all " +
                            "SELECT " +
                            "count(c.idCarpetaJudicial) total " +
                            "from  htsj_laboral.tblcarpetasjudiciales c " +
                            "left join htsj_laboral.tbljuzgados as j on c.cveJuzgado = j.cveJuzgado " +
                            "left join htsj_laboral.tbltiposterminaciones t on t.cveTipoTerminacion = c.cveTipoTerminacion " +
                            "where c.fechaSolucionado between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                            "and c.cveTipoTerminacion = 1 " +
                            "and c.cvetipocarpeta = 1 " +
                            "AND j.cveJuzgado not in (9, 10, 11) " +
                            "AND c.activo = 'S' " +
                            "AND j.activo = 'S' " +
                            "union all " +
                            "SELECT " +
                            "count(c.idCarpetaJudicial) total " +
                            "from  htsj_laboral.tblcarpetasjudiciales c " +
                            "left join htsj_laboral.tbljuzgados as j on c.cveJuzgado = j.cveJuzgado " +
                            "left join htsj_laboral.tbltiposterminaciones t on t.cveTipoTerminacion = c.cveTipoTerminacion " +
                            "where c.fechaSolucionado between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                            "and c.cvetipocarpeta = 1 " +
                            "and t.desctipoterminacion not like 'SENTENCIA' " +
                            "and t.desctipoterminacion not like 'CONVENIO' " +
                            "AND j.cveJuzgado not in (9, 10, 11) " +
                            "AND c.activo = 'S' AND j.activo = 'S' " +
                            ") t ";

                MySqlCommand cmd1 = new MySqlCommand(q, con1);
                MySqlDataReader r1 = cmd1.ExecuteReader();


                while (r1.Read())
                {
                    re.Laboral = r1.GetString("total");

                }
                con1.Close();

                /*Penal Escrito*/
                MySqlConnection con2 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con2.Open();

                string qu = "select " +
                            "count(distinct t.cveter) total " +
                            "from estadistica.tblinijuzpen i " +
                            "inner " +
                            "join estadistica.tblterjuzpen t on t.cveini = i.cveini " +
                            "where t.FechaTer between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                            "and i.Estado = 1 ";

                MySqlCommand cmd2 = new MySqlCommand(qu, con2);
                MySqlDataReader r2 = cmd2.ExecuteReader();


                while (r2.Read())
                {
                    re.Penal = r2.GetString("total");

                }
                con2.Close();

                /*Penal acusatorio*/
                MySqlConnection con3 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["sigejupe-local"]);
                con3.Open();


                string q3 = "SELECT " +
                            "count(distinct ct.idCarpetajudicial) total " +
                            "FROM htsj_sigejupe.tblcarpetasjudiciales c " +
                            "inner join htsj_sigejupe.tblcarpetasjudicialesterminadas ct on ct.idCarpetaJudicial = c.idCarpetaJudicial " +
                            "where " +
                            "ct.fechatermino between '" + fecIni.Text + " " + " " + horaIni + " ' and '" + fecFin.Text + " " + " " + horaFin + " ' " +
                            "and cveTipoCarpeta in (2, 3, 4) and cvejuzgado != 11353  and activo = 'S' ";

                MySqlCommand cmd3 = new MySqlCommand(q3, con3);
                MySqlDataReader r3 = cmd3.ExecuteReader();


                while (r3.Read())
                {
                    re.PenalAcu = r3.GetString("total");

                }
                con3.Close();

                /*Ejecucion*/
                MySqlConnection con4 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con4.Open();


                string q4 = "select count(distinct idConcluido) total " + 
                                "from htsj_ejecucion_tradicional.tblconcluidos " +
                                "where fecha_conclusion between '" + fecIni.Text + "' and '" + fecFin.Text + "' ";

                MySqlCommand cmd4 = new MySqlCommand(q4, con4);
                MySqlDataReader r4 = cmd4.ExecuteReader();


                while (r4.Read())
                {
                    re.Ejecucion = r4.GetString("total");

                }
                con4.Close();

                worksheet.Cells[27, 4].Value = re.Penal;
                worksheet.Cells[29, 9].Value = re.Laboral;
                worksheet.Cells[26, 3].Value = re.PenalAcu;
                worksheet.Cells[28, 3].Value = re.Ejecucion;


                excel.Save();
                excel.Dispose();
                Debug.WriteLine("Se termino el proceso 2 ...");
            }
            catch (Exception i)
            {
                Debug.WriteLine("No se establecio la conexion :" + i);

            }
        }

        public void PrimerReporteCuadro3()
        {
            string fichero = "C:\\Reportes/ReporteUno.xlsx";
            ExcelPackage excel = new ExcelPackage(new FileInfo(fichero));
            ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];

            RepoUno re = new RepoUno();

            try
            {

                MySqlConnection con = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con.Open();

                query = "select t.matjuzgado tipoJuzgado, " +
                        "ifnull(sum(if (t.matJuicio = 'CIVIL ORAL', t.total,null)),0) 'CO', " +
                        "ifnull(sum(if (t.matJuicio = 'CIVIL ESCRITO', t.total,null)),0) 'CT', " +
                        "ifnull(sum(if (t.matJuicio = 'FAMILIAR ORAL', t.total,null)),0) 'FO', " +
                        "ifnull(sum(if (t.matJuicio = 'FAMILIAR ESCRITO', t.total,null)),0) 'FT', " +
                        "ifnull(sum(if (t.matJuicio = 'MERCANTIL ORAL', t.total,null)),0) 'MO', " +
                        "ifnull(sum(if (t.matJuicio = 'MERCANTIL ESCRITO', t.total,null)),0) 'MT' " +
                        "from " +
                        "(select  " +
                        "case " +
                        "When j.NomJuzgado like '%Civil%' or j.NomJuzgado like '%Menor%' or j.NomJuzgado like '%Mixto%'  or j.NomJuzgado like '%C.M.%'or j.NomJuzgado like '%C. M.%' then 'JUZGADOS CIVILES' " +
                        "When j.NomJuzgado like '%Mercantil%' then 'JUZGADOS MERCANTILES' " +
                        "When j.NomJuzgado like '%Familiar%' or j.nomJuzgado like '%Adopcion%' then 'JUZGADOS FAMILIARES' " +
                        "When j.NomJuzgado like '%Usucapion%' then 'JUZGADOS USUCAPION' " +
                        "Else j.NomJuzgado  end  matJuzgado, " +
                        "case " +
                        "when r.CveJuicioDelito = 'A102242' then 'CIVIL ORAL' " +
                        "when r.CveJuicioDelito = 'A103024' or((c.tipo = 'C' or c.tipo like 'C%') and r.CveJuicioDelito != 'A102242') then 'CIVIL ESCRITO' " +
                        "when((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121', 'A103067') and(c.descripcion not like '%Sucesorio%' and c.descripcion not like '%Tramitacion%' " +
                        "and c.descripcion not like '%testament%' and c.descripcion not like '%here%')) then 'FAMILIAR ORAL' " +
                        "when(((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121')  and(c.descripcion like '%Sucesorio%' or c.descripcion like '%Tramitacion%' " +
                        "or c.descripcion like '%testament%' or c.descripcion like '%here%') or r.CveJuicioDelito = 'A103067')) then 'FAMILIAR ESCRITO' " +
                        "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion like '%oral%') then 'MERCANTIL ORAL' " +
                        "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion not like '%oral%') or r.cveJuicioDelito = 'A103121' then 'MERCANTIL ESCRITO' " +
                        "else if (c.descripcion = 'NO EXISTE EXPEDIENTE', " +
                        "(case when j.NomJuzgado like '%Civil%' or j.NomJuzgado like '%Menor%' or j.NomJuzgado like '%Mixto%'  or j.NomJuzgado like '%C.M.%' or j.NomJuzgado like '%Usucapion%'  then 'CIVIL ESCRITO' " +
                        "when j.NomJuzgado like '%Familiar%' or j.nomJuzgado like '%Adopcion%' then 'FAMILIAR ORAL' " +
                        "when j.NomJuzgado like '%Mercantil%' then 'MERCANTIL ESCRITO' " +
                        "else j.NomJuzgado end), c.descripcion) " +
                        "end matJuicio, " +
                        "count(distinct t.cveter) total " +
                        "from estadistica.tblterjuzgados t " +
                        "inner join estadistica.tblinijuzgados i on i.cveini = t.cveini " +
                        "inner join estadistica.tblrepjuidel r on i.cveini = r.cveini " +
                        "inner join estadistica.util_tblcatalogos c on c.cveCatalogo = r.CveJuicioDelito " +
                        "inner join estadistica.tbljuzgados j on j.cveAdscripcion = i.cveJuzgado " +
                        "where t.fechater between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                        "and i.Estado = 1 and t.CveResolucion in ('A031001') " +
                        "group by matJuzgado, matJuicio) as t " +
                        "group by t.matjuzgado " +
                        "order by field(t.matjuzgado,'JUZGADOS CIVILES','JUZGADOS USUCAPION','JUZGADOS MERCANTILES','JUZGADOS FAMILIARES') ";

                MySqlCommand cmd = new MySqlCommand(query, con);
                cmd.CommandTimeout = 1800;
                MySqlDataReader r = cmd.ExecuteReader();


                while (r.Read())
                {
                    JuicioMateriaTotales j = new JuicioMateriaTotales();
                    j.TipoJuzgado = r.GetString("tipoJuzgado");
                    j.CivilOral = r.GetString("CO");
                    j.CivilTradicional = r.GetString("CT");
                    j.FamiliarlOral = r.GetString("FO");
                    j.FamiliarTradicional = r.GetString("FT");
                    j.MercantilOral1 = r.GetString("MO");
                    j.MercantilTradicional1 = r.GetString("MT");
                    re.Jui.Add(j);

                }
                con.Close();

                /*Civil*/
                List<JuicioMateriaTotales> civiles = new List<JuicioMateriaTotales>();
                civiles = re.Jui;
                int col = 37;

                foreach (JuicioMateriaTotales c in civiles)
                {
                    worksheet.Cells[col, 5].Value = c.CivilOral;
                    col++;
                }

                List<JuicioMateriaTotales> civilesT = new List<JuicioMateriaTotales>();
                civilesT = re.Jui;
                int col1 = 37;

                foreach (JuicioMateriaTotales c in civilesT)
                {
                    worksheet.Cells[col1, 6].Value = c.CivilTradicional;
                    col1++;
                }

                /*Familiar*/
                List<JuicioMateriaTotales> familiar = new List<JuicioMateriaTotales>();
                familiar = re.Jui;
                int col2 = 37;

                foreach (JuicioMateriaTotales c in familiar)
                {
                    worksheet.Cells[col2, 7].Value = c.FamiliarlOral;
                    col2++;
                }

                List<JuicioMateriaTotales> familiarT = new List<JuicioMateriaTotales>();
                familiarT = re.Jui;
                int col3 = 37;

                foreach (JuicioMateriaTotales c in familiarT)
                {
                    worksheet.Cells[col3, 8].Value = c.FamiliarTradicional;
                    col3++;
                }

                /*Mercantil*/
                List<JuicioMateriaTotales> mercantil = new List<JuicioMateriaTotales>();
                mercantil = re.Jui;
                int col4 = 37;

                foreach (JuicioMateriaTotales c in mercantil)
                {
                    worksheet.Cells[col4, 12].Value = c.MercantilOral1;
                    col4++;
                }

                List<JuicioMateriaTotales> mercantilT = new List<JuicioMateriaTotales>();
                mercantilT = re.Jui;
                int col5 = 37;

                foreach (JuicioMateriaTotales c in mercantilT)
                {
                    worksheet.Cells[col5, 13].Value = c.MercantilTradicional1;
                    col5++;
                }

                /*Laboral*/
                MySqlConnection con1 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql"]);
                con1.Open();

                        string q = "SELECT " +
                        "count(distinct s.idSentencia) total " +
                        "FROM htsj_laboral.tblsentencias s " +
                        "inner join htsj_laboral.tblactuaciones a on a.idActuacion = s.idActuacion " +
                        "inner join htsj_laboral.tblcarpetasjudiciales c on c.idCarpetaJudicial = a.idCarpetaJudicial " +
                        "left join htsj_laboral.tbljuicios ju on ju.cveJuicio = c.cveJuicio " +
                        "left join htsj_laboral.tbljuzgados j on j.cveJuzgado = a.cveJuzgado " +
                        "where s.cveTipoSentencia = 2 and s.fechaSentencia between '" + fecIni.Text + "' and '" + fecFin.Text + "' ";

                MySqlCommand cmd1 = new MySqlCommand(q, con1);
                MySqlDataReader r1 = cmd1.ExecuteReader();


                while (r1.Read())
                {
                    re.Laboral = r1.GetString("total");

                }
                con1.Close();

                /*Penal Escrito*/
                MySqlConnection con2 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con2.Open();

                string qu = "select " +
                    "count(distinct ter.cveTer) total " +
                    "from Estadistica.tblterjuzpen ter " +
                    "where " +
                    "ter.cveResolucion in ('A041001','A041002','A041003') " +
                    "and ter.fechater between '" + fecIni.Text + "' and '" + fecFin.Text + "' "; ;

                MySqlCommand cmd2 = new MySqlCommand(qu, con2);
                MySqlDataReader r2 = cmd2.ExecuteReader();


                while (r2.Read())
                {
                    re.Penal = r2.GetString("total");

                }
                con2.Close();

                /*Penal acusatorio*/
                MySqlConnection con3 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["sigejupe"]);
                con3.Open();


                string q3 = "select count(distinct idActuacion) total " +
                            "from htsj_sigejupe.tblactuaciones " +
                            "where " +
                            "fechasentencia between '" + fecIni.Text + " " + " " + horaIni + " ' and '" + fecFin.Text + " " + " " + horaFin + " ' " +
                            "and cvetipoactuacion = 3 and cvetipocarpeta in (2, 3, 4) and cvejuzgado != 11353 and activo = 'S' ";

                MySqlCommand cmd3 = new MySqlCommand(q3, con3);
                MySqlDataReader r3 = cmd3.ExecuteReader();
                cmd3.CommandTimeout = 1800;
                


                while (r3.Read())
                {
                    re.PenalAcu = r3.GetString("total");

                }
                con3.Close();



                worksheet.Cells[42, 4].Value = re.Penal;
                worksheet.Cells[43, 9].Value = re.Laboral;
                worksheet.Cells[41, 3].Value = re.PenalAcu;

                excel.Save();
                excel.Dispose();
                Debug.WriteLine("Se termino el proceso 3 ...");
            }
            catch (Exception i)
            {
                Debug.WriteLine("No se establecio la conexion :" + i);

            }
        }

        public void PrimerReporteCuadro4()
        {
            string fichero = "C:\\Reportes/ReporteUno.xlsx";
            ExcelPackage excel = new ExcelPackage(new FileInfo(fichero));
            ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];

            RepoUno re = new RepoUno();

            try
            {

                MySqlConnection con = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con.Open();

                query = "select t.matjuzgado tipoJuzgado, " +
                        "ifnull(sum(if (t.matJuicio = 'CIVIL ORAL', t.total,null)),0) 'CO', " +
                        "ifnull(sum(if (t.matJuicio = 'CIVIL ESCRITO', t.total,null)),0) 'CT', " +
                        "ifnull(sum(if (t.matJuicio = 'FAMILIAR ORAL', t.total,null)),0) 'FO', " +
                        "ifnull(sum(if (t.matJuicio = 'FAMILIAR ESCRITO', t.total,null)),0) 'FT', " +
                        "ifnull(sum(if (t.matJuicio = 'MERCANTIL ORAL', t.total,null)),0) 'MO', " +
                        "ifnull(sum(if (t.matJuicio = 'MERCANTIL ESCRITO', t.total,null)),0) 'MT' " +
                        "from " +
                        "(select  " +
                        "case " +
                        "When j.NomJuzgado like '%Civil%' or j.NomJuzgado like '%Menor%' or j.NomJuzgado like '%Mixto%'  or j.NomJuzgado like '%C.M.%'or j.NomJuzgado like '%C. M.%' then 'JUZGADOS CIVILES' " +
                        "When j.NomJuzgado like '%Mercantil%' then 'JUZGADOS MERCANTILES' " +
                        "When j.NomJuzgado like '%Familiar%' or j.nomJuzgado like '%Adopcion%' then 'JUZGADOS FAMILIARES' " +
                        "When j.NomJuzgado like '%Usucapion%' then 'JUZGADOS USUCAPION' " +
                        "Else j.NomJuzgado  end  matJuzgado, " +
                        "case " +
                        "when r.CveJuicioDelito = 'A102242' then 'CIVIL ORAL' " +
                        "when r.CveJuicioDelito = 'A103024' or((c.tipo = 'C' or c.tipo like 'C%') and r.CveJuicioDelito != 'A102242') then 'CIVIL ESCRITO' " +
                        "when((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121', 'A103067') and(c.descripcion not like '%Sucesorio%' and c.descripcion not like '%Tramitacion%' " +
                        "and c.descripcion not like '%testament%' and c.descripcion not like '%here%')) then 'FAMILIAR ORAL' " +
                        "when(((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121')  and(c.descripcion like '%Sucesorio%' or c.descripcion like '%Tramitacion%' " +
                        "or c.descripcion like '%testament%' or c.descripcion like '%here%') or r.CveJuicioDelito = 'A103067')) then 'FAMILIAR ESCRITO' " +
                        "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion like '%oral%') then 'MERCANTIL ORAL' " +
                        "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion not like '%oral%') or r.cveJuicioDelito = 'A103121' then 'MERCANTIL ESCRITO' " +
                        "else if (c.descripcion = 'NO EXISTE EXPEDIENTE', " +
                        "(case when j.NomJuzgado like '%Civil%' or j.NomJuzgado like '%Menor%' or j.NomJuzgado like '%Mixto%'  or j.NomJuzgado like '%C.M.%' or j.NomJuzgado like '%Usucapion%'  then 'CIVIL ESCRITO' " +
                        "when j.NomJuzgado like '%Familiar%' or j.nomJuzgado like '%Adopcion%' then 'FAMILIAR ORAL' " +
                        "when j.NomJuzgado like '%Mercantil%' then 'MERCANTIL ESCRITO' " +
                        "else j.NomJuzgado end), c.descripcion) " +
                        "end matJuicio, " +
                        "count(distinct t.cveter) total " +
                        "from estadistica.tblterjuzgados t " +
                        "inner join estadistica.tblinijuzgados i on i.cveini = t.cveini " +
                        "inner join estadistica.tblrepjuidel r on i.cveini = r.cveini " +
                        "inner join estadistica.util_tblcatalogos c on c.cveCatalogo = r.CveJuicioDelito " +
                        "inner join estadistica.tbljuzgados j on j.cveAdscripcion = i.cveJuzgado " +
                        "where t.fechater between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                        "and i.Estado = 1 and t.CveResolucion in ('A031002') " +
                        "group by matJuzgado, matJuicio) as t " +
                        "group by t.matjuzgado " +
                        "order by field(t.matjuzgado,'JUZGADOS CIVILES','JUZGADOS USUCAPION','JUZGADOS MERCANTILES','JUZGADOS FAMILIARES') ";

                MySqlCommand cmd = new MySqlCommand(query, con);
                cmd.CommandTimeout = 1800;
                MySqlDataReader r = cmd.ExecuteReader();


                while (r.Read())
                {
                    JuicioMateriaTotales j = new JuicioMateriaTotales();
                    j.TipoJuzgado = r.GetString("tipoJuzgado");
                    j.CivilOral = r.GetString("CO");
                    j.CivilTradicional = r.GetString("CT");
                    j.FamiliarlOral = r.GetString("FO");
                    j.FamiliarTradicional = r.GetString("FT");
                    j.MercantilOral1 = r.GetString("MO");
                    j.MercantilTradicional1 = r.GetString("MT");
                    re.Jui.Add(j);

                }
                con.Close();

                /*Civil*/
                List<JuicioMateriaTotales> civiles = new List<JuicioMateriaTotales>();
                civiles = re.Jui;
                int col = 51;

                foreach (JuicioMateriaTotales c in civiles)
                {
                    worksheet.Cells[col, 5].Value = c.CivilOral;
                    col++;
                }

                List<JuicioMateriaTotales> civilesT = new List<JuicioMateriaTotales>();
                civilesT = re.Jui;
                int col1 = 51;

                foreach (JuicioMateriaTotales c in civilesT)
                {
                    worksheet.Cells[col1, 6].Value = c.CivilTradicional;
                    col1++;
                }

                /*Familiar*/
                List<JuicioMateriaTotales> familiar = new List<JuicioMateriaTotales>();
                familiar = re.Jui;
                int col2 = 51;

                foreach (JuicioMateriaTotales c in familiar)
                {
                    worksheet.Cells[col2, 7].Value = c.FamiliarlOral;
                    col2++;
                }

                List<JuicioMateriaTotales> familiarT = new List<JuicioMateriaTotales>();
                familiarT = re.Jui;
                int col3 = 51;

                foreach (JuicioMateriaTotales c in familiarT)
                {
                    worksheet.Cells[col3, 8].Value = c.FamiliarTradicional;
                    col3++;
                }

                /*Mercantil*/
                List<JuicioMateriaTotales> mercantil = new List<JuicioMateriaTotales>();
                mercantil = re.Jui;
                int col4 = 51;

                foreach (JuicioMateriaTotales c in mercantil)
                {
                    worksheet.Cells[col4, 12].Value = c.MercantilOral1;
                    col4++;
                }

                List<JuicioMateriaTotales> mercantilT = new List<JuicioMateriaTotales>();
                mercantilT = re.Jui;
                int col5 = 51;

                foreach (JuicioMateriaTotales c in mercantilT)
                {
                    worksheet.Cells[col5, 13].Value = c.MercantilTradicional1;
                    col5++;
                }

                //worksheet.Cells[53, 9].Value = 0;


                excel.Save();
                excel.Dispose();
                Debug.WriteLine("Se termino el proceso 4 ...");
            }
            catch (Exception i)
            {
                Debug.WriteLine("No se establecio la conexion :" + i);

            }
        }

        public void PrimerReporteCuadro5()
        {
            string fichero = "C:\\Reportes/ReporteUno.xlsx";
            ExcelPackage excel = new ExcelPackage(new FileInfo(fichero));
            ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];

            RepoUno re = new RepoUno();

            try
            {

                MySqlConnection con = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con.Open();

                query = "select t.matjuzgado tipoJuzgado, " +
                "ifnull(sum(if (t.matJuicio = 'CIVIL', t.total,null)),0) 'C', " +
                "ifnull(sum(if (t.matJuicio = 'FAMILIAR', t.total,null)),0) 'F', " +
                "ifnull(sum(if (t.matJuicio = 'MERCANTIL', t.total,null)),0) 'M' " +
                "from " +
                "(select " +
                "case " +
                "When jj.NomJuzgado like '%Civil%' or jj.NomJuzgado like '%Menor%' or jj.NomJuzgado like '%Mixto%'  or jj.NomJuzgado like '%C.M.%'or jj.NomJuzgado like '%C. M.%' or jj.NomJuzgado like '%Mercantil%' or jj.NomJuzgado like '%Usucapion%' then 'SALAS CIVILES' " +
                "When jj.NomJuzgado like '%Familiar%' or jj.nomJuzgado like '%Adopcion%' then 'SALAS FAMILIARES' " +
                "Else jj.NomJuzgado  end  matJuzgado, " +
                "case " +
                "when r.CveJuicioDelito = 'A102242' then 'CIVIL' " +
                "when r.CveJuicioDelito = 'A103024' or((c.tipo = 'C' or c.tipo like 'C%') and r.CveJuicioDelito != 'A102242') then 'CIVIL' " +
                "when((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121', 'A103067') and(c.descripcion not like '%Sucesorio%' and c.descripcion not like '%Tramitacion%' " +
                "and c.descripcion not like '%testament%' and c.descripcion not like '%here%')) then 'FAMILIAR' " +
                "when(((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121')  and(c.descripcion like '%Sucesorio%' or c.descripcion like '%Tramitacion%' " +
                "or c.descripcion like '%testament%' or c.descripcion like '%here%') or r.CveJuicioDelito = 'A103067')) then 'FAMILIAR' " +
                "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion like '%oral%') then 'MERCANTIL' " +
                "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion not like '%oral%') or r.cveJuicioDelito = 'A103121' then 'MERCANTIL' " +
                "else if (c.descripcion = 'NO EXISTE EXPEDIENTE', " +
                "(case when j.NomJuzgado like '%Civil%' or j.NomJuzgado like '%Menor%' or j.NomJuzgado like '%Mixto%'  or j.NomJuzgado like '%C.M.%' or j.NomJuzgado like '%Usucapion%'  then 'CIVIL' " +
                "when j.NomJuzgado like '%Familiar%' or j.nomJuzgado like '%Adopcion%' then 'FAMILIAR' " +
                "when j.NomJuzgado like '%Mercantil%' then 'MERCANTIL' " +
                "else j.NomJuzgado end), c.descripcion) " +
                "end matJuicio, " +
                "count(distinct i.cveini) total " +
                "from estadistica.tblinisalas i " +
                "inner join estadistica.tblinijuzgados ii on ii.cveExp = i.cveExp and ii.anioexp = i.anioexp and i.cveJuzProc = ii.cveJuzgado " +
                "inner join estadistica.tblrepjuidel r on ii.cveini = r.cveini " +
                "inner join estadistica.util_tblcatalogos c on c.cveCatalogo = r.CveJuicioDelito " +
                "inner join estadistica.tbljuzgados j on j.cveAdscripcion = ii.cveJuzgado " +
                "inner join estadistica.tbljuzgados jj on jj.cveAdscripcion = i.cveSala " +
                "where i.fechaRad between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                "and  i.Estado = 1 " +
                "group by matjuzgado, matJuicio) as t " +
                "group by t.matjuzgado " +
                "order by field(t.matjuzgado,'SALAS CIVILES','SALAS FAMILIARES'); ";

                MySqlCommand cmd = new MySqlCommand(query, con);
                cmd.CommandTimeout = 1800;
                MySqlDataReader r = cmd.ExecuteReader();


                while (r.Read())
                {
                    JuicioMateriaTotales j = new JuicioMateriaTotales();
                    j.TipoJuzgado = r.GetString("tipoJuzgado");
                    j.Civil = r.GetString("C");
                    j.Familiar = r.GetString("F");
                    j.Mercantil = r.GetString("M");
                    re.Jui.Add(j);

                }
                con.Close();

                /*Civil*/
                List<JuicioMateriaTotales> civiles = new List<JuicioMateriaTotales>();
                civiles = re.Jui;
                int col = 66;

                foreach (JuicioMateriaTotales c in civiles)
                {
                    worksheet.Cells[col, 5].Value = c.Civil;
                    col++;
                }



                /*Familiar*/
                List<JuicioMateriaTotales> familiar = new List<JuicioMateriaTotales>();
                familiar = re.Jui;
                int col2 = 66;

                foreach (JuicioMateriaTotales c in familiar)
                {
                    worksheet.Cells[col2, 6].Value = c.Familiar;
                    col2++;
                }



                /*Mercantil*/
                List<JuicioMateriaTotales> mercantil = new List<JuicioMateriaTotales>();
                mercantil = re.Jui;
                int col4 = 66;

                foreach (JuicioMateriaTotales c in mercantil)
                {
                    worksheet.Cells[col4, 10].Value = c.Mercantil;
                    col4++;
                }

                /*PENAL*/
                MySqlConnection con1 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con1.Open();

                query = "select " +
                        "SUM(if (t.matjuzgado = 'ORAL', t.total,0)) 'PENAL ORAL', " +
                        "SUM(if (t.matjuzgado = 'TRADICIONAL', t.total,0)) 'PENAL TRADICIONAL' " +
                        "from " +
                        "(select case " +
                        "when j.NomJuzgado like '%CONTROL%' or j.NomJuzgado like '%JUICIO ORAL%' or j.NomJuzgado like '%TRIBUNAL ORAL%' or j.NomJuzgado like '%EJECUCION DE SENTENCIAS SIST.%'then 'ORAL' " +
                        "when ii.cveini is not null or j.NomJuzgado like '%PENAL%' or j.NomJuzgado like  '%EJECUCION DE SENTENCIAS DE%' then 'TRADICIONAL' " +
                        "when(j.NomJuzgado like '%MIXTO%' or j.NomJuzgado like '%C.M%' or j.NomJuzgado like '%C. M%') and(jj.NomJuzgado like '%ALZADA%' or jj.NomJuzgado like '%PENAL%')  then 'TRADICIONAL' " +
                        "end matJuzgado, " +
                        "count(distinct i.cveini) total " +
                        "from estadistica.tblinisalas i " +
                        "left " +
                        "join estadistica.tblinijuzpen ii on ii.cveExp = i.cveExp and ii.anioexp = i.anioexp and i.cveJuzProc = ii.cveJuzgado " +
                        "inner join estadistica.tbljuzgados j on j.cveAdscripcion = i.cveJuzProc " +
                        "inner join estadistica.tbljuzgados jj on jj.cveAdscripcion = i.cvesala " +
                        "where i.fechaRad between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                        "and  i.Estado = 1 " +
                        "group by matJuzgado " +
                        "having matjuzgado is not null) t; ";

                MySqlCommand cmd1 = new MySqlCommand(query, con1);
                cmd1.CommandTimeout = 1800;
                MySqlDataReader r1 = cmd1.ExecuteReader();


                while (r1.Read())
                {


                    re.NsjPenal = r1.GetString("PENAL ORAL");
                    re.Penal = r1.GetString("PENAL TRADICIONAL");


                }
                con.Close();


                worksheet.Cells[68, 3].Value = re.NsjPenal;
                worksheet.Cells[68, 4].Value = re.Penal;


                excel.Save();
                excel.Dispose();
                Debug.WriteLine("Se termino el proceso 5 ...");
            }
            catch (Exception i)
            {
                Debug.WriteLine("No se establecio la conexion :" + i);

            }
        }

        public void PrimerReporteCuadro6()
        {
            string fichero = "C:\\Reportes/ReporteUno.xlsx";
            ExcelPackage excel = new ExcelPackage(new FileInfo(fichero));
            ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];

            RepoUno re = new RepoUno();

            try
            {

                MySqlConnection con = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con.Open();

                query = "Create Temporary Table IF NOT EXISTS estadistica.util_tblcatalogos " +
                        "(cveCatalogo char(7) COLLATE utf8_general_ci, " +
                        "descripcion varchar(200) CHARSET utf8 COLLATE utf8_general_ci, " +
                        "tipo char(2), " +
                        "INDEX cveCatalogo_util_idx(cveCatalogo), " +
                        "INDEX cveCatalogo_descripcion_idx(descripcion), " +
                        "INDEX cveCatalogo_tipo_idx(tipo)) " +
                        "(Select concat(codigo, rango) cveCatalogo, descripcion, tipo " +
                        "From estadistica.tblcatalogos c); " +
                        "select t.matjuzgado tipoJuzgado, " +
                        "ifnull(sum(if (t.matJuicio = 'CIVIL', t.total,null)),0) 'C', " +
                        "ifnull(sum(if (t.matJuicio = 'FAMILIAR', t.total,null)),0) 'F', " +
                        "ifnull(sum(if (t.matJuicio = 'MERCANTIL', t.total,null)),0) 'M' " +
                        "from " +
                        "(select " +
                        "case " +
                        "When jj.NomJuzgado like '%Civil%' or jj.NomJuzgado like '%Menor%' or jj.NomJuzgado like '%Mixto%'  or jj.NomJuzgado like '%C.M.%'or jj.NomJuzgado like '%C. M.%' or jj.NomJuzgado like '%Mercantil%' or jj.NomJuzgado like '%Usucapion%' then 'SALAS CIVILES' " +
                        "When jj.NomJuzgado like '%Familiar%' or jj.nomJuzgado like '%Adopcion%' then 'SALAS FAMILIARES' " +
                        "Else jj.NomJuzgado  end  matJuzgado, " +
                        "case " +
                        "when r.CveJuicioDelito = 'A102242' then 'CIVIL' " +
                        "when r.CveJuicioDelito = 'A103024' or((c.tipo = 'C' or c.tipo like 'C%') and r.CveJuicioDelito != 'A102242') then 'CIVIL' " +
                        "when((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121', 'A103067') and(c.descripcion not like '%Sucesorio%' and c.descripcion not like '%Tramitacion%' " +
                        "and c.descripcion not like '%testament%' and c.descripcion not like '%here%')) then 'FAMILIAR' " +
                        "when(((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121')  and(c.descripcion like '%Sucesorio%' or c.descripcion like '%Tramitacion%' " +
                        "or c.descripcion like '%testament%' or c.descripcion like '%here%') or r.CveJuicioDelito = 'A103067')) then 'FAMILIAR' " +
                        "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion like '%oral%') then 'MERCANTIL' " +
                        "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion not like '%oral%') or r.cveJuicioDelito = 'A103121' then 'MERCANTIL' " +
                        "else if (c.descripcion = 'NO EXISTE EXPEDIENTE', " +
                        "(case when j.NomJuzgado like '%Civil%' or j.NomJuzgado like '%Menor%' or j.NomJuzgado like '%Mixto%'  or j.NomJuzgado like '%C.M.%' or j.NomJuzgado like '%Usucapion%'  then 'CIVIL' " +
                        "when j.NomJuzgado like '%Familiar%' or j.nomJuzgado like '%Adopcion%' then 'FAMILIAR' " +
                        "when j.NomJuzgado like '%Mercantil%' then 'MERCANTIL' " +
                        "else j.NomJuzgado end), c.descripcion) " +
                        "end matJuicio, " +
                        "count(distinct t.cveter) total " +
                        "from estadistica.tblinisalas i " +
                        "inner join estadistica.tbltersalas t on t.cveini = i.cveini " +
                        "inner join estadistica.tblinijuzgados ii on ii.cveExp = i.cveExp and ii.anioexp = i.anioexp and i.cveJuzProc = ii.cveJuzgado " +
                        "inner join estadistica.tblrepjuidel r on ii.cveini = r.cveini " +
                        "inner join estadistica.util_tblcatalogos c on c.cveCatalogo = r.CveJuicioDelito " +
                        "inner join estadistica.tbljuzgados j on j.cveAdscripcion = ii.cveJuzgado " +
                        "inner join estadistica.tbljuzgados jj on jj.cveAdscripcion = i.cveSala " +
                        "where t.fechaTer between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                        "and  i.Estado = 1 and t.cveResolucion not like 'A053%' " +
                        "group by matjuzgado, matJuicio) as t " +
                        "group by t.matjuzgado " +
                        "order by field(t.matjuzgado,'SALAS CIVILES','SALAS FAMILIARES'); ";

                MySqlCommand cmd = new MySqlCommand(query, con);
                cmd.CommandTimeout = 1800;
                MySqlDataReader r = cmd.ExecuteReader();


                while (r.Read())
                {
                    JuicioMateriaTotales j = new JuicioMateriaTotales();
                    j.TipoJuzgado = r.GetString("tipoJuzgado");
                    j.Civil = r.GetString("C");
                    j.Familiar = r.GetString("F");
                    j.Mercantil = r.GetString("M");
                    re.Jui.Add(j);

                }
                con.Close();

                /*Civil*/
                List<JuicioMateriaTotales> civiles = new List<JuicioMateriaTotales>();
                civiles = re.Jui;
                int col = 75;

                foreach (JuicioMateriaTotales c in civiles)
                {
                    worksheet.Cells[col, 5].Value = c.Civil;
                    col++;
                }



                /*Familiar*/
                List<JuicioMateriaTotales> familiar = new List<JuicioMateriaTotales>();
                familiar = re.Jui;
                int col2 = 75;

                foreach (JuicioMateriaTotales c in familiar)
                {
                    worksheet.Cells[col2, 6].Value = c.Familiar;
                    col2++;
                }



                /*Mercantil*/
                List<JuicioMateriaTotales> mercantil = new List<JuicioMateriaTotales>();
                mercantil = re.Jui;
                int col4 = 75;

                foreach (JuicioMateriaTotales c in mercantil)
                {
                    worksheet.Cells[col4, 10].Value = c.Mercantil;
                    col4++;
                }

                /*PENAL*/
                MySqlConnection con1 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con1.Open();

                query = "select " +
                            "SUM(if (t.matjuzgado = 'ORAL', t.total,0)) 'PENAL ORAL', " +
                            "SUM(if (t.matjuzgado = 'TRADICIONAL', t.total,0)) 'PENAL TRADICIONAL' " +
                            "from " +
                            "(select case " +
                            "when j.NomJuzgado like '%CONTROL%' or j.NomJuzgado like '%JUICIO ORAL%' or j.NomJuzgado like '%TRIBUNAL ORAL%' or j.NomJuzgado like '%EJECUCION DE SENTENCIAS SIST.%'then 'ORAL' " +
                            "when ii.cveini is not null or j.NomJuzgado like '%PENAL%' or j.NomJuzgado like  '%EJECUCION DE SENTENCIAS DE%' then 'TRADICIONAL' " +
                            "when(j.NomJuzgado like '%MIXTO%' or j.NomJuzgado like '%C.M%' or j.NomJuzgado like '%C. M%') and(jj.NomJuzgado like '%ALZADA%' or jj.NomJuzgado like '%PENAL%')  then 'TRADICIONAL' " +
                            "end matJuzgado, " +
                            "count(distinct t.cveter) total " +
                            "from estadistica.tblinisalas i " +
                            "inner " +
                            "join estadistica.tbltersalas t on t.cveini = i.cveini " +
                            "left " +
                            "join estadistica.tblinijuzpen ii on ii.cveExp = i.cveExp and ii.anioexp = i.anioexp and i.cveJuzProc = ii.cveJuzgado " +
                            "inner join estadistica.tbljuzgados j on j.cveAdscripcion = i.cveJuzProc " +
                            "inner join estadistica.tbljuzgados jj on jj.cveAdscripcion = i.cvesala " +
                            "where t.fechaTer between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                            "and  i.Estado = 1 and t.cveResolucion not like '%A063%' " +
                            "group by matJuzgado " +
                            "having matjuzgado is not null) t; ";

                MySqlCommand cmd1 = new MySqlCommand(query, con1);
                cmd1.CommandTimeout = 1800;
                MySqlDataReader r1 = cmd1.ExecuteReader();


                while (r1.Read())
                {


                    re.NsjPenal = r1.GetString("PENAL ORAL");
                    re.Penal = r1.GetString("PENAL TRADICIONAL");


                }
                con.Close();


                worksheet.Cells[77, 3].Value = re.NsjPenal;
                worksheet.Cells[77, 4].Value = re.Penal;


                excel.Save();
                excel.Dispose();
                Debug.WriteLine("Se termino el proceso 6 ...");
            }
            catch (Exception i)
            {
                Debug.WriteLine("No se establecio la conexion :" + i);

            }
        }

        public void PrimerReporteCuadro7()
        {
            string fichero = "C:\\Reportes/ReporteUno.xlsx";
            ExcelPackage excel = new ExcelPackage(new FileInfo(fichero));
            ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];

            RepoUno re = new RepoUno();

            try
            {

                MySqlConnection con = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con.Open();

                query = "Create Temporary Table IF NOT EXISTS estadistica.util_tblcatalogos " +
                        "(cveCatalogo char(7) COLLATE utf8_general_ci, " +
                        "descripcion varchar(200) CHARSET utf8 COLLATE utf8_general_ci, " +
                        "tipo char(2), " +
                        "INDEX cveCatalogo_util_idx(cveCatalogo), " +
                        "INDEX cveCatalogo_descripcion_idx(descripcion), " +
                        "INDEX cveCatalogo_tipo_idx(tipo)) " +
                        "(Select concat(codigo, rango) cveCatalogo, descripcion, tipo " +
                        "From estadistica.tblcatalogos c); " +
                        "select t.matjuzgado tipoJuzgado, " +
                        "ifnull(sum(if (t.matJuicio = 'CIVIL', t.total,null)),0) 'C', " +
                        "ifnull(sum(if (t.matJuicio = 'FAMILIAR', t.total,null)),0) 'F', " +
                        "ifnull(sum(if (t.matJuicio = 'MERCANTIL', t.total,null)),0) 'M' " +
                        "from " +
                        "(select " +
                        "case " +
                        "When jj.NomJuzgado like '%Civil%' or jj.NomJuzgado like '%Menor%' or jj.NomJuzgado like '%Mixto%'  or jj.NomJuzgado like '%C.M.%'or jj.NomJuzgado like '%C. M.%' or jj.NomJuzgado like '%Mercantil%' or jj.NomJuzgado like '%Usucapion%' then 'SALAS CIVILES' " +
                        "When jj.NomJuzgado like '%Familiar%' or jj.nomJuzgado like '%Adopcion%' then 'SALAS FAMILIARES' " +
                        "Else jj.NomJuzgado  end  matJuzgado, " +
                        "case " +
                        "when r.CveJuicioDelito = 'A102242' then 'CIVIL' " +
                        "when r.CveJuicioDelito = 'A103024' or((c.tipo = 'C' or c.tipo like 'C%') and r.CveJuicioDelito != 'A102242') then 'CIVIL' " +
                        "when((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121', 'A103067') and(c.descripcion not like '%Sucesorio%' and c.descripcion not like '%Tramitacion%' " +
                        "and c.descripcion not like '%testament%' and c.descripcion not like '%here%')) then 'FAMILIAR' " +
                        "when(((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121')  and(c.descripcion like '%Sucesorio%' or c.descripcion like '%Tramitacion%' " +
                        "or c.descripcion like '%testament%' or c.descripcion like '%here%') or r.CveJuicioDelito = 'A103067')) then 'FAMILIAR' " +
                        "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion like '%oral%') then 'MERCANTIL' " +
                        "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion not like '%oral%') or r.cveJuicioDelito = 'A103121' then 'MERCANTIL' " +
                        "else if (c.descripcion = 'NO EXISTE EXPEDIENTE', " +
                        "(case when j.NomJuzgado like '%Civil%' or j.NomJuzgado like '%Menor%' or j.NomJuzgado like '%Mixto%'  or j.NomJuzgado like '%C.M.%' or j.NomJuzgado like '%Usucapion%'  then 'CIVIL' " +
                        "when j.NomJuzgado like '%Familiar%' or j.nomJuzgado like '%Adopcion%' then 'FAMILIAR' " +
                        "when j.NomJuzgado like '%Mercantil%' then 'MERCANTIL' " +
                        "else j.NomJuzgado end), c.descripcion) " +
                        "end matJuicio, " +
                        "count(distinct t.cveter) total " +
                        "from estadistica.tblinisalas i " +
                        "inner join estadistica.tbltersalas t on t.cveini = i.cveini " +
                        "inner join estadistica.tblinijuzgados ii on ii.cveExp = i.cveExp and ii.anioexp = i.anioexp and i.cveJuzProc = ii.cveJuzgado " +
                        "inner join estadistica.tblrepjuidel r on ii.cveini = r.cveini " +
                        "inner join estadistica.util_tblcatalogos c on c.cveCatalogo = r.CveJuicioDelito " +
                        "inner join estadistica.tbljuzgados j on j.cveAdscripcion = ii.cveJuzgado " +
                        "inner join estadistica.tbljuzgados jj on jj.cveAdscripcion = i.cveSala " +
                        "where t.fechaTer between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                        "and  i.Estado = 1 and t.cveResolucion in ('A051001') " +
                        "group by matjuzgado, matJuicio) as t " +
                        "group by t.matjuzgado " +
                        "order by field(t.matjuzgado,'SALAS CIVILES','SALAS FAMILIARES'); ";

                MySqlCommand cmd = new MySqlCommand(query, con);
                cmd.CommandTimeout = 1800;
                MySqlDataReader r = cmd.ExecuteReader();


                while (r.Read())
                {
                    JuicioMateriaTotales j = new JuicioMateriaTotales();
                    j.TipoJuzgado = r.GetString("tipoJuzgado");
                    j.Civil = r.GetString("C");
                    j.Familiar = r.GetString("F");
                    j.Mercantil = r.GetString("M");
                    re.Jui.Add(j);

                }
                con.Close();

                /*Civil*/
                List<JuicioMateriaTotales> civiles = new List<JuicioMateriaTotales>();
                civiles = re.Jui;
                int col = 83;

                foreach (JuicioMateriaTotales c in civiles)
                {
                    worksheet.Cells[col, 5].Value = c.Civil;
                    col++;
                }



                /*Familiar*/
                List<JuicioMateriaTotales> familiar = new List<JuicioMateriaTotales>();
                familiar = re.Jui;
                int col2 = 83;

                foreach (JuicioMateriaTotales c in familiar)
                {
                    worksheet.Cells[col2, 6].Value = c.Familiar;
                    col2++;
                }



                /*Mercantil*/
                List<JuicioMateriaTotales> mercantil = new List<JuicioMateriaTotales>();
                mercantil = re.Jui;
                int col4 = 83;

                foreach (JuicioMateriaTotales c in mercantil)
                {
                    worksheet.Cells[col4, 10].Value = c.Mercantil;
                    col4++;
                }

                /*PENAL*/
                MySqlConnection con1 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con1.Open();

                query = "select " +
                        "SUM(if (t.matjuzgado = 'ORAL', t.total,0)) 'PENAL ORAL', " +
                        "SUM(if (t.matjuzgado = 'TRADICIONAL', t.total,0)) 'PENAL TRADICIONAL' " +
                        "from " +
                        "(select case " +
                        "when j.NomJuzgado like '%CONTROL%' or j.NomJuzgado like '%JUICIO ORAL%' or j.NomJuzgado like '%TRIBUNAL ORAL%' or j.NomJuzgado like '%EJECUCION DE SENTENCIAS SIST.%'then 'ORAL' " +
                        "when ii.cveini is not null or j.NomJuzgado like '%PENAL%' or j.NomJuzgado like  '%EJECUCION DE SENTENCIAS DE%' then 'TRADICIONAL' " +
                        "when(j.NomJuzgado like '%MIXTO%' or j.NomJuzgado like '%C.M%' or j.NomJuzgado like '%C. M%') and(jj.NomJuzgado like '%ALZADA%' or jj.NomJuzgado like '%PENAL%')  then 'TRADICIONAL' " +
                        "end matJuzgado, " +
                        "count(distinct t.cveter) total " +
                        "from estadistica.tblinisalas i " +
                        "inner " +
                        "join estadistica.tbltersalas t on t.cveini = i.cveini " +
                        "left " +
                        "join estadistica.tblinijuzpen ii on ii.cveExp = i.cveExp and ii.anioexp = i.anioexp and i.cveJuzProc = ii.cveJuzgado " +
                        "inner join estadistica.tbljuzgados j on j.cveAdscripcion = i.cveJuzProc " +
                        "inner join estadistica.tbljuzgados jj on jj.cveAdscripcion = i.cvesala " +
                        "where t.fechaTer between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                        "and  i.Estado = 1 and t.cveResolucion = 'A061002' " +
                        "group by matJuzgado " +
                        "having matjuzgado is not null) t; ";

                MySqlCommand cmd1 = new MySqlCommand(query, con1);
                cmd1.CommandTimeout = 1800;
                MySqlDataReader r1 = cmd1.ExecuteReader();


                while (r1.Read())
                {


                    re.NsjPenal = r1.GetString("PENAL ORAL");
                    re.Penal = r1.GetString("PENAL TRADICIONAL");


                }
                con.Close();


                worksheet.Cells[85, 3].Value = re.NsjPenal;
                worksheet.Cells[85, 4].Value = re.Penal;


                excel.Save();
                excel.Dispose();
                Debug.WriteLine("Se termino el proceso 7 ...");
            }
            catch (Exception i)
            {
                Debug.WriteLine("No se establecio la conexion :" + i);

            }
        }

        public void PrimerReporteCuadro8()
        {
            string fichero = "C:\\Reportes/ReporteUno.xlsx";
            ExcelPackage excel = new ExcelPackage(new FileInfo(fichero));
            ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];

            RepoUno re = new RepoUno();

            try
            {

                MySqlConnection con = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con.Open();

                query = "Create Temporary Table IF NOT EXISTS estadistica.util_tblcatalogos " +
                        "(cveCatalogo char(7) COLLATE utf8_general_ci, " +
                        "descripcion varchar(200) CHARSET utf8 COLLATE utf8_general_ci, " +
                        "tipo char(2), " +
                        "INDEX cveCatalogo_util_idx(cveCatalogo), " +
                        "INDEX cveCatalogo_descripcion_idx(descripcion), " +
                        "INDEX cveCatalogo_tipo_idx(tipo)) " +
                        "(Select concat(codigo, rango) cveCatalogo, descripcion, tipo " +
                        "From estadistica.tblcatalogos c); " +
                        "select t.matjuzgado tipoJuzgado, " +
                        "ifnull(sum(if (t.matJuicio = 'CIVIL', t.total,null)),0) 'C', " +
                        "ifnull(sum(if (t.matJuicio = 'FAMILIAR', t.total,null)),0) 'F', " +
                        "ifnull(sum(if (t.matJuicio = 'MERCANTIL', t.total,null)),0) 'M' " +
                        "from " +
                        "(select " +
                        "case " +
                        "When jj.NomJuzgado like '%Civil%' or jj.NomJuzgado like '%Menor%' or jj.NomJuzgado like '%Mixto%'  or jj.NomJuzgado like '%C.M.%'or jj.NomJuzgado like '%C. M.%' or jj.NomJuzgado like '%Mercantil%' or jj.NomJuzgado like '%Usucapion%' then 'SALAS CIVILES' " +
                        "When jj.NomJuzgado like '%Familiar%' or jj.nomJuzgado like '%Adopcion%' then 'SALAS FAMILIARES' " +
                        "Else jj.NomJuzgado  end  matJuzgado, " +
                        "case " +
                        "when r.CveJuicioDelito = 'A102242' then 'CIVIL' " +
                        "when r.CveJuicioDelito = 'A103024' or((c.tipo = 'C' or c.tipo like 'C%') and r.CveJuicioDelito != 'A102242') then 'CIVIL' " +
                        "when((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121', 'A103067') and(c.descripcion not like '%Sucesorio%' and c.descripcion not like '%Tramitacion%' " +
                        "and c.descripcion not like '%testament%' and c.descripcion not like '%here%')) then 'FAMILIAR' " +
                        "when(((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121')  and(c.descripcion like '%Sucesorio%' or c.descripcion like '%Tramitacion%' " +
                        "or c.descripcion like '%testament%' or c.descripcion like '%here%') or r.CveJuicioDelito = 'A103067')) then 'FAMILIAR' " +
                        "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion like '%oral%') then 'MERCANTIL' " +
                        "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion not like '%oral%') or r.cveJuicioDelito = 'A103121' then 'MERCANTIL' " +
                        "else if (c.descripcion = 'NO EXISTE EXPEDIENTE', " +
                        "(case when j.NomJuzgado like '%Civil%' or j.NomJuzgado like '%Menor%' or j.NomJuzgado like '%Mixto%'  or j.NomJuzgado like '%C.M.%' or j.NomJuzgado like '%Usucapion%'  then 'CIVIL' " +
                        "when j.NomJuzgado like '%Familiar%' or j.nomJuzgado like '%Adopcion%' then 'FAMILIAR' " +
                        "when j.NomJuzgado like '%Mercantil%' then 'MERCANTIL' " +
                        "else j.NomJuzgado end), c.descripcion) " +
                        "end matJuicio, " +
                        "count(distinct t.cveter) total " +
                        "from estadistica.tblinisalas i " +
                        "inner join estadistica.tbltersalas t on t.cveini = i.cveini " +
                        "inner join estadistica.tblinijuzgados ii on ii.cveExp = i.cveExp and ii.anioexp = i.anioexp and i.cveJuzProc = ii.cveJuzgado " +
                        "inner join estadistica.tblrepjuidel r on ii.cveini = r.cveini " +
                        "inner join estadistica.util_tblcatalogos c on c.cveCatalogo = r.CveJuicioDelito " +
                        "inner join estadistica.tbljuzgados j on j.cveAdscripcion = ii.cveJuzgado " +
                        "inner join estadistica.tbljuzgados jj on jj.cveAdscripcion = i.cveSala " +
                        "where t.fechaTer between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                        "and  i.Estado = 1 and t.cveResolucion in ('A051003') " +
                        "group by matjuzgado, matJuicio) as t " +
                        "group by t.matjuzgado " +
                        "order by field(t.matjuzgado,'SALAS CIVILES','SALAS FAMILIARES'); ";

                MySqlCommand cmd = new MySqlCommand(query, con);
                cmd.CommandTimeout = 1800;
                MySqlDataReader r = cmd.ExecuteReader();


                while (r.Read())
                {
                    JuicioMateriaTotales j = new JuicioMateriaTotales();
                    j.TipoJuzgado = r.GetString("tipoJuzgado");
                    j.Civil = r.GetString("C");
                    j.Familiar = r.GetString("F");
                    j.Mercantil = r.GetString("M");
                    re.Jui.Add(j);

                }
                con.Close();

                /*Civil*/
                List<JuicioMateriaTotales> civiles = new List<JuicioMateriaTotales>();
                civiles = re.Jui;
                int col = 91;

                foreach (JuicioMateriaTotales c in civiles)
                {
                    worksheet.Cells[col, 5].Value = c.Civil;
                    col++;
                }



                /*Familiar*/
                List<JuicioMateriaTotales> familiar = new List<JuicioMateriaTotales>();
                familiar = re.Jui;
                int col2 = 91;

                foreach (JuicioMateriaTotales c in familiar)
                {
                    worksheet.Cells[col2, 6].Value = c.Familiar;
                    col2++;
                }



                /*Mercantil*/
                List<JuicioMateriaTotales> mercantil = new List<JuicioMateriaTotales>();
                mercantil = re.Jui;
                int col4 = 91;

                foreach (JuicioMateriaTotales c in mercantil)
                {
                    worksheet.Cells[col4, 10].Value = c.Mercantil;
                    col4++;
                }

                /*PENAL*/
                MySqlConnection con1 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con1.Open();

                query = "select " +
                            "SUM(if (t.matjuzgado = 'ORAL', t.total,0)) 'PENAL ORAL', " +
                            "SUM(if (t.matjuzgado = 'TRADICIONAL', t.total,0)) 'PENAL TRADICIONAL' " +
                            "from " +
                            "(select case " +
                            "when j.NomJuzgado like '%CONTROL%' or j.NomJuzgado like '%JUICIO ORAL%' or j.NomJuzgado like '%TRIBUNAL ORAL%' or j.NomJuzgado like '%EJECUCION DE SENTENCIAS SIST.%'then 'ORAL' " +
                            "when ii.cveini is not null or j.NomJuzgado like '%PENAL%' or j.NomJuzgado like  '%EJECUCION DE SENTENCIAS DE%' then 'TRADICIONAL' " +
                            "when(j.NomJuzgado like '%MIXTO%' or j.NomJuzgado like '%C.M%' or j.NomJuzgado like '%C. M%') and(jj.NomJuzgado like '%ALZADA%' or jj.NomJuzgado like '%PENAL%')  then 'TRADICIONAL' " +
                            "end matJuzgado, " +
                            "count(distinct t.cveter) total " +
                            "from estadistica.tblinisalas i " +
                            "inner " +
                            "join estadistica.tbltersalas t on t.cveini = i.cveini " +
                            "left " +
                            "join estadistica.tblinijuzpen ii on ii.cveExp = i.cveExp and ii.anioexp = i.anioexp and i.cveJuzProc = ii.cveJuzgado " +
                            "inner join estadistica.tbljuzgados j on j.cveAdscripcion = i.cveJuzProc " +
                            "inner join estadistica.tbljuzgados jj on jj.cveAdscripcion = i.cvesala " +
                            "where t.fechaTer between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                            "and  i.Estado = 1 and t.cveResolucion = 'A061001' " +
                            "group by matJuzgado " +
                            "having matjuzgado is not null) t; ";

                MySqlCommand cmd1 = new MySqlCommand(query, con1);
                cmd1.CommandTimeout = 1800;
                MySqlDataReader r1 = cmd1.ExecuteReader();


                while (r1.Read())
                {


                    re.NsjPenal = r1.GetString("PENAL ORAL");
                    re.Penal = r1.GetString("PENAL TRADICIONAL");


                }
                con.Close();


                worksheet.Cells[93, 3].Value = re.NsjPenal;
                worksheet.Cells[93, 4].Value = re.Penal;


                excel.Save();
                excel.Dispose();
                Debug.WriteLine("Se termino el proceso 8 ...");
            }
            catch (Exception i)
            {
                Debug.WriteLine("No se establecio la conexion :" + i);

            }
        }

        public void PrimerReporteCuadro9()
        {
            string fichero = "C:\\Reportes/ReporteUno.xlsx";
            ExcelPackage excel = new ExcelPackage(new FileInfo(fichero));
            ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];

            RepoUno re = new RepoUno();

            try
            {

                MySqlConnection con = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con.Open();

                query = "Create Temporary Table IF NOT EXISTS estadistica.util_tblcatalogos " +
                        "(cveCatalogo char(7) COLLATE utf8_general_ci, " +
                        "descripcion varchar(200) CHARSET utf8 COLLATE utf8_general_ci, " +
                        "tipo char(2), " +
                        "INDEX cveCatalogo_util_idx(cveCatalogo), " +
                        "INDEX cveCatalogo_descripcion_idx(descripcion), " +
                        "INDEX cveCatalogo_tipo_idx(tipo)) " +
                        "(Select concat(codigo, rango) cveCatalogo, descripcion, tipo " +
                        "From estadistica.tblcatalogos c); " +
                        "select t.matjuzgado tipoJuzgado, " +
                        "ifnull(sum(if (t.matJuicio = 'CIVIL', t.total,null)),0) 'C', " +
                        "ifnull(sum(if (t.matJuicio = 'FAMILIAR', t.total,null)),0) 'F', " +
                        "ifnull(sum(if (t.matJuicio = 'MERCANTIL', t.total,null)),0) 'M' " +
                        "from " +
                        "(select " +
                        "case " +
                        "When jj.NomJuzgado like '%Civil%' or jj.NomJuzgado like '%Menor%' or jj.NomJuzgado like '%Mixto%'  or jj.NomJuzgado like '%C.M.%'or jj.NomJuzgado like '%C. M.%' or jj.NomJuzgado like '%Mercantil%' or jj.NomJuzgado like '%Usucapion%' then 'SALAS CIVILES' " +
                        "When jj.NomJuzgado like '%Familiar%' or jj.nomJuzgado like '%Adopcion%' then 'SALAS FAMILIARES' " +
                        "Else jj.NomJuzgado  end  matJuzgado, " +
                        "case " +
                        "when r.CveJuicioDelito = 'A102242' then 'CIVIL' " +
                        "when r.CveJuicioDelito = 'A103024' or((c.tipo = 'C' or c.tipo like 'C%') and r.CveJuicioDelito != 'A102242') then 'CIVIL' " +
                        "when((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121', 'A103067') and(c.descripcion not like '%Sucesorio%' and c.descripcion not like '%Tramitacion%' " +
                        "and c.descripcion not like '%testament%' and c.descripcion not like '%here%')) then 'FAMILIAR' " +
                        "when(((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121')  and(c.descripcion like '%Sucesorio%' or c.descripcion like '%Tramitacion%' " +
                        "or c.descripcion like '%testament%' or c.descripcion like '%here%') or r.CveJuicioDelito = 'A103067')) then 'FAMILIAR' " +
                        "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion like '%oral%') then 'MERCANTIL' " +
                        "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion not like '%oral%') or r.cveJuicioDelito = 'A103121' then 'MERCANTIL' " +
                        "else if (c.descripcion = 'NO EXISTE EXPEDIENTE', " +
                        "(case when j.NomJuzgado like '%Civil%' or j.NomJuzgado like '%Menor%' or j.NomJuzgado like '%Mixto%'  or j.NomJuzgado like '%C.M.%' or j.NomJuzgado like '%Usucapion%'  then 'CIVIL' " +
                        "when j.NomJuzgado like '%Familiar%' or j.nomJuzgado like '%Adopcion%' then 'FAMILIAR' " +
                        "when j.NomJuzgado like '%Mercantil%' then 'MERCANTIL' " +
                        "else j.NomJuzgado end), c.descripcion) " +
                        "end matJuicio, " +
                        "count(distinct t.cveter) total " +
                        "from estadistica.tblinisalas i " +
                        "inner join estadistica.tbltersalas t on t.cveini = i.cveini " +
                        "inner join estadistica.tblinijuzgados ii on ii.cveExp = i.cveExp and ii.anioexp = i.anioexp and i.cveJuzProc = ii.cveJuzgado " +
                        "inner join estadistica.tblrepjuidel r on ii.cveini = r.cveini " +
                        "inner join estadistica.util_tblcatalogos c on c.cveCatalogo = r.CveJuicioDelito " +
                        "inner join estadistica.tbljuzgados j on j.cveAdscripcion = ii.cveJuzgado " +
                        "inner join estadistica.tbljuzgados jj on jj.cveAdscripcion = i.cveSala " +
                        "where t.fechaTer between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                        "and  i.Estado = 1 and t.cveResolucion in ('A051002') " +
                        "group by matjuzgado, matJuicio) as t " +
                        "group by t.matjuzgado " +
                        "order by field(t.matjuzgado,'SALAS CIVILES','SALAS FAMILIARES'); ";

                MySqlCommand cmd = new MySqlCommand(query, con);
                cmd.CommandTimeout = 1800;
                MySqlDataReader r = cmd.ExecuteReader();


                while (r.Read())
                {
                    JuicioMateriaTotales j = new JuicioMateriaTotales();
                    j.TipoJuzgado = r.GetString("tipoJuzgado");
                    j.Civil = r.GetString("C");
                    j.Familiar = r.GetString("F");
                    j.Mercantil = r.GetString("M");
                    re.Jui.Add(j);

                }
                con.Close();

                /*Civil*/
                List<JuicioMateriaTotales> civiles = new List<JuicioMateriaTotales>();
                civiles = re.Jui;
                int col = 99;

                foreach (JuicioMateriaTotales c in civiles)
                {
                    worksheet.Cells[col, 5].Value = c.Civil;
                    col++;
                }



                /*Familiar*/
                List<JuicioMateriaTotales> familiar = new List<JuicioMateriaTotales>();
                familiar = re.Jui;
                int col2 = 99;

                foreach (JuicioMateriaTotales c in familiar)
                {
                    worksheet.Cells[col2, 6].Value = c.Familiar;
                    col2++;
                }



                /*Mercantil*/
                List<JuicioMateriaTotales> mercantil = new List<JuicioMateriaTotales>();
                mercantil = re.Jui;
                int col4 = 99;

                foreach (JuicioMateriaTotales c in mercantil)
                {
                    worksheet.Cells[col4, 10].Value = c.Mercantil;
                    col4++;
                }

                /*PENAL*/
                MySqlConnection con1 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con1.Open();

                query = "select " +
                        "SUM(if (t.matjuzgado = 'ORAL', t.total,0)) 'PENAL ORAL', " +
                        "SUM(if (t.matjuzgado = 'TRADICIONAL', t.total,0)) 'PENAL TRADICIONAL' " +
                        "from " +
                        "(select case " +
                        "when j.NomJuzgado like '%CONTROL%' or j.NomJuzgado like '%JUICIO ORAL%' or j.NomJuzgado like '%TRIBUNAL ORAL%' or j.NomJuzgado like '%EJECUCION DE SENTENCIAS SIST.%'then 'ORAL' " +
                        "when ii.cveini is not null or j.NomJuzgado like '%PENAL%' or j.NomJuzgado like  '%EJECUCION DE SENTENCIAS DE%' then 'TRADICIONAL' " +
                        "when(j.NomJuzgado like '%MIXTO%' or j.NomJuzgado like '%C.M%' or j.NomJuzgado like '%C. M%') and(jj.NomJuzgado like '%ALZADA%' or jj.NomJuzgado like '%PENAL%')  then 'TRADICIONAL' " +
                        "end matJuzgado, " +
                        "count(distinct t.cveter) total " +
                        "from estadistica.tblinisalas i " +
                        "inner " +
                        "join estadistica.tbltersalas t on t.cveini = i.cveini " +
                        "left " +
                        "join estadistica.tblinijuzpen ii on ii.cveExp = i.cveExp and ii.anioexp = i.anioexp and i.cveJuzProc = ii.cveJuzgado " +
                        "inner join estadistica.tbljuzgados j on j.cveAdscripcion = i.cveJuzProc " +
                        "inner join estadistica.tbljuzgados jj on jj.cveAdscripcion = i.cvesala " +
                        "where t.fechaTer between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                        "and  i.Estado = 1 and t.cveResolucion = 'A061003' " +
                        "group by matJuzgado " +
                        "having matjuzgado is not null) t; ";

                MySqlCommand cmd1 = new MySqlCommand(query, con1);
                cmd1.CommandTimeout = 1800;
                MySqlDataReader r1 = cmd1.ExecuteReader();


                while (r1.Read())
                {


                    re.NsjPenal = r1.GetString("PENAL ORAL");
                    re.Penal = r1.GetString("PENAL TRADICIONAL");


                }
                con.Close();


                worksheet.Cells[101, 3].Value = re.NsjPenal;
                worksheet.Cells[101, 4].Value = re.Penal;


                excel.Save();
                excel.Dispose();
                Debug.WriteLine("Se termino el proceso 9 ...");
            }
            catch (Exception i)
            {
                Debug.WriteLine("No se establecio la conexion :" + i);

            }
        }

        public void PrimerReporteCuadro10()
        {
            string fichero = "C:\\Reportes/ReporteUno.xlsx";
            ExcelPackage excel = new ExcelPackage(new FileInfo(fichero));
            ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];

            RepoUno re = new RepoUno();

            try
            {

                MySqlConnection con = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con.Open();

                query = "Create Temporary Table IF NOT EXISTS estadistica.util_tblcatalogos " +
                        "(cveCatalogo char(7) COLLATE utf8_general_ci, " +
                        "descripcion varchar(200) CHARSET utf8 COLLATE utf8_general_ci, " +
                        "tipo char(2), " +
                        "INDEX cveCatalogo_util_idx(cveCatalogo), " +
                        "INDEX cveCatalogo_descripcion_idx(descripcion), " +
                        "INDEX cveCatalogo_tipo_idx(tipo)) " +
                        "(Select concat(codigo, rango) cveCatalogo, descripcion, tipo " +
                        "From estadistica.tblcatalogos c); " +
                            "select t.matjuzgado tipoJuzgado, " +
                            "ifnull(sum(if (t.matJuicio = 'CIVIL', t.total,null)),0) 'C', " +
                            "ifnull(sum(if (t.matJuicio = 'FAMILIAR', t.total,null)),0) 'F', " +
                            "ifnull(sum(if (t.matJuicio = 'MERCANTIL', t.total,null)),0) 'M' " +
                            "from " +
                            "(select " +
                            "case " +
                            "When jj.NomJuzgado like '%Civil%' or jj.NomJuzgado like '%Menor%' or jj.NomJuzgado like '%Mixto%'  or jj.NomJuzgado like '%C.M.%'or jj.NomJuzgado like '%C. M.%' or jj.NomJuzgado like '%Mercantil%' or jj.NomJuzgado like '%Usucapion%' then 'SALAS CIVILES' " +
                            "When jj.NomJuzgado like '%Familiar%' or jj.nomJuzgado like '%Adopcion%' then 'SALAS FAMILIARES' " +
                            "Else jj.NomJuzgado  end  matJuzgado, " +
                            "case " +
                            "when r.CveJuicioDelito = 'A102242' then 'CIVIL' " +
                            "when r.CveJuicioDelito = 'A103024' or((c.tipo = 'C' or c.tipo like 'C%') and r.CveJuicioDelito != 'A102242') then 'CIVIL' " +
                            "when((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121', 'A103067') and(c.descripcion not like '%Sucesorio%' and c.descripcion not like '%Tramitacion%' " +
                            "and c.descripcion not like '%testament%' and c.descripcion not like '%here%')) then 'FAMILIAR' " +
                            "when(((c.tipo LIKE 'F%' or c.tipo like 'F') and r.CveJuicioDelito not in  ('A103024', 'A103121')  and(c.descripcion like '%Sucesorio%' or c.descripcion like '%Tramitacion%' " +
                            "or c.descripcion like '%testament%' or c.descripcion like '%here%') or r.CveJuicioDelito = 'A103067')) then 'FAMILIAR' " +
                            "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion like '%oral%') then 'MERCANTIL' " +
                            "when((c.tipo like 'M%' or c.tipo like 'M') and c.descripcion not like '%oral%') or r.cveJuicioDelito = 'A103121' then 'MERCANTIL' " +
                            "else if (c.descripcion = 'NO EXISTE EXPEDIENTE', " +
                            "(case when j.NomJuzgado like '%Civil%' or j.NomJuzgado like '%Menor%' or j.NomJuzgado like '%Mixto%'  or j.NomJuzgado like '%C.M.%' or j.NomJuzgado like '%Usucapion%'  then 'CIVIL' " +
                            "when j.NomJuzgado like '%Familiar%' or j.nomJuzgado like '%Adopcion%' then 'FAMILIAR' " +
                            "when j.NomJuzgado like '%Mercantil%' then 'MERCANTIL' " +
                            "else j.NomJuzgado end), c.descripcion) " +
                            "end matJuicio, " +
                            "count(distinct i.cveini) total " +
                            "from estadistica.tblinisalas i " +
                            "inner join estadistica.tblinijuzgados ii on ii.cveExp = i.cveExp and ii.anioexp = i.anioexp and i.cveJuzProc = ii.cveJuzgado " +
                            "inner join estadistica.tblrepjuidel r on ii.cveini = r.cveini " +
                            "inner join estadistica.util_tblcatalogos c on c.cveCatalogo = r.CveJuicioDelito " +
                            "inner join estadistica.tbljuzgados j on j.cveAdscripcion = ii.cveJuzgado " +
                            "inner join estadistica.tbljuzgados jj on jj.cveAdscripcion = i.cveSala " +
                            "where i.fechaRad between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                            "and  i.Estado = 1 and i.NatRecurso like 'A021%' " +
                            "group by matjuzgado, matJuicio) as t " +
                            "group by t.matjuzgado " +
                            "order by field(t.matjuzgado,'SALAS CIVILES','SALAS FAMILIARES'); ";

                MySqlCommand cmd = new MySqlCommand(query, con);
                cmd.CommandTimeout = 1800;
                MySqlDataReader r = cmd.ExecuteReader();


                while (r.Read())
                {
                    JuicioMateriaTotales j = new JuicioMateriaTotales();
                    j.TipoJuzgado = r.GetString("tipoJuzgado");
                    j.Civil = r.GetString("C");
                    j.Familiar = r.GetString("F");
                    j.Mercantil = r.GetString("M");
                    re.Jui.Add(j);

                }
                con.Close();

                /*Civil*/
                List<JuicioMateriaTotales> civiles = new List<JuicioMateriaTotales>();
                civiles = re.Jui;
                int col = 108;

                foreach (JuicioMateriaTotales c in civiles)
                {
                    worksheet.Cells[col, 5].Value = c.Civil;
                    col++;
                }



                /*Familiar*/
                List<JuicioMateriaTotales> familiar = new List<JuicioMateriaTotales>();
                familiar = re.Jui;
                int col2 = 108;

                foreach (JuicioMateriaTotales c in familiar)
                {
                    worksheet.Cells[col2, 6].Value = c.Familiar;
                    col2++;
                }



                /*Mercantil*/
                List<JuicioMateriaTotales> mercantil = new List<JuicioMateriaTotales>();
                mercantil = re.Jui;
                int col4 = 108;

                foreach (JuicioMateriaTotales c in mercantil)
                {
                    worksheet.Cells[col4, 10].Value = c.Mercantil;
                    col4++;
                }

                /*PENAL*/
                MySqlConnection con1 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con1.Open();

                query = "select " +
                        "SUM(if (t.matjuzgado = 'ORAL', t.total,0)) 'PENAL ORAL', " +
                        "SUM(if (t.matjuzgado = 'TRADICIONAL', t.total,0)) 'PENAL TRADICIONAL' " +
                        "from " +
                        "(select case " +
                        "when j.NomJuzgado like '%CONTROL%' or j.NomJuzgado like '%JUICIO ORAL%' or j.NomJuzgado like '%TRIBUNAL ORAL%' or j.NomJuzgado like '%EJECUCION DE SENTENCIAS SIST.%'then 'ORAL' " +
                        "when ii.cveini is not null or j.NomJuzgado like '%PENAL%' or j.NomJuzgado like  '%EJECUCION DE SENTENCIAS DE%' then 'TRADICIONAL' " +
                        "when(j.NomJuzgado like '%MIXTO%' or j.NomJuzgado like '%C.M%' or j.NomJuzgado like '%C. M%') and(jj.NomJuzgado like '%ALZADA%' or jj.NomJuzgado like '%PENAL%')  then 'TRADICIONAL' " +
                        "end matJuzgado, " +
                        "count(distinct i.cveini) total " +
                        "from estadistica.tblinisalas i " +
                        "left " +
                        "join estadistica.tblinijuzpen ii on ii.cveExp = i.cveExp and ii.anioexp = i.anioexp and i.cveJuzProc = ii.cveJuzgado " +
                        "inner join estadistica.tbljuzgados j on j.cveAdscripcion = i.cveJuzProc " +
                        "inner join estadistica.tbljuzgados jj on jj.cveAdscripcion = i.cvesala " +
                        "where i.fechaRad between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                        "and  i.Estado = 1 and i.NatRecurso like '%A012%' " +
                        "group by matJuzgado " +
                        "having matjuzgado is not null) t; ";

                MySqlCommand cmd1 = new MySqlCommand(query, con1);
                cmd1.CommandTimeout = 1800;
                MySqlDataReader r1 = cmd1.ExecuteReader();


                while (r1.Read())
                {


                    re.NsjPenal = r1.GetString("PENAL ORAL");
                    re.Penal = r1.GetString("PENAL TRADICIONAL");


                }
                con.Close();


                worksheet.Cells[110, 3].Value = re.NsjPenal;
                worksheet.Cells[110, 4].Value = re.Penal;


                excel.Save();
                excel.Dispose();
                Debug.WriteLine("Se termino el proceso 10 ...");
                //mascara.Visible = false;
            }
            catch (Exception i)
            {
                Debug.WriteLine("No se establecio la conexion :" + i);

            }
        }

        public void PrimerReporteCuadro11()
        {
            string fichero = "C:\\Reportes/ReporteUno.xlsx";
            ExcelPackage excel = new ExcelPackage(new FileInfo(fichero));
            ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];

            RepoUno re = new RepoUno();

            try
            {

                MySqlConnection con = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con.Open();

                query = "select " +
                        "case " +
                        "When j.NomJuzgado like '%Civil%' or j.NomJuzgado like '%Menor%' or j.NomJuzgado like '%Mixto%'  or j.NomJuzgado like '%C.M.%'or j.NomJuzgado like '%C. M.%' then 'JUZGADOS CIVILES' " +
                        "When j.NomJuzgado like '%Mercantil%' then 'JUZGADOS MERCANTILES' " +
                        "When j.NomJuzgado like '%Familiar%' or j.nomJuzgado like '%Adopcion%' then 'JUZGADOS FAMILIARES' " +
                        "When j.NomJuzgado like '%Usucapion%' then 'JUZGADOS USUCAPION' " +
                        "Else j.NomJuzgado end  matJuzgado, " +
                        "sum(if (d.CveDatAdi = 'A072001', d.Total, null)) 'CELEBRADAS', " +
                        "sum(if (d.CveDatAdi = 'A072002', d.Total, null)) 'NO CELEBRADAS' " +
                        "from estadistica.tbldatadicionales d " +
                        "inner " +
                        "join estadistica.tbljuzgados j on j.cveAdscripcion = d.cveJuzgado " +
                        "where " +
                        "d.FechaRep between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                        "and d.CveDatAdi like '%A072%' " +
                        "group by matjuzgado " +
                        "order by field(matjuzgado,'JUZGADOS CIVILES','JUZGADOS USUCAPION','JUZGADOS FAMILIARES','JUZGADOS MERCANTILES'); ";

                MySqlCommand cmd = new MySqlCommand(query, con);
                cmd.CommandTimeout = 1800;
                MySqlDataReader r = cmd.ExecuteReader();


                while (r.Read())
                {
                    JuicioMateriaTotales j = new JuicioMateriaTotales();
                    j.TipoJuzgado = r.GetString("matJuzgado");
                    j.Celebradas = r.GetString("CELEBRADAS");
                    j.NoCelebradas = r.GetString("NO CELEBRADAS");

                    re.Jui.Add(j);

                }
                con.Close();

                /*Celebradas*/
                List<JuicioMateriaTotales> celebradas = new List<JuicioMateriaTotales>();
                celebradas = re.Jui;
                int col = 118;

                foreach (JuicioMateriaTotales c in celebradas)
                {
                    worksheet.Cells[col, 3].Value = c.Celebradas;
                    col++;
                }



                /*No celebradas*/
                List<JuicioMateriaTotales> noCelebradas = new List<JuicioMateriaTotales>();
                noCelebradas = re.Jui;
                int col2 = 118;

                foreach (JuicioMateriaTotales c in noCelebradas)
                {
                    worksheet.Cells[col2, 4].Value = c.NoCelebradas;
                    col2++;
                }





                /*PENAL TRADICIONAL*/
                MySqlConnection con1 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con1.Open();

                query = "select " +
                        "sum(ifnull(d.Total, 0)) 'CELEBRADAS' " +
                        "from estadistica.tbldatadijuzpen d " +
                        "inner " +
                        "join estadistica.tbljuzgados j on j.cveAdscripcion = d.cveJuzgado " +
                        "where " +
                        "d.FechaRep between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                        "and d.CveDatAdi in ('A084003'); ";

                MySqlCommand cmd1 = new MySqlCommand(query, con1);
                cmd1.CommandTimeout = 1800;
                MySqlDataReader r1 = cmd1.ExecuteReader();


                while (r1.Read())
                {

                    re.Penal = r1.GetString("CELEBRADAS");


                }
                con.Close();

                worksheet.Cells[122, 3].Value = re.Penal;

                /*PENAL ACUSATORIO*/
                MySqlConnection con2 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["sigejupe"]);
                con2.Open();

                query = "select " +
                        "count(distinct if (a.cveEstatusAudiencia in (2, 3), a.idaudiencia, null)) 'CELEBRADAS', " +
                        "count(distinct if (a.cveEstatusAudiencia = 1, a.idaudiencia, null)) 'NO CELEBRADAS' " +
                        "from htsj_sigejupe.tblcarpetasjudiciales c " +
                        "inner " +
                        "join htsj_sigejupe.tblaudiencias a  on c.idCarpetaJudicial = a.idReferencia " +
                        "where " +
                        "fechaInicial between '" + fecIni.Text + " " + " " + horaIni + " ' and '" + fecFin.Text + " " + " " + horaFin + " ' " +
                        "and c.cveTipoCarpeta in (2, 3, 4) and a.activo = 'S' and c.activo = 'S' " +
                        "and c.cvejuzgado != 11353; ";

                MySqlCommand cmd2 = new MySqlCommand(query, con2);
                cmd2.CommandTimeout = 1800;
                MySqlDataReader r2 = cmd2.ExecuteReader();


                while (r2.Read())
                {
                    re.Celebradas = r2.GetString("CELEBRADAS");
                    re.NoCelebradas = r2.GetString("NO CELEBRADAS");


                }
                con.Close();

                worksheet.Cells[123, 3].Value = re.Celebradas;
                worksheet.Cells[123, 4].Value = re.NoCelebradas;


                excel.Save();
                excel.Dispose();
                Debug.WriteLine("Se termino el proceso 11 ...");
                //mascara.Visible = false;
            }
            catch (Exception i)
            {
                Debug.WriteLine("No se establecio la conexion :" + i);

            }
        }

        public void PrimerReporteCuadro12()
        {
            string fichero = "C:\\Reportes/ReporteUno.xlsx";
            ExcelPackage excel = new ExcelPackage(new FileInfo(fichero));
            ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];

            RepoUno re = new RepoUno();

            try
            {

                MySqlConnection con = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con.Open();

                query = "select " +
                            "case " +
                            "When j.NomJuzgado like '%Civil%' or j.NomJuzgado like '%Menor%' or j.NomJuzgado like '%Mixto%'  or j.NomJuzgado like '%C.M.%'or j.NomJuzgado like '%C. M.%' then 'JUZGADOS CIVILES' " +
                            "When j.NomJuzgado like '%Mercantil%' then 'JUZGADOS MERCANTILES' " +
                            "When j.NomJuzgado like '%Familiar%' or j.nomJuzgado like '%Adopcion%' then 'JUZGADOS FAMILIARES' " +
                            "When j.NomJuzgado like '%Usucapion%' then 'JUZGADOS USUCAPION' " +
                            "Else j.NomJuzgado end  matJuzgado, " +
                            "sum(if (d.CveDatAdi = 'A071001', d.Total, null)) 'RECIBIDOS', " +
                            "sum(if (d.CveDatAdi = 'A071002', d.Total, null)) 'DILIGENCIADOS' " +
                            "from estadistica.tbldatadicionales d " +
                            "inner " +
                            "join estadistica.tbljuzgados j on j.cveAdscripcion = d.cveJuzgado " +
                            "where " +
                            "d.FechaRep between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                            "and d.CveDatAdi in ('A071001', 'A071002') " +
                            "group by matjuzgado " +
                            "order by field(matjuzgado,'JUZGADOS CIVILES','JUZGADOS USUCAPION','JUZGADOS FAMILIARES','JUZGADOS MERCANTILES');  ";

                MySqlCommand cmd = new MySqlCommand(query, con);
                cmd.CommandTimeout = 1800;
                MySqlDataReader r = cmd.ExecuteReader();


                while (r.Read())
                {
                    JuicioMateriaTotales j = new JuicioMateriaTotales();
                    j.TipoJuzgado = r.GetString("matJuzgado");
                    j.Celebradas = r.GetString("RECIBIDOS");
                    j.NoCelebradas = r.GetString("DILIGENCIADOS");

                    re.Jui.Add(j);

                }
                con.Close();

                /*Celebradas*/
                List<JuicioMateriaTotales> celebradas = new List<JuicioMateriaTotales>();
                celebradas = re.Jui;
                int col = 129;

                foreach (JuicioMateriaTotales c in celebradas)
                {
                    worksheet.Cells[col, 3].Value = c.Celebradas;
                    col++;
                }



                /*No celebradas*/
                List<JuicioMateriaTotales> noCelebradas = new List<JuicioMateriaTotales>();
                noCelebradas = re.Jui;
                int col2 = 129;

                foreach (JuicioMateriaTotales c in noCelebradas)
                {
                    worksheet.Cells[col2, 4].Value = c.NoCelebradas;
                    col2++;
                }

                /*PENAL TRADICIONAL*/
                MySqlConnection con1 = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["mysql-local"]);
                con1.Open();

                query = "select " +
                        "sum(if (d.CveDatAdi = 'A081001', ifnull(d.Total, 0), 0)) 'RECIBIDOS', " +
                        "sum(if (d.CveDatAdi = 'A081002', ifnull(d.Total, 0), 0)) 'DILIGENCIADOS' " +
                        "from estadistica.tbldatadijuzpen d " +
                        "inner " +
                        "join estadistica.tbljuzgados j on j.cveAdscripcion = d.cveJuzgado " +
                        "where " +
                        "d.FechaRep between '" + fecIni.Text + "' and '" + fecFin.Text + "' " +
                        "and d.CveDatAdi in ('A081001', 'A081002'); ";

                MySqlCommand cmd1 = new MySqlCommand(query, con1);
                cmd1.CommandTimeout = 1800;
                MySqlDataReader r1 = cmd1.ExecuteReader();


                while (r1.Read())
                {

                    re.Celebradas = r1.GetString("RECIBIDOS");
                    re.NoCelebradas = r1.GetString("DILIGENCIADOS");


                }
                con.Close();

                worksheet.Cells[133, 3].Value = re.Celebradas;
                worksheet.Cells[133, 4].Value = re.NoCelebradas;

                


                excel.Save();
                excel.Dispose();
                Debug.WriteLine("Se termino el proceso 12 ...");
                //mascara.Visible = false;
            }
            catch (Exception i)
            {
                Debug.WriteLine("No se establecio la conexion :" + i);

            }
        }

        public void PrimerReporteCuadro13()
        {
            string fichero = "C:\\Reportes/ReporteUno.xlsx";
            ExcelPackage excel = new ExcelPackage(new FileInfo(fichero));
            ExcelWorksheet worksheet = excel.Workbook.Worksheets[0];

            RepoUno re = new RepoUno();

            try
            {

                MySqlConnection con = new MySqlConnection(System.Configuration.ConfigurationManager.AppSettings["sigejupe"]);
                con.Open();

                query = "select " +
                        "case " +
                        "when d.cvedelito = 128 then 'FEMINICIDIO' " +
                        "else 'TRATA DE PERSONAS' " +
                        "end delito, " +
                        "count(distinct if (c.cvetipocarpeta in (2) ,c.idCarpetaJudicial,null)) 'CONTROL', " +
                        "count(distinct if (c.cvetipocarpeta in (3, 4) ,c.idCarpetaJudicial,null)) 'TRIBUNAL', " +
                        "COUNT(distinct c.idCarpetaJudicial) TOTAL " +
                        "from htsj_sigejupe.tblcarpetasjudiciales c " +
                        "inner join htsj_sigejupe.tblimpofedelcarpetas im on im.idCarpetaJudicial = c.idCarpetaJudicial " +
                        "inner join htsj_sigejupe.tbldelitoscarpetas dc on dc.idDelitoCarpeta = im.idDelitoCarpeta " +
                        "inner join htsj_sigejupe.tbldelitos d on d.cvedelito = dc.cvedelito " +
                        "where c.fechaRadicacion between '" + fecIni.Text + " " + " " + horaIni + " ' and '" + fecFin.Text + " " + " " + horaFin + " ' " +
                        "and c.cveTipoCarpeta in (2, 3, 4) and(dc.cveDelito = 128 or d.cvecatdelito = 92) " +
                        "and c.cveJuzgado != 11353 and c.activo = 'S' and im.activo = 'S' and dc.activo = 'S' " +
                        "group by delito; ";

                MySqlCommand cmd = new MySqlCommand(query, con);
                cmd.CommandTimeout = 1800;
                MySqlDataReader r = cmd.ExecuteReader();


                while (r.Read())
                {
                    JuicioMateriaTotales j = new JuicioMateriaTotales();
                    j.TipoJuzgado = r.GetString("delito");
                    j.Control = r.GetString("CONTROL");
                    j.Tribunal = r.GetString("TRIBUNAL");
                    j.Total = r.GetString("TOTAL");

                    re.Jui.Add(j);

                }
                con.Close();

                /*control*/
                List<JuicioMateriaTotales> control = new List<JuicioMateriaTotales>();
                control = re.Jui;
                int col = 118;

                foreach (JuicioMateriaTotales c in control)
                {
                    worksheet.Cells[col, 7].Value = c.Control;
                    col++;
                }



                /*tribunal*/
                List<JuicioMateriaTotales> tribunal = new List<JuicioMateriaTotales>();
                tribunal = re.Jui;
                int col2 = 118;

                foreach (JuicioMateriaTotales c in tribunal)
                {
                    worksheet.Cells[col2, 8].Value = c.Tribunal;
                    col2++;
                }

                /*total*/
                List<JuicioMateriaTotales> total = new List<JuicioMateriaTotales>();
                total = re.Jui;
                int col3 = 118;

                foreach (JuicioMateriaTotales c in total)
                {
                    worksheet.Cells[col3, 9].Value = c.Total;
                    col3++;
                }

                excel.Save();
                excel.Dispose();
                Debug.WriteLine("Se termino el proceso 13 ...");
                Debug.WriteLine("fin ...");
                //mascara.Visible = false;
            }
            catch (Exception i)
            {
                Debug.WriteLine("No se establecio la conexion :" + i);

            }
        }
    }
}