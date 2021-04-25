using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace TEST_01.mypages
{
    public partial class wfFacturas : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {

               
                Dictionary<int, string> Tipos = cargarTipos();
                foreach (var TipoCliente in Tipos)
                {
                    ListItem it = new ListItem();
                    it.Value = TipoCliente.Key.ToString();
                    it.Text = TipoCliente.Value;
                    lstClientes.Items.Add(it);
                }
                imbExcel.Visible = false;
            }
        }
        Dictionary<int, string> cargarTipos()
        {
            string queryString = "SELECT  [Tipo]      ,[Descripcion] ,[estado] FROM [Test].[dbo].[Catalogo_Tipos_Cliente]";
            string connectionString = ConfigurationManager.ConnectionStrings["myConnectionString"].ConnectionString;
            Dictionary<int, string> temp = new Dictionary<int, string>();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(queryString, connection);
                //command.Parameters.AddWithValue("@tPatSName", "Your-Parm-Value");
                connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        temp.Add(Convert.ToInt32(reader["Tipo"]), reader["Descripcion"].ToString());

                        /*Console.WriteLine(String.Format("{0}, {1}",
                        reader["tPatCulIntPatIDPk"], reader["tPatSFirstname"]));// etc
                        */
                    }
                }
                finally
                {
                    // Always call Close when done reading.
                    reader.Close();
                }
            }
            return temp;
        }

        void listGrid(string qry)
        {
            if (qry == "" ||  string.IsNullOrEmpty(qry))
            {
                qry = "Select f.NIT, c.Nombre, c.Tel, f.Serie_factura, f.Numero_factura, CAST(f.Fecha_Emision as varchar) as Fecha_Emision, CAST(f.Monto as varchar) as Monto from Test.dbo.facturas f left join Test.dbo.clientes c on f.NIT = c.NIT";
            }
            string query = qry;
            System.Data.DataTable dataTable = new System.Data.DataTable();
            string connString = ConfigurationManager.ConnectionStrings["myConnectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand(query, conn);
            conn.Open();

            // create data adapter
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            // this will query your database and return the result to your datatable
            da.Fill(dataTable);
            conn.Close();
            da.Dispose();
            foreach (DataRow dr in dataTable.Rows)
            {
                double t = Convert.ToDouble(dr[6].ToString());
                string g = t.ToString("N2");
                dr[6] = "Q" + g;
                dr[5] = DateTime.Parse((dr[5].ToString())).ToString("dd/MM/yyyy");
                //  string m = dr[6].ToString();

            }
            GridView1.DataSource = dataTable;
            GridView1.DataBind();
        }
        string mesLestra(int mes)
        {
            switch (mes)
            {
                case 2:
                    return "Febrero";
                case 3:
                    return "Marzo";
                case 4:
                    return "Abril";
                case 5:
                    return "Mayo";
                case 6:
                    return "Junio";
                case 7:
                    return "Julio";
                case 8:
                    return "Agosto";
                case 9:
                    return "Septiembre";
                case 10:
                    return "Octubre";
                case 11:
                    return "Noviembre";
                case 12:
                    return "Diciembre";
                default:
                    return "Enero";
            }
        }
        
        void generarExcel(string titulo)
        {
            try
            {
                Response.ClearHeaders();
                Response.Clear();
                Response.AppendHeader("content-disposition", "attachment; filename=Facturas.xls");
                Response.ContentType = "application/excel";

                StringWriter stringWriter = new StringWriter();
                HtmlTextWriter htmlTextWriter = new HtmlTextWriter(stringWriter);

                GridView1.RenderControl(htmlTextWriter);
                string t = stringWriter.ToString();
                string nuevotext = "";
                int c = 0;
                foreach (var item in t.ToCharArray())
                {
                    if (item.ToString().CompareTo("<") == 0)
                    {
                        c++;
                    }
                    if (c == 3)
                    {
                        nuevotext = nuevotext + " " + titulo;
                        c++;
                    }
                    nuevotext = nuevotext + item.ToString();
                    
                }
                Response.Write(nuevotext/* stringWriter.ToString()*/);
                Response.End();

            }
            catch (Exception ex)
            {

            }
           
        }

        public override void VerifyRenderingInServerForm(Control control)
        {
            /* Confirms that an HtmlForm control is rendered for the specified ASP.NET
               server control at run time. */
        }

        protected void BtVerInfo_Click(object sender, EventArgs e)
        {
            imbExcel.Visible = true;
            string monto = tbMontoFactura.Text;
            List<string> tiposcl = new List<string>();
            foreach (ListItem item in lstClientes.Items)
            {
                if (item.Selected)
                {
                    tiposcl.Add(item.Value);
                }
            }
            DateTime t = Convert.ToDateTime(Calendar1.SelectedDate);
           DateTime t2 = Convert.ToDateTime(Calendar2.SelectedDate);

            string query = "Select f.NIT, c.Nombre, c.Tel, f.Serie_factura, f.Numero_factura, CAST(f.Fecha_Emision as varchar) as Fecha_Emision, CAST(f.Monto as varchar) as Monto from Test.dbo.facturas f left join Test.dbo.clientes c on f.NIT = c.NIT";
            string where = " where f.Anulado <> 'A'";
            if (String.IsNullOrEmpty(monto) == false)
            {
                where = where + " and f.Monto >= " + monto;
            }
            if (tiposcl.Count() > 0)
            {
                if (String.IsNullOrEmpty(where) == false)
                {
                    where = where + " and c.Tipo_Cliente in (" + tiposcl.Aggregate((i, j) => i + "," + j)+" )";
                }
                else {
                    where = where + " where c.Tipo_Cliente in " + tiposcl.Aggregate((i, j) => i + "," + j)+" )";
                }
            }
            if (String.IsNullOrEmpty(t.ToString())==false)
            {
                if (String.IsNullOrEmpty(where) == false)
                {
                    where = where + " and f.Fecha_Emision BETWEEN '" + formatoFecha(t.Year.ToString()+"-"+t.Month.ToString() + "-" + t.Day.ToString()) + "' and '"+ formatoFecha(t2.Year.ToString() + "-" + t2.Month.ToString() + "-" + t2.Day.ToString()) +"'" ;
                }
                else
                {
                    where = where + " where f.Fecha_Emision BETWEEN '" + formatoFecha(t.Year.ToString() + "-" + t.Month.ToString() + "-" + t.Day.ToString()) + "' and '" + formatoFecha(t2.Year.ToString() + "-" + t2.Month.ToString() + "-" + t2.Day.ToString()) +"'";
                }
            }
            listGrid(query + where);
        }

        string formatoFecha(string f)
        {
            string r = "";
            string[] g = f.Split('-');
            return g[0] + "-" + g[1] + "-" + g[2] + "  00:00:00.000";
        }
        string formatoFecha2(DateTime T)
        {
            string r = "";
            if (T.Day < 10)
            {
                r = "0";
            }
            r = r + T.Day.ToString() + "/";
            if (T.Month < 10)
            {
                r = r + "0";
            }
            r = r + T.Month.ToString()+"/"+T.Year.ToString();
            
            return r;
        }

        protected void imbExcel_Click(object sender, ImageClickEventArgs e)
        {
            DateTime t = Convert.ToDateTime(Calendar1.SelectedDate);
            DateTime t2 = Convert.ToDateTime(Calendar2.SelectedDate);
            string valor = tbMontoFactura.Text;
            double d = Convert.ToDouble(valor);
            string titulo = "<tr><th colspan = '7' scope = 'col' > REPORTE DE FACTURAS CON EL MONTO MINIMO DE Q."+ d.ToString("N2") + "  DEL "+ formatoFecha2(t)+ " AL " + formatoFecha2(t2) + "</th> </ tr > ";
            generarExcel(titulo);
        }


       

    }
}