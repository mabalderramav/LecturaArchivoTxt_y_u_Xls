using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;

namespace GeoPlus.CargaArchivoWeb
{
    public partial class index : System.Web.UI.Page
    {
        #region "Metodos ADO.NET"

        /// <summary>
        /// Armando un data table para la carga de los archivos
        /// </summary>
        /// <returns></returns>
        private System.Data.DataTable armandoColumnDataTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("column0");
            dt.Columns.Add("column1");
            dt.Columns.Add("column2");
            dt.Columns.Add("column3");
            dt.Columns.Add("column4");
            dt.Columns.Add("column5");
            dt.Columns.Add("column6");
            dt.Columns.Add("column7");
            dt.Columns.Add("column8");
            dt.Columns.Add("column9");
            dt.Columns.Add("column10");
            dt.Columns.Add("column11");

            return dt;
        }

        /// <summary>
        /// Obtiene el data table de un archivo con extension .csv que se leyo previamente
        /// </summary>
        /// <param name="lista">Lista de datos del archivo csv</param>
        /// <returns></returns>
        private System.Data.DataTable getDataTableCsv(List<string> lista)
        {
            System.Data.DataTable dt = armandoColumnDataTable();
            int cont = 0;
            foreach (string item in lista)
            {
                if (cont > 8)
                {
                    string[] ListItems = item.Split(';');
                    dt.Rows.Add(ListItems);
                }
                cont++;
            }
            return dt;
        }

        private string getCadenaConexion(string extencion)
        {
            string ruta = Server.MapPath("~/tmp/archivotmp.xls");
            string sConnectionString = "";
            if (extencion.Equals(".xls"))
            {
                sConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ruta + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
            }
            else
            {
                sConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ruta + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            }
            return sConnectionString;
        }
        /// <summary>
        /// Obtiene el DataSet de un archivo excel con extension .xls y .xlsx
        /// </summary>
        /// <param name="extension"></param>
        /// <returns></returns>
        private DataSet leerArchivoExcel(string extension)
        {

            // Crear un nuevo OleDbDataAdapter que se usa para generar un DataSet // basado en la instrucción SQL SELECT anterior. 
            OleDbDataAdapter da = new OleDbDataAdapter();

            // Crear el nuevo DataSet para que aloje la información de la hoja de cálculo. 
            DataSet ds = new DataSet();

            // Crear el objeto de conexión utilizando la cadena de conexión anterior.
            
            using (OleDbConnection conexion = new OleDbConnection(getCadenaConexion(extension)))
            {
                // Abrir la conexión con la base de datos.
                conexion.Open();

                // El código utiliza un comando SQL SELECT para mostrar los datos de la hoja de cálculo.
                // Crear un nuevo OleDbCommand para devolver los datos de la hoja de cálculo. 
                OleDbCommand com = new OleDbCommand("SELECT * FROM [rPatrolRecord$]", conexion);

                // Pasar el comando Select al adaptador. 
                da.SelectCommand = com;

                // Llenar el DataSet con la información de la hoja de cálculo. 
                da.Fill(ds, "XLData");    
            }
            return ds;
        }

        #endregion

        #region "Metodos"

        /// <summary>
        /// Leé el archivo csv para despues procesarlo
        /// </summary>
        /// <returns></returns>
        private List<string> leerArchivoCsv()
        {
            int counter = 0;
            string line;
            List<string> listaLine = new List<string>();
            string ruta = Server.MapPath("~/tmp/archivotmp.csv");

            using (System.IO.StreamReader archivo = new System.IO.StreamReader (ruta))
            {
                while ((line = archivo.ReadLine()) != null)
                {
                    listaLine.Add(line);
                    counter++;
                }
            }

            return listaLine;
        }

        /// <summary>
        /// Procesa el archivo con extensión .csv
        /// </summary>
        private void procesarArchivoCsv()
        {
            List<string> lista = leerArchivoCsv();
            System.Data.DataTable dt = getDataTableCsv(lista);
            gv.DataSource = dt;
            gv.DataBind();
        }


        /// <summary>
        /// Procesa el archivo Excel con extension .xls y .xlsx
        /// </summary>
        /// <param name="extencion">Extension del archivo excel. (Permite .xls y .xlsx)</param>
        private void procesarArchivoXls(string extencion)
        {
            

            // Crear el nuevo DataSet para que aloje la información de la hoja de cálculo. 
            DataSet ds = leerArchivoExcel(extencion);

            // Enlazar los datos al control DataGrid. 
            gv.DataSource = ds.Tables[0].DefaultView; 
            gv.DataBind();
        }

        /// <summary>
        /// Verifica si la extension es correcta del archivo.
        /// Las extenciones válidas son txt, xls, xlsx, csv
        /// </summary>
        /// <param name="extencion">Exteción del archivo a evaluar</param>
        /// <returns></returns>
        private bool isExtencion(string extencion)
        {
            List<string> listaExteciones = new List<string> 
                {
                    ".txt",
                    ".xls",
                    ".xlsx",
                    ".csv"
                };
            return listaExteciones.Exists(ex => ex == extencion.ToLower());
        }

        #endregion

        #region "Contructor"
        protected void Page_Load(object sender, EventArgs e)
        {
            
            
        }
        #endregion

        /// <summary>
        /// Procesa los archivos suvidos por el uploadFile :)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void btnProcesar_Click(object sender, EventArgs e)
        {
            bool isArchivo = false;
            string ruta = Server.MapPath("~/tmp/archivotmp");

            //Verifica si el componente uploadFile contiene un archivo
            if (fuPrueba.HasFile)
            {
                
                string archivoExtension = System.IO.Path.GetExtension(fuPrueba.FileName).ToLower();
                isArchivo = this.isExtencion(archivoExtension);//Verifica si las exteciones son correctas

                if (isArchivo)
                {
                    try
                    {
                        if (System.IO.File.Exists(ruta + archivoExtension))
                        {
                            System.IO.File.Delete(ruta + archivoExtension);//Si existe el archivo temporal se lo elimina para crear el nuevo archivo
                        }

                        fuPrueba.PostedFile.SaveAs(ruta + archivoExtension);
                        lblMensaje.Text = "Archivo cargado correctamente";
                        lblMensaje.ForeColor = System.Drawing.Color.Green;

                        switch (archivoExtension)
                        {
                            case ".csv":
                                procesarArchivoCsv();
                                break;
                            case ".xls":
                                procesarArchivoXls(archivoExtension.ToLower());
                                break;
                            case ".xlsx":
                                procesarArchivoXls(archivoExtension.ToLower());
                                break;
                            default:
                                //Not tiene implementación :(
                                break;
                        }
                    }
                    catch (Exception m)
                    {
                        lblMensaje.Text = "ocurrio un problema en el cargado del archivo"
                            + Environment.NewLine + m.Message;
                        lblMensaje.ForeColor = System.Drawing.Color.Red;
                    }
                }
                else
                {
                    lblMensaje.Text = "No se puede cargar el archivo ya que no es valido";
                    lblMensaje.ForeColor = System.Drawing.Color.Red;
                }
            }
        }
    }
}