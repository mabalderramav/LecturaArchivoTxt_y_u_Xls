using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace GeoPlus.CargaArchivoWeb
{
    public partial class index : System.Web.UI.Page
    {
        #region "Metodos"
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
            return listaExteciones.Exists(ex => ex == extencion);
        }

        #endregion

        #region "Contructor"
        protected void Page_Load(object sender, EventArgs e)
        {
            
            
        }
        #endregion
      
        protected void btnProcesar_Click(object sender, EventArgs e)
        {
            bool isArchivo = false;
            string ruta = Server.MapPath("~/tmp/");
            if (fuPrueba.HasFile)
            {
                string archivoExtension = System.IO.Path.GetExtension(fuPrueba.FileName).ToLower();
                isArchivo = this.isExtencion(archivoExtension);

                if (isArchivo)
                {
                    try
                    {
                        fuPrueba.PostedFile.SaveAs(ruta + fuPrueba.FileName);
                        lblMensaje.Text = "Archivo cargado correctamente";
                        lblMensaje.ForeColor = System.Drawing.Color.Green;
                    }
                    catch (Exception m)
                    {
                        lblMensaje.Text = "ocurrio un problema en el cargado del archivo "
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