using System;
using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportePTGPS
{
    public class ConnectionPT
    {
        #region "ATRIBUTOS DE LA CLASE"
        public string strCnn;     //Cadena de conexion a la base de datos.
        private SqlConnection cnn;  //Objeto para definir la conexión a la base de datos.
        private SqlCommand com;     //Objeto para ejecutar comandos en la base de datos.
        private SqlTransaction tra; //Objeto para ejecutar transacciones en la base de datos.
        #endregion

        #region "PROPIEDADES"
        /// <summary>
        /// Objeto de conexion
        /// </summary>
        public SqlConnection Cnn
        {
            get
            {
                return cnn;
            }
            set
            {
                cnn = value;
            }
        }

        /// <summary>
        /// Objeto para ejecutar instrucciones en la base de datos.
        /// </summary>
        public SqlCommand Com
        {
            get
            {
                return com;
            }
            set
            {
                com = value;
            }
        }

        /// <summary>
        /// Objeto para ejecutar transacciones en la base de datos.
        /// </summary>
        public SqlTransaction Tra
        {
            get
            {
                return tra;
            }
            set
            {
                tra = value;
            }
        }

        /// <summary>
        /// AppSettings shorcut
        /// </summary>
        private NameValueCollection AppSettings { get; set; }
        #endregion

        #region "METODOS CONSTRUCTORES"
        public ConnectionPT(string strCnn)
        {
            this.strCnn = strCnn;
            iniciarObjetos();
        }

        public ConnectionPT()
        {
            this.AppSettings = ConfigurationSettings.AppSettings;
            this.strCnn = this.AppSettings["dbPT"].ToString();
            iniciarObjetos();
        }

        private void iniciarObjetos()
        {
            cnn = new SqlConnection(this.strCnn);
            com = new SqlCommand();
            com.Connection = cnn;
            com.CommandType = CommandType.StoredProcedure;
        }

        #endregion

        #region "METODOS DE LA CLASE"
        /// <summary>
        /// Método para extraer datos de la base de datos.
        /// </summary>
        /// <returns>Retornar el conjunto de datos en un DataTable.</returns>
        public DataTable Seleccionar()
        {
            DataTable dt = new DataTable(); //Definir el de depósito de datos a retornar.
            try
            {
                //Objeto de comandos a ejecutar.
                com.CommandType = CommandType.StoredProcedure;
                com.Connection.Open();
                SqlDataReader dr = com.ExecuteReader(CommandBehavior.CloseConnection);
                dt.Load(dr); //Llenar el depósito de datos con la instrucción a ejecutar.
            }
            catch
            {
                dt = null;
            }
            finally
            {
                if (com.Connection != null)
                    com.Connection.Close();
            }
            return dt;
        }

        /// <summary>
        /// Método para extraer datos de la base de datos.
        /// </summary>
        /// <returns>Retornar el conjunto de datos en un DataTable.</returns>
        public DataTable Seleccionar(string query)
        {
            DataTable dt = new DataTable(); //Definir el de depósito de datos a retornar.
            try
            {
                //Objeto de comandos a ejecutar.
                com.CommandText = query;
                com.CommandType = CommandType.Text;
                com.Connection.Open();
                SqlDataReader dr = com.ExecuteReader(CommandBehavior.CloseConnection);
                dt.Load(dr); //Llenar el depósito de datos con la instrucción a ejecutar.
            }
            catch
            {
                dt = null;
            }
            finally
            {
                if (com.Connection != null)
                    com.Connection.Close();
            }
            return dt;
        }


        /// <summary>
        /// Método para ejecutar instrucciones SQL en la base de datos.
        /// </summary>
        /// <param name="auto">Inidica si si va a insertar en una tabla con autonumérico.</param>
        /// <returns>Retorna el númeor de registros afectados al ejecutar el comando.</returns>
        public int Ejecutar(bool auto)
        {
            int resultado;

            try
            {
                com.Connection.Open();
                //this.Tra = com.Connection.BeginTransaction();
                if (auto)
                    resultado = Convert.ToInt32(com.ExecuteScalar());
                else
                    resultado = com.ExecuteNonQuery();
                //this.Tra.Commit();
            }
            catch (SqlException ex)
            {
                resultado = ex.ErrorCode;
            }
            finally
            {
                if (com.Connection != null) com.Connection.Close();
            }
            return resultado;
        }

        /// <summary>
        /// Método para ejecutar instrucciones SQL en la base de datos.
        /// </summary>
        /// <param name="auto">Inidica si si va a insertar en una tabla con autonumérico.</param>
        /// <returns>Retorna el númeor de registros afectados al ejecutar el comando.</returns>
        public int EjecutarTransaccion()
        {
            try
            {
                return com.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
                return ex.ErrorCode;
            }
        }
        #endregion
    }
}
