using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
//using System.Data.OracleClient;
using Oracle.DataAccess.Client;

namespace Transferencia
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        DataView importardatos (String nombrearchivo)
        {
            String conexion = String.Format("Provider= Microsoft.ACE.OLEDB.12.0; Data Source=C:/Users/Elena Pineda/Documents/Transfe.xlsx; Extended Properties='Excel 12.0;'", nombrearchivo);
            OleDbConnection conector = new OleDbConnection(conexion);
            conector.Open();

            OleDbCommand consulta = new OleDbCommand("select * from [Hoja2$]", conector);
            OleDbDataAdapter adaptador = new OleDbDataAdapter
            {
                SelectCommand = consulta

            };
            DataSet ds = new DataSet();
            adaptador.Fill(ds);
            conector.Close();
            return ds.Tables[0].DefaultView;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            {
                OracleConnection ora = new OracleConnection("DATA SOURCE=xe; PASSWORD=12345;USER ID=Elena;");

                ora.Open();
                MessageBox.Show("Conectado");
                ora.Close();

            }
        }

        DataView importardatoss(String nombrearchivo)
        {
            String conexion = String.Format("Provider= Microsoft.ACE.OLEDB.12.0; Data Source=C:/Users/Elena Pineda/Documents/Prueba.xlsx; Extended Properties='Excel 12.0;'", nombrearchivo);
            OleDbConnection conector = new OleDbConnection(conexion);
            conector.Open();

            OleDbCommand consulta = new OleDbCommand("select * from [Hoja1$]", conector);
            OleDbDataAdapter adaptador = new OleDbDataAdapter
            {
                SelectCommand = consulta

            };
            DataSet ds = new DataSet();
            adaptador.Fill(ds);


            dataGridView1.DataSource = ds.Tables[0];

            OracleConnection ora = new OracleConnection();
            ora.ConnectionString = ("DATA SOURCE=xe; PASSWORD=12345;USER ID=Elena;");
            ora.Open();

            OracleBulkCopy importar = default(OracleBulkCopy);
            importar = new OracleBulkCopy(ora);
            importar.DestinationTableName = "AREA1";
            importar.WriteToServer(ds.Tables[0]);
            ora.Close();

            conector.Close();
            return ds.Tables[0].DefaultView;
        }
        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void BtnImportarDatos_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel | *.xls;*.xlsx; ",
                Title = "Seleccionar Archivo "
            };

            if (openFileDialog.ShowDialog()==DialogResult.OK)

            {
                dataGridView1.DataSource = importardatoss(openFileDialog.FileName);
            }
}

        private void Button2_Click(object sender, EventArgs e)
        {

        }
    }
}