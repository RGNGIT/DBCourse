using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.OleDb;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _DBWorks
{

    public partial class DBWorks : Form
    {

        public DBWorks()
        {
            InitializeComponent();
            labelConnectionInfo.Text = String.Empty;
            textBoxDBPath.Text = "Database.mdb";
            textBoxOLEDBP.Text = "Microsoft.Jet.OLEDB.4.0";
        }

        public static string DBConnectionProperties;
        public OleDbConnection CurrentConnection = null;

        // Параметры/Подключение БД

        void Connect()
        {
            try
            {
                if (this.CurrentConnection != null)
                {
                    this.CurrentConnection.Dispose();
                    this.CurrentConnection.Close();
                }
                this.CurrentConnection = new OleDbConnection(DBWorks.DBConnectionProperties);
                this.CurrentConnection.Open();
                labelConnectionInfo.Text = "База данных успешно подключена и готова к работе!";
                tabPage7.Text = $"Кастомный SQL запрос ({textBoxOLEDBP.Text})";
            }
            catch (ArgumentException)
            {
                labelConnectionInfo.Text = "Ошибка инициализации базы данных и/или OleDB провайдера!";
            }
            catch (InvalidOperationException)
            {
                labelConnectionInfo.Text = "Данный провайдер не представлен в Вашей системе!";
            }
            catch (OleDbException)
            {
                labelConnectionInfo.Text = "Данная БД не найдена!";
            }
        }

        void SetSettings(string getProvider, string getPath)
        {
            DBWorks.DBConnectionProperties = $"Provider = {getProvider}; Data Source = {getPath};";
            Connect();
        }

        // Работа с SQL

        string sqlRequest = String.Empty;

        void Output(string dbTable)
        {
            this.sqlRequest = $"SELECT * FROM [{dbTable}]";
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(this.sqlRequest, this.CurrentConnection);
            DataSet DataSet = new DataSet();
            dataAdapter.Fill(DataSet, $"[{dbTable}]");
            dataGridViewMain.DataSource = DataSet.Tables[0].DefaultView;
        }

        void CustomSQL(string sqlRequest)
        {
            this.sqlRequest = sqlRequest;
            OleDbCommand command = new OleDbCommand(this.sqlRequest, this.CurrentConnection);
            command.ExecuteNonQuery();
        }

        void Insert(string dbTable, string dbValues)
        {
            this.sqlRequest = $"INSERT INTO [{dbTable}] VALUES ({dbValues})";
            OleDbCommand command = new OleDbCommand(this.sqlRequest, this.CurrentConnection);
            command.ExecuteNonQuery();
            Output(dbTable);
        }

        void Delete(string dbTable, string dbKeyField, string dbKeyValue, bool isTable)
        {
            if (!isTable)
            {
                this.sqlRequest = $"DELETE FROM [{dbTable}] WHERE {dbKeyField} = {dbKeyValue}";
                OleDbCommand command = new OleDbCommand(this.sqlRequest, this.CurrentConnection);
                command.ExecuteNonQuery();
                Output(dbTable);
            }
            else
            {
                this.sqlRequest = $"DROP TABLE [{dbTable}]";
                OleDbCommand command = new OleDbCommand(this.sqlRequest, this.CurrentConnection);
                command.ExecuteNonQuery();
            }
        }

        void Edit(string dbTable, string dbTab, string dbKeyField, string dbKeyValue, string dbEditValue)
        {
            this.sqlRequest = $"UPDATE [{dbTable}] SET [{dbTab}] = {dbEditValue} WHERE {dbKeyField} = {dbKeyValue}";
            OleDbCommand command = new OleDbCommand(this.sqlRequest, this.CurrentConnection);
            command.ExecuteNonQuery();
            Output(dbTable);
        }

        void AddTableSQLAssembly(string dbTableName, string dbTabName, string varType, bool notNull)
        {

        }

        void AddTable()
        {

        }

        void ClearRequest()
        {

        }

        private void SQLRequestOutputDraw(object sender, PaintEventArgs e)
        {
            using (Font font = new Font("Arial", 10))
            {
                e.Graphics.DrawString(this.sqlRequest, font, Brushes.Black, new Point(2, 2));
            }
        }

        // Работа со схемами данных



        // Банк вызовов

        private void Call_ApplySettings(object sender, EventArgs e) => SetSettings(textBoxOLEDBP.Text, textBoxDBPath.Text);
        private void Call_InsertDBAction(object sender, EventArgs e) => Insert(textBoxAddTable.Text, textBoxAddParam.Text);
        private void Call_DeleteDBAction(object sender, EventArgs e) => Delete(textBoxDelTable.Text, textBoxDelKeyField.Text, textBoxDelKeyValue.Text, checkBoxTableDeletion.Checked);
        private void Call_EditDBAction(object sender, EventArgs e) => Edit(textBoxEditTable.Text, textBoxEditTab.Text, textBoxEditKeyField.Text, textBoxEditKeyValue.Text, textBoxToEdit.Text);
        private void Call_OutputDBAction(object sender, EventArgs e) => Output(textBoxOut.Text);
        private void Call_CustomSQLDBAction(object sender, EventArgs e) => CustomSQL(textBoxSQLRequest.Text);

        private void buttonAddTToRequest_Click(object sender, EventArgs e)
        {

        }

        private void DBWorks_Load(object sender, EventArgs e)
        {

        }
    }

}
