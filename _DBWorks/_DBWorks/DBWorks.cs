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
            textBoxDBPath.Text = "DB_Course.mdb";
            textBoxOLEDBP.Text = "Microsoft.Jet.OLEDB.4.0";
        }

        public static string DBConnectionProperties;
        public OleDbConnection CurrentConnection = null;
        List<string> Tables = new List<string>() { "ОБЛОРГ", "СПОБЪЕКТ", "МЕРОПРИЯТИЕ" };

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
            tabPage1.Text = $"Работа с БД (Таблица в базе данных: {dbTable})";
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

        void Insert()
        {
            switch (tabControlAdd.SelectedIndex)
            {
                case 0:
                    this.sqlRequest = $"INSERT INTO [ОБЛОРГ] VALUES ({int.Parse(textBoxOrgCodeAdd.Text)}, '{textBoxOrgNameAdd.Text}', '{textBoxOrgShNameAdd.Text}', '{textBoxOrgAddrAdd.Text}', '{textBoxOrgTelAdd.Text}', '{textBoxOrgEmailAdd.Text}')";
                    break;
                case 1:
                    this.sqlRequest = $"INSERT INTO [СПОБЪЕКТ] VALUES ('{textBoxSpPlaceAdd.Text}', {int.Parse(textBoxSpNumAdd.Text)}, '{textBoxSpNameAdd.Text}', '{textBoxSpShNameAdd.Text}', '{textBoxSpTypeAdd.Text}', {int.Parse(textBoxSpSquareAdd.Text)}, {int.Parse(textBoxSpCapacityAdd.Text)}, '{textBoxSpOrgAdd.Text}', '{textBoxSpBalDateAdd.Text}', '{textBoxSpEventAdd.Text}')";
                    break;
                case 2:
                    this.sqlRequest = $"INSERT INTO [МЕРОПРИЯТИЕ] VALUES ('{textBoxEventTypeAdd.Text}', '{textBoxEventNameAdd.Text}', '{textBoxEventDateAdd.Text}', {int.Parse(textBoxEventsVisitorsAdd.Text)})";
                    break;
            }
            OleDbCommand command = new OleDbCommand(this.sqlRequest, this.CurrentConnection);
            command.ExecuteNonQuery();
            Output(this.Tables[tabControlAdd.SelectedIndex]);
        }

        void Delete(int dbTable, string dbKeyField, string dbKeyValue)
        {
            try
            {
                this.sqlRequest = $"DELETE FROM [{this.Tables[dbTable]}] WHERE {dbKeyField} = '{dbKeyValue}'";
                OleDbCommand command = new OleDbCommand(this.sqlRequest, this.CurrentConnection);
                command.ExecuteNonQuery();
                Output(this.Tables[dbTable]);
            }
            catch (OleDbException)
            {
                this.sqlRequest = $"DELETE FROM [{this.Tables[dbTable]}] WHERE {dbKeyField} = {dbKeyValue}";
                OleDbCommand command = new OleDbCommand(this.sqlRequest, this.CurrentConnection);
                command.ExecuteNonQuery();
                Output(this.Tables[dbTable]);
            }
        }

        void Edit(int dbTable, string dbTab, string dbKeyField, string dbKeyValue, string dbEditValue)
        {
            this.sqlRequest = $"UPDATE [{this.Tables[dbTable]}] SET [{dbTab}] = {dbEditValue} WHERE {dbKeyField} = {dbKeyValue}";
            OleDbCommand command = new OleDbCommand(this.sqlRequest, this.CurrentConnection);
            command.ExecuteNonQuery();
            Output(this.Tables[dbTable]);
        }

        // Банк вызовов

        private void Call_ApplySettings(object sender, EventArgs e) => SetSettings(textBoxOLEDBP.Text, textBoxDBPath.Text);
        private void Call_InsertDBAction(object sender, EventArgs e) => Insert();
        private void Call_DeleteDBAction(object sender, EventArgs e) => Delete(comboBoxDelList.SelectedIndex, textBoxDelKeyField.Text, textBoxDelKeyValue.Text);
        private void Call_EditDBAction(object sender, EventArgs e) => Edit(comboBoxRedList.SelectedIndex, textBoxEditTab.Text, textBoxEditKeyField.Text, textBoxEditKeyValue.Text, textBoxToEdit.Text);
        private void Call_OutputDBAction(object sender, EventArgs e) => Output(this.Tables[comboBoxOutTable.SelectedIndex]);
        private void Call_CustomSQLDBAction(object sender, EventArgs e) => CustomSQL(textBoxSQLRequest.Text);
        private void Call_DelListUpdate(object sender, EventArgs e) => Output(this.Tables[comboBoxDelList.SelectedIndex]);
        private void Call_RedListUpdate(object sender, EventArgs e) => Output(this.Tables[comboBoxRedList.SelectedIndex]);
        private void Call_AddTab(object sender, EventArgs e) => Output(this.Tables[tabControlAdd.SelectedIndex]);

    }

}
