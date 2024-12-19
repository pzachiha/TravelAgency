using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace UserForm
{
    public partial class DelTable : Form
    {
        private static string Str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=TravelAgencyDB.mdb;";
        private OleDbConnection connection = new OleDbConnection(Str);
        public DelTable()
        {
            InitializeComponent();
            LoadTableNames();
        }
        private void LoadTableNames()
        {
            try
            {
                connection.Open();
                DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                foreach (DataRow row in schemaTable.Rows)
                {
                    string tableName = row["TABLE_NAME"].ToString();
                    if (!tableName.StartsWith("MSys") && !tableName.StartsWith("~"))
                    {
                        comboBox1.Items.Add(tableName);
                    }
                }
                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке имен таблиц: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }
        private void DeleteRelatedFields(string tableName)
        {
            DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Foreign_Keys, null);
            foreach (DataRow row in schemaTable.Rows)
            {
                string fkTableName = row["FK_TABLE_NAME"].ToString();
                string fkColumnName = row["FK_COLUMN_NAME"].ToString(); string pkTableName = row["PK_TABLE_NAME"].ToString();
                string pkColumnName = row["PK_COLUMN_NAME"].ToString();
                if (pkTableName == tableName)
                {
                    // Проверяем, существует ли поле в таблице
                    if (FieldExists(fkTableName, fkColumnName))
                    {
                        DropAllForeignKeys(tableName); string deleteFieldQuery = $"ALTER TABLE {fkTableName} DROP COLUMN {fkColumnName}";
                        OleDbCommand deleteCommand = new OleDbCommand(deleteFieldQuery, connection); deleteCommand.ExecuteNonQuery();
                    }
                }
            }
        }
        private bool FieldExists(string tableName, string fieldName)
        {
            DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, tableName, fieldName }); return schemaTable.Rows.Count > 0;
        }
        public void DropAllForeignKeys(string tableName)
        {
            DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Foreign_Keys, null); foreach (DataRow row in schemaTable.Rows)
            {
                string fkTableName = row["FK_TABLE_NAME"].ToString();
                string fkColumnName = row["FK_COLUMN_NAME"].ToString();
                string pkTableName = row["PK_TABLE_NAME"].ToString();
                string pkColumnName = row["PK_COLUMN_NAME"].ToString();
                if (pkTableName == tableName || fkTableName == tableName)
                {
                    string constraintName = row["FK_NAME"].ToString();
                    string dropForeignKeyQuery = $"ALTER TABLE {fkTableName} DROP CONSTRAINT {constraintName}";
                    OleDbCommand dropCommand = new OleDbCommand(dropForeignKeyQuery, connection);
                    dropCommand.ExecuteNonQuery();
                }
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Пожалуйста, выберите таблицу для удаления."); return;
            }
            string tableName = comboBox1.SelectedItem.ToString();
            try
            {
                connection.Open();     // Разрываем все связи с выбранной таблицей
                DeleteRelatedFields(tableName);

                // Удаляем таблицу
                string dropTableQuery = $"DROP TABLE {tableName}"; OleDbCommand dropCommand = new OleDbCommand(dropTableQuery, connection);
                dropCommand.ExecuteNonQuery(); MessageBox.Show("Таблица успешно удалена.");
                // Обновляем список таблиц     comboBox1.Items.Clear();
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при удалении таблицы: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }
    }
}
