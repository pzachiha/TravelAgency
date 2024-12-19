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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace UserForm
{
    public partial class AddTable : Form
    {
        private static string Str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=TravelAgencyDB.mdb;";
        private OleDbConnection connection = new OleDbConnection(Str);
        private List<string> fields = new List<string>();
        private List<string> listname = new List<string>();
        public AddTable()
        {
            InitializeComponent();
            InitializeListView();
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
                        comboBox2.Items.Add(tableName);
                    }
                }
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
        private void button1_Click(object sender, EventArgs e)
        {
            string fieldName = textBox1.Text;
            string type = comboBox1.SelectedItem.ToString();
            bool key = checkBox1.Checked;
            if (string.IsNullOrEmpty(fieldName))
            {
                MessageBox.Show("Пожалуйста, введите название поля.");
                return;
            }

            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Пожалуйста, выберите тип поля.");
                return;

            }
            string fieldDefinition = $"{fieldName} {type}";
            if (key)
            {
                fieldDefinition += " PRIMARY KEY";
            }

            fields.Add(fieldDefinition);
            listname.Add(fieldName);

            // Добавляем поле в ListView 
            listView1.Items.Add(new ListViewItem(new[] { fieldName, type, key ? "Да" : "Нет" }));

            // Очищаем поля ввода 
            textBox1.Text = "";
            comboBox1.SelectedIndex = -1;
            checkBox1.Checked = false;
        }
        private void InitializeListView()
        {
            // Настройка ListView 
            listView1.View = View.Details;
            listView1.Columns.Add("Название поля", 150);
            listView1.Columns.Add("Тип поля", 100);
            listView1.Columns.Add("Ключевое поле", 100);
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (fields.Count == 0)
            {
                MessageBox.Show("Нет полей для создания таблицы."); return;
            }
            string tableName = textBox2.Text; if (string.IsNullOrEmpty(tableName))
            {
                MessageBox.Show("Пожалуйста, введите название таблицы.");
                return;
            }
            string createTableQuery = $"CREATE TABLE {tableName} ({string.Join(", ", fields)})";
            try
            {
                connection.Open();

                OleDbCommand command = new OleDbCommand(createTableQuery, connection); command.ExecuteNonQuery();

                MessageBox.Show("Таблица успешно создана.");
                AddField();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при создании таблицы: " + ex.Message);
            }
            finally
            {
                connection.Close();
                fields.Clear(); listView1.Items.Clear();
            }

        }
        public void AddField()
        {
            string addFieldTable = comboBox2.SelectedItem.ToString(); foreach (string fieldDefinition in fields)
            {
                string tableName = textBox2.Text;
                string[] fieldParts = fieldDefinition.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries); string fieldName = fieldParts[0]; // Получаем имя поля
                string fieldType = fieldParts[1]; // Получаем тип поля
                // Проверяем, является ли поле ключевым
                bool isKeyField = fieldParts.Any(part => part.Equals("PRIMARY", StringComparison.OrdinalIgnoreCase) &&
                fieldParts.Any(p => p.Equals("KEY", StringComparison.OrdinalIgnoreCase)));
                if (isKeyField)
                {
                    try
                    {   // Добавление ключевого поля в выбранную таблицу
                        string addFieldQuery = $"ALTER TABLE {addFieldTable} ADD COLUMN {fieldName} {fieldType}"; OleDbCommand addFieldCommand = new OleDbCommand(addFieldQuery, connection);
                        addFieldCommand.ExecuteNonQuery(); MessageBox.Show($"Поле '{fieldName}' добавлено в таблицу '{addFieldTable}'.");
                        string createForeignKeyQuery = $"ALTER TABLE {addFieldTable} ADD CONSTRAINT FK_{fieldName} FOREIGN KEY ({fieldName}) REFERENCES {tableName}({fieldName})";
                        OleDbCommand createForeignKeyCommand = new OleDbCommand(createForeignKeyQuery, connection); createForeignKeyCommand.ExecuteNonQuery();
                        MessageBox.Show($"Связь между полем '{fieldName}' в таблице '{addFieldTable}' и таблице '{tableName}' создана.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при добавлении поля '{fieldName}': {ex.Message}");
                    }
                }
            }
        }
    }
}
