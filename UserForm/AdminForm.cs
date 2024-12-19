using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UserForm
{
    public partial class AdminForm : Form
    {
        private static string Str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=TravelAgencyDB.mdb;";
        private OleDbConnection connection = new OleDbConnection(Str);
        private Dictionary<string, string> idColumns = new Dictionary<string, string>
        {
            { "Orders", "order_id" },
            { "Tours", "tour_id" },
            {"Clients", "client_id" },
            {"Services", "tour_id" },
            {"Users", "login" },
            {"Food", "type_food" },
            {"Hotels", "tour_id" }
        };
        public AdminForm()
        {
            InitializeComponent();
            LoadTableNames();
        }
        private void LoadTableNames()
        {
            comboBox1.Items.Clear();

            try
            {
                connection.Open();
                DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                foreach (DataRow row in schemaTable.Rows)
                {
                    comboBox1.Items.Add(row["TABLE_NAME"].ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedTable = comboBox1.SelectedItem.ToString();
            LoadData(selectedTable);
        }
        private void LoadData(string tableName)
        {
            OleDbDataAdapter adapter = null;
            DataTable tableData = new DataTable();
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=TravelAgencyDB.mdb;";

            try
            {
                string query = $"SELECT * FROM {tableName}";
                adapter = new OleDbDataAdapter(query, connectionString);

                // Fill DataTable with data
                adapter.Fill(tableData);

                // Set ListView to Details view
                listView1.View = View.Details;
                listView1.FullRowSelect = true;
                listView1.GridLines = true;

                listView1.Columns.Clear();
                listView1.Items.Clear();

                foreach (DataColumn column in tableData.Columns)
                {
                    listView1.Columns.Add(column.ColumnName);
                }

                listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);

                foreach (DataRow row in tableData.Rows)
                {
                    ListViewItem item = new ListViewItem();
                    for (int i = 0; i < row.ItemArray.Length; i++)
                    {
                        string value = row[i] == DBNull.Value ? string.Empty : row[i].ToString();
                        if (i == 0)
                        {
                            item.Text = value;
                        }
                        else
                        {
                            item.SubItems.Add(value);
                        }
                    }
                    listView1.Items.Add(item);
                }
            }
            catch (OleDbException ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (adapter != null)
                {
                    adapter.Dispose();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string selectedTable = comboBox1.SelectedItem.ToString();
            if (selectedTable == "Orders")
            {
                FieldOrder orderForm = new FieldOrder(connection, comboBox1.SelectedItem.ToString());
                if (orderForm.ShowDialog() == DialogResult.OK)
                {
                    LoadData(comboBox1.SelectedItem.ToString());
                }
            }
            else if (selectedTable == "Tours")
            {
                FieldTour tourForm = new FieldTour(connection, comboBox1.SelectedItem.ToString());
                if (tourForm.ShowDialog() == DialogResult.OK)
                {
                    LoadData(comboBox1.SelectedItem.ToString());
                }
            }
            else
            {
                MessageBox.Show("Для выбранной таблицы не предусмотрено редактирование");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string selectedTable = comboBox1.SelectedItem.ToString();
            if (selectedTable == "Orders")
            {
                if (listView1.SelectedItems.Count == 0)
                {
                    MessageBox.Show("Выберите запись для редактирования.");
                    return;
                }

                int orderId = int.Parse(listView1.SelectedItems[0].SubItems[0].Text);
                FieldOrder orderForm = new FieldOrder(connection, comboBox1.SelectedItem.ToString(), orderId);
                if (orderForm.ShowDialog() == DialogResult.OK)
                {
                    LoadData(comboBox1.SelectedItem.ToString());
                }
            }
            else if (selectedTable == "Tours")
            {
                if (listView1.SelectedItems.Count == 0)
                {
                    MessageBox.Show("Выберите запись для редактирования.");
                    return;
                }

                int tourId = int.Parse(listView1.SelectedItems[0].SubItems[0].Text);
                FieldTour tourForm = new FieldTour(connection, comboBox1.SelectedItem.ToString(), tourId);
                if (tourForm.ShowDialog() == DialogResult.OK)
                {
                    LoadData(comboBox1.SelectedItem.ToString());
                }
            }
            else
            {
                MessageBox.Show("Для выбранной таблицы не предусмотрено редактирование");
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0)
            {
                MessageBox.Show("Выберите запись для удаления.");
                return;
            }

            string selectedTable = comboBox1.SelectedItem.ToString();
            int id = int.Parse(listView1.SelectedItems[0].SubItems[0].Text);

            DeleteRecord(selectedTable, id);
        }
        private void DeleteRecord(string tableName, int id)
        {
            if (!idColumns.ContainsKey(tableName))
            {
                MessageBox.Show("Для выбранной таблицы не предусмотрено удаление");
                return;
            }

            string idColumn = idColumns[tableName];

            try
            {
                connection.Open();
                string query = $"DELETE FROM {tableName} WHERE {idColumn} = @id";
                OleDbCommand command = new OleDbCommand(query, connection);
                command.Parameters.AddWithValue("@id", id);

                command.ExecuteNonQuery();
                MessageBox.Show("Запись успешно удалена!");

                LoadData(tableName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при удалении записи: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DelTable del = new DelTable();
            del.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            AddTable add = new AddTable();
            add.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            LoadTableNames();
        }
    }
}