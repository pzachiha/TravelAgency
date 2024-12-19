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

namespace UserForm
{
    public partial class FieldOrder : Form
    {
        private static string Str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=TravelAgencyDB.mdb;";
        private OleDbConnection connection = new OleDbConnection(Str);
        private string tableName;
        private int orderId;
        private bool isNewRecord;
        public FieldOrder(OleDbConnection connection, string tableName, int orderId = 0)
        {
            InitializeComponent();
            this.connection = connection;
            this.tableName = tableName;
            this.orderId = orderId;
            this.isNewRecord = orderId == 0;
            comboBox1.Items.AddRange(new string[] { "Оплачен", "Неоплачен", "Отменен" });
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList; // Запрещаем редактирование

            if (!isNewRecord)
            {
                LoadOrderData();
            }
        }
        private void LoadOrderData()
        {
            try
            {
                connection.Open();
                string query = $"SELECT * FROM {tableName} WHERE order_id = @orderId";
                OleDbCommand command = new OleDbCommand(query, connection);
                command.Parameters.AddWithValue("@orderId", orderId);
                OleDbDataReader reader = command.ExecuteReader();

                if (reader.Read())
                {
                    textBox1.Text = reader["client_id"].ToString();
                    textBox2.Text = reader["tour_id"].ToString();
                    dateTimePicker1.Value = Convert.ToDateTime(reader["order_date"]);
                    comboBox1.SelectedItem = reader["payment_status"].ToString();
                }

                reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке данных: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                connection.Open();
                if (isNewRecord)
                {
                    InsertOrder();
                }
                else
                {
                    UpdateOrder();
                }
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при сохранении записи: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }
        private void InsertOrder()
        {
            try
            {
                string query = $"INSERT INTO {tableName} (client_id, tour_id, order_date, payment_status) VALUES (@clientId, @tourId, @orderDate, @paymentStatus)";
                OleDbCommand command = new OleDbCommand(query, connection);

                command.Parameters.AddWithValue("@clientId", textBox1.Text);
                command.Parameters.AddWithValue("@tourId", textBox2.Text);
                command.Parameters.AddWithValue("@orderDate", dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                command.Parameters.AddWithValue("@paymentStatus", comboBox1.SelectedItem.ToString());

                command.ExecuteNonQuery();
                MessageBox.Show("Запись успешно добавлена!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при добавлении записи: " + ex.Message);
            }
        }
        private void UpdateOrder()
        {
            try
            {
                string query = $"UPDATE {tableName} SET client_id = @clientId, tour_id = @tourId, order_date = @orderDate, payment_status = @paymentStatus WHERE order_id = @orderId";
                OleDbCommand command = new OleDbCommand(query, connection);

                command.Parameters.AddWithValue("@clientId", textBox1.Text);
                command.Parameters.AddWithValue("@tourId", textBox2.Text);
                command.Parameters.AddWithValue("@orderDate", dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                command.Parameters.AddWithValue("@paymentStatus", comboBox1.SelectedItem.ToString());
                command.Parameters.AddWithValue("@orderId", orderId);

                command.ExecuteNonQuery();
                MessageBox.Show("Запись успешно обновлена!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при обновлении записи: " + ex.Message);
            }
        }
    }
}