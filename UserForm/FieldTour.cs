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
    public partial class FieldTour : Form
    {
        private static string Str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=TravelAgencyDB.mdb;";
        private OleDbConnection connection = new OleDbConnection(Str);
        private string tableName;
        private int tourId;
        private bool isNewRecord;
        public FieldTour(OleDbConnection connection, string tableName, int tourId = 0)
        {
            InitializeComponent();
            this.connection = connection;
            this.tableName = tableName;
            this.tourId = tourId;
            this.isNewRecord = tourId == 0;

            // Настройка ComboBox для типа еды
            comboBox1.Items.AddRange(new string[] { "AL", "BB", "FB", "FB+", "HB", "HB+", "RO" });
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList; // Запрещаем редактирование

            if (!isNewRecord)
            {
                LoadTourData();
            }
        }
        private void LoadTourData()
        {
            try
            {
                connection.Open();
                string query = $"SELECT * FROM {tableName} WHERE tour_id = @tourId";
                OleDbCommand command = new OleDbCommand(query, connection);
                command.Parameters.AddWithValue("@tourId", tourId);
                OleDbDataReader reader = command.ExecuteReader();

                if (reader.Read())
                {
                    textBox1.Text = reader["name"].ToString();
                    textBox2.Text = reader["destination"].ToString();
                    comboBox1.SelectedItem = reader["type_food"].ToString();
                    dateTimePicker1.Value = Convert.ToDateTime(reader["start_date"]);
                    dateTimePicker2.Value = Convert.ToDateTime(reader["end_date"]);
                    textBox3.Text = reader["price"].ToString();
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
         private void SaveTour()
        {
            try
            {
                string query;
                OleDbCommand command;

                if (isNewRecord)
                {
                    query = $"INSERT INTO {tableName} (name, destination, type_food, start_date, end_date, price) VALUES (@name, @destination, @typeFood, @startDate, @endDate, @price)";
                    command = new OleDbCommand(query, connection);
                }
                else
                {
                    query = $"UPDATE {tableName} SET name = @name, destination = @destination, type_food = @typeFood, start_date = @startDate, end_date = @endDate, price = @price WHERE tour_id = @tourId";
                    command = new OleDbCommand(query, connection);
                    command.Parameters.AddWithValue("@tourId", tourId);
                }

                command.Parameters.AddWithValue("@name", textBox1.Text);
                command.Parameters.AddWithValue("@destination", textBox2.Text);
                command.Parameters.AddWithValue("@typeFood", comboBox1.SelectedItem.ToString());
                command.Parameters.AddWithValue("@startDate", dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                command.Parameters.AddWithValue("@endDate", dateTimePicker2.Value.ToString("yyyy-MM-dd"));
                command.Parameters.AddWithValue("@price", textBox3.Text);

                command.ExecuteNonQuery();
                MessageBox.Show("Запись успешно сохранена!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при сохранении записи: " + ex.Message);
            }
        }
        private void UpdateTour()
        {
            try
            {
                connection.Open();
                string query = $"UPDATE {tableName} SET name = @name, destination = @destination, type_food = @typeFood, start_date = @startDate, end_date = @endDate, price = @price WHERE tour_id = @tourId";
                OleDbCommand command = new OleDbCommand(query, connection);

                command.Parameters.AddWithValue("@name", textBox1.Text);
                command.Parameters.AddWithValue("@destination", textBox2.Text);
                command.Parameters.AddWithValue("@typeFood", comboBox1.SelectedItem.ToString());
                command.Parameters.AddWithValue("@startDate", dateTimePicker1.Value.ToString("yyyy-MM-dd"));
                command.Parameters.AddWithValue("@endDate", dateTimePicker2.Value.ToString("yyyy-MM-dd"));
                command.Parameters.AddWithValue("@price", textBox3.Text);
                command.Parameters.AddWithValue("@tourId", tourId);

                command.ExecuteNonQuery();
                MessageBox.Show("Запись успешно обновлена!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при обновлении записи: " + ex.Message);
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
                    SaveTour();
                }
                else
                {
                    UpdateTour();
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
    }
}
