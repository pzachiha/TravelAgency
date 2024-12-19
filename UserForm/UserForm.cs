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
using System.Collections;
using System.Text.RegularExpressions;
using System.Net.Mail;
using System.Net;
using System.IO;
using OfficeOpenXml;

namespace UserForm
{ 
    public partial class UserForm : Form
    {
        private static string Str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=TravelAgencyDB.mdb;";
        private OleDbConnection connection = new OleDbConnection(Str);

        public UserForm()
        {
            InitializeComponent();
            textBox5.KeyPress += textBox5_KeyPress;
        }

        private void UserForm_Load(object sender, EventArgs e)
        {
            // Заполнение таблицы
            listView1.Items.Clear();
            listView1.Columns.Clear();
            listView1.Columns.Add("ID");
            listView1.Columns.Add("Наименование");
            listView1.Columns.Add("Страна");
            listView1.Columns.Add("Тип питания");
            listView1.Columns.Add("Дата начала");
            listView1.Columns.Add("Дата конца");
            listView1.Columns.Add("Цена");
            listView1.View = View.Details;
            PopulateListView();

            // Выбор страны
            OleDbDataAdapter adapter2 = new OleDbDataAdapter("SELECT destination FROM Tours", connection);
            DataTable dataTable1 = new DataTable();
            adapter2.Fill(dataTable1);
            comboBox1.DataSource = dataTable1;
            comboBox1.DisplayMember = "destination";
        }

        private void Reader(string query)
        {
            listView1.Items.Clear();

            try
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand(query, connection);
                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    ListViewItem item = new ListViewItem(reader["tour_id"].ToString());
                    item.SubItems.Add(reader["name"].ToString());
                    item.SubItems.Add(reader["destination"].ToString());
                    item.SubItems.Add(reader["type_food"].ToString());
                    item.SubItems.Add(reader["start_date"].ToString());
                    item.SubItems.Add(reader["end_date"].ToString());
                    item.SubItems.Add(reader["price"].ToString());
                    listView1.Items.Add(item);
                }

                reader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }

        private void PopulateListView()
        {
            string query = "SELECT * FROM Tours";
            Reader(query);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            string destination = comboBox1.SelectedItem != null ? ((DataRowView)comboBox1.SelectedItem)["destination"].ToString() : string.Empty;
            string priceFrom = textBox1.Text;
            string priceTo = textBox6.Text;
            // Страна и цена
            if (comboBox1.SelectedIndex >= 0 && !string.IsNullOrEmpty(priceFrom) && !string.IsNullOrEmpty(priceTo))
            {
                string query = $"SELECT * FROM Tours WHERE (destination = \"{destination}\" AND price >= {priceFrom} AND price <= {priceTo});";
                Reader(query);
            }
            // Страна
            else if (comboBox1.SelectedIndex >= 0)
            {
                string query = $"SELECT * FROM Tours WHERE destination = \"{destination}\";";
                Reader(query);
            }
            // Цена
            else if (!string.IsNullOrEmpty(priceFrom) && !string.IsNullOrEmpty(priceTo))
            {
                string query = $"SELECT * FROM Tours WHERE (price >= {priceFrom} AND price <= {priceTo});";
                Reader(query);
            }
            else
            {
                PopulateListView();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PopulateListView();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!IsValidPhoneNumber(textBox5.Text))
            {
                MessageBox.Show("Номер телефона должен состоять из 11 цифр.");
                return;
            }

            connection.Open();
            // Проверяем, выбрана ли строка в ListView
            if (listView1.SelectedItems.Count == 0)
            {
                MessageBox.Show("Пожалуйста, выберите тур из списка."); return;
            }
            // Получаем данные выбранной строки из ListView
            ListViewItem selectedItem = listView1.SelectedItems[0];
            int tourId = int.Parse(selectedItem.SubItems[0].Text);
            DateTime orderDate = DateTime.Now;
            string payment_status = "Неоплачен";

            // Получаем паспорт из LoginForm
            string passport = Form1.LoggedInPassport;
            string phoneNumber = textBox5.Text;
            string address = textBox4.Text;

            int client_Id;

            // Проверка наличия клиента по паспорту
            string checkClientQuery = $"SELECT client_id FROM Clients WHERE passport = '{passport}'";
            OleDbCommand checkClientCommand = new OleDbCommand(checkClientQuery, connection);
            object result = checkClientCommand.ExecuteScalar();

            if (result != null)
            {
                // Клиент найден, используем его client_id
                client_Id = Convert.ToInt32(result);

                // Обновляем информацию о телефоне и адресе
                string updateClientQuery = $"UPDATE Clients SET phone = '{phoneNumber}', address = '{address}' WHERE passport = '{passport}'";
                OleDbCommand updateClientCommand = new OleDbCommand(updateClientQuery, connection);
                updateClientCommand.ExecuteNonQuery();
            }
            else
            {
                // Клиент не найден, проверяем наличие пользователя по паспорту
                string checkUserQuery = $"SELECT last_name, first_name, middle_name FROM Users WHERE passport = '{passport}'";
                OleDbCommand checkUserCommand = new OleDbCommand(checkUserQuery, connection);
                OleDbDataReader reader = checkUserCommand.ExecuteReader();

                if (reader.Read())
                {
                    string lastName = reader["last_name"].ToString();
                    string firstName = reader["first_name"].ToString();
                    string middleName = reader["middle_name"].ToString();

                    // Создаем нового клиента с данными из таблицы Users
                    string insertClientQuery = $"INSERT INTO Clients (last_name, first_name, middle_name, passport, phone, address) VALUES ('{lastName}', '{firstName}', '{middleName}', '{passport}', '{phoneNumber}', '{address}')";
                    OleDbCommand insertClientCommand = new OleDbCommand(insertClientQuery, connection);
                    insertClientCommand.ExecuteNonQuery();

                    // Получаем client_id нового клиента
                    string getClientIdQuery = "SELECT @@IDENTITY";
                    OleDbCommand getClientIdCommand = new OleDbCommand(getClientIdQuery, connection);
                    client_Id = Convert.ToInt32(getClientIdCommand.ExecuteScalar());
                }
                else
                {
                    MessageBox.Show("Пользователь с таким паспортом не найден.");
                    connection.Close();
                    return;
                }
            }

            // Создаем заказ
            string insertOrderQuery = $"INSERT INTO Orders (client_id, tour_id, order_date, payment_status) VALUES ({client_Id}, {tourId}, #{orderDate:yyyy-MM-dd}#, '{payment_status}')";
            OleDbCommand insertOrderCommand = new OleDbCommand(insertOrderQuery, connection);
            insertOrderCommand.ExecuteNonQuery();
            string email = textBox2.Text;
            try
            {
                SendEmailAsync(email);
                MessageBox.Show("Заказ успешно оформлен! Письмо отправлено на ваш email.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при отправке письма: " + ex.Message);
            }
            
            connection.Close();
        }
        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Разрешаем только цифры и ограничиваем количество символов до 11
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
            else if (textBox5.Text.Length >= 11 && e.KeyChar != (char)Keys.Back)
            {
                e.Handled = true;
            }
        }
        private bool IsValidPhoneNumber(string phoneNumber)
        {
            // Проверка, что номер телефона состоит из 11 цифр
            return phoneNumber.Length == 11 && phoneNumber.All(char.IsDigit);
        }
        private bool ValidatePassport(string passport)
        {
            string pattern = @"^\d{4} \d{6}$";
            return Regex.IsMatch(passport, pattern);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0)
            {
                MessageBox.Show("Пожалуйста, выберите тур из списка.");
                return;
            }

            // Получаем данные выбранной строки из ListView
            ListViewItem selectedItem = listView1.SelectedItems[0];
            int tourId = int.Parse(selectedItem.SubItems[0].Text); // Получаем tour_id

            // Получаем название отеля по tour_id
            string hotelName = GetHotelNameByTourId(tourId);

            if (string.IsNullOrEmpty(hotelName))
            {
                MessageBox.Show("Не удалось найти отель для выбранного тура.");
                return;
            }

            // Открываем форму с картой и передаем название отеля
            MapForm mapForm = new MapForm(hotelName);
            mapForm.Show();
        }
        private string GetHotelNameByTourId(int tourId)
        {
            string hotelName = string.Empty;

            try
            {
                connection.Open();
                string query = "SELECT hotel FROM Hotels WHERE tour_id = @tourId";
                OleDbCommand command = new OleDbCommand(query, connection);
                command.Parameters.AddWithValue("@tourId", tourId);
                object result = command.ExecuteScalar();

                if (result != null)
                {
                    hotelName = result.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }

            return hotelName;
        }
        public async Task SendEmailAsync(string toEmail)
        {
            try
            {
                string fromEmail = "anastasuasuhanova05@gmail.com"; 
                string password = "zazrpxbrsiozhgwz";
                MailAddress from = new MailAddress(fromEmail, "BR_Manager"); 
                MailAddress to = new MailAddress(toEmail);
                MailMessage m = new MailMessage(from, to); 
                m.Subject = "Заказ в турфирме";
                m.Body = print(); 
                SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
                smtp.Credentials = new NetworkCredential(fromEmail, password); 
                smtp.EnableSsl = true;
                await smtp.SendMailAsync(m); 
                MessageBox.Show("Письмо отправлено");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при отправке письма: " + ex.Message);
            }
        }
        public string print()
        {
            // Получаем данные выбранного тура
            ListViewItem selectedItem = listView1.SelectedItems[0];
            string name = selectedItem.SubItems[1].Text;
            string destination = selectedItem.SubItems[2].Text; 
            string typeFood = selectedItem.SubItems[3].Text;
            string startDate = selectedItem.SubItems[4].Text; 
            string endDate = selectedItem.SubItems[5].Text;
            string price = selectedItem.SubItems[6].Text;
            // Формируем тело сообщения
            string body = $@"
            Ваш заказ успешно оформлен.    
            Информация о туре:

            Название тура: { name}
            Место назначения: { destination}
            Тип питания: { typeFood}
            Дата начала: { startDate}
            Дата окончания: { endDate}
            Цена: { price}
            ";    
            return body;
        }

        private void GenerateReport(string passport)
        {
            try
            {
                connection.Open();

                // Получаем данные о клиенте
                string clientQuery = $"SELECT * FROM Clients WHERE passport = '{passport}'";
                OleDbCommand clientCommand = new OleDbCommand(clientQuery, connection);
                OleDbDataReader clientReader = clientCommand.ExecuteReader();

                if (!clientReader.Read())
                {
                    MessageBox.Show("Клиент с таким паспортом не найден.");
                    return;
                }

                string clientId = clientReader["client_id"].ToString();
                string lastName = clientReader["last_name"].ToString();
                string firstName = clientReader["first_name"].ToString();
                string middleName = clientReader["middle_name"].ToString();
                string phone = clientReader["phone"].ToString();
                string address = clientReader["address"].ToString();

                clientReader.Close();

                // Получаем данные о заказах клиента
                string ordersQuery = $"SELECT Orders.*, Tours.name AS tour_name, Tours.destination AS tour_destination FROM Orders INNER JOIN Tours ON Orders.tour_id = Tours.tour_id WHERE Orders.client_id = {clientId}";
                OleDbCommand ordersCommand = new OleDbCommand(ordersQuery, connection);
                OleDbDataReader ordersReader = ordersCommand.ExecuteReader();

                // Создаем Excel файл
                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Отчет");

                    // Заголовки столбцов
                    worksheet.Cells[1, 1].Value = "Фамилия";
                    worksheet.Cells[1, 2].Value = "Имя";
                    worksheet.Cells[1, 3].Value = "Отчество";
                    worksheet.Cells[1, 4].Value = "Телефон";
                    worksheet.Cells[1, 5].Value = "Адрес";
                    worksheet.Cells[1, 6].Value = "Наименование тура";
                    worksheet.Cells[1, 7].Value = "Страна";
                    worksheet.Cells[1, 8].Value = "Дата заказа";
                    worksheet.Cells[1, 9].Value = "Статус оплаты";

                    // Данные клиента
                    worksheet.Cells[2, 1].Value = lastName;
                    worksheet.Cells[2, 2].Value = firstName;
                    worksheet.Cells[2, 3].Value = middleName;
                    worksheet.Cells[2, 4].Value = phone;
                    worksheet.Cells[2, 5].Value = address;

                    int row = 3;
                    while (ordersReader.Read())
                    {
                        worksheet.Cells[row, 6].Value = ordersReader["tour_name"].ToString();
                        worksheet.Cells[row, 7].Value = ordersReader["tour_destination"].ToString();
                        worksheet.Cells[row, 8].Value = ordersReader["order_date"].ToString();
                        worksheet.Cells[row, 9].Value = ordersReader["payment_status"].ToString();
                        row++;
                    }

                    ordersReader.Close();

                    // Сохраняем файл
                    string filePath = Path.Combine(Application.StartupPath, "Report.xlsx");
                    FileInfo file = new FileInfo(filePath);
                    package.SaveAs(file);

                    MessageBox.Show("Отчет успешно создан и сохранен в файл: " + filePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            // Получаем паспорт из LoginForm
            string passport = Form1.LoggedInPassport;

            if (string.IsNullOrEmpty(passport))
            {
                MessageBox.Show("Паспорт не найден.");
                return;
            }
            GenerateReport(passport);
        }
    }
}
