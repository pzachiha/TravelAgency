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
using System.Security.Cryptography;
using Microsoft.Win32;

namespace UserForm
{
    public partial class Form1 : Form
    {
        private static string Str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=TravelAgencyDB.mdb;";
        private OleDbConnection connection = new OleDbConnection(Str);
        public static string LoggedInPassport { get; private set; }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Дополнительные настройки при загрузке формы
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string login = textBox1.Text;
            string password = textBox2.Text;

            try
            {
                connection.Open();
                string query = "SELECT password, rights, passport FROM Users WHERE login = @name";
                OleDbCommand command = new OleDbCommand(query, connection);
                command.Parameters.AddWithValue("@login", login);
                OleDbDataReader reader = command.ExecuteReader();
                if (reader.Read())
                {

                    string storedHash = reader["password"].ToString();
                    int role = Convert.ToInt32(reader["rights"]);
                    string passport = reader["passport"].ToString(); // Получаем паспорт

                    string enterpassword = VerifyPassword(password);
                    if (enterpassword == storedHash && role == 1)
                    {
                        LoggedInPassport = passport; // Сохраняем паспорт
                        UserForm userForm = new UserForm();
                        userForm.Show();
                        MessageBox.Show("Пароль верный");
                    }
                    if (enterpassword == storedHash && role == 2)
                    {
                        LoggedInPassport = passport; // Сохраняем паспорт
                        AdminForm adminForm = new AdminForm(); adminForm.Show();
                        MessageBox.Show("Вы зашли как администратор");
                    }
                    if (enterpassword != storedHash)
                    {
                        Console.WriteLine(password);
                        MessageBox.Show("Пароль или логин неверный");
                    }
                }
                else
                {
                    MessageBox.Show("Такого пользователя нет");
                }
                reader.Close(); connection.Close();
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
        public string VerifyPassword(string enteredPassword)
        {
            string enteredHash = GetMd5Hash(enteredPassword); return enteredHash;
        }
        public string GetMd5Hash(string s)
        {
            using (var md5 = MD5.Create())
            {
                var hash = md5.ComputeHash(Encoding.UTF8.GetBytes(s));
                return Convert.ToBase64String(hash);
            }
        }

    private void button2_Click(object sender, EventArgs e)
        {
            RegisterForm registerForm = new RegisterForm();
            registerForm.Show();
        }
    }
}
