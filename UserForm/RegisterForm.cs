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
using System.Text.RegularExpressions;

namespace UserForm
{
    public partial class RegisterForm : Form
    {
        private static string Str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=TravelAgencyDB.mdb;"; 
        OleDbConnection connection = new OleDbConnection(Str);
        public RegisterForm()
        {
            InitializeComponent();
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string login1 = textBox1.Text;
            string password1 = GetMd5Hash(textBox2.Text.Trim());
            int rights1 = 1;
            string last1 = textBox3.Text;
            string first1 = textBox4.Text;
            string middl1 = textBox5.Text;
            string passport1 = textBox6.Text;
            if (!ValidateName(last1) || !ValidateName(first1) || !ValidateName(middl1))
            {
                MessageBox.Show("Фамилия, имя и отчество должны содержать только русские буквы, пробелы и дефисы.");
                return;
            }
            if (!ValidatePassport(passport1))
            {
                MessageBox.Show("Паспорт должен быть в формате '**** ******'.");
                return;
            }

            try
            {
                connection.Open();
                string query = "INSERT INTO [Users] ([login], [password], [rights], [last_name], [first_name], [middle_name], [passport]) VALUES (?, ?, ?, ?, ?, ?, ?)";

                using (OleDbCommand cmd = new OleDbCommand(query, connection))
                {
                    cmd.Parameters.AddWithValue("?", login1);
                    cmd.Parameters.AddWithValue("?", password1);
                    cmd.Parameters.AddWithValue("?", rights1);
                    cmd.Parameters.AddWithValue("?", last1);
                    cmd.Parameters.AddWithValue("?", first1);
                    cmd.Parameters.AddWithValue("?", middl1);
                    cmd.Parameters.AddWithValue("?", passport1);

                    cmd.ExecuteNonQuery();
                }

                MessageBox.Show("Пользователь успешно добавлен.");
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
        public string GetMd5Hash(string s)
        {
            using (var md5 = MD5.Create())
            {
                var hash = md5.ComputeHash(Encoding.UTF8.GetBytes(s)); return Convert.ToBase64String(hash);
            }
        }
        private static bool ValidateName(string name)
        {
            // Регулярное выражение для проверки, что строка содержит только русские буквы и не содержит английские буквы и цифры
            string pattern = @"^[а-яА-ЯёЁ\s-]+$";

            return Regex.IsMatch(name, pattern);
        }
        private bool ValidatePassport(string passport)
        {
            string pattern = @"^\d{4} \d{6}$";
            return Regex.IsMatch(passport, pattern);
        }
    }
}

