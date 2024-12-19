using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CefSharp;
using CefSharp.WinForms;


namespace UserForm
{
    public partial class MapForm : Form
    {
        private string hotelName;
        public MapForm(string hotelName)
        {
            InitializeComponent();
            this.hotelName = hotelName;
            InitializeChromium();
            this.WindowState = FormWindowState.Maximized; // Устанавливаем форму на весь экран
        }

        private void InitializeChromium()
        {
            chromiumWebBrowser1.LoadUrl($"https://www.google.com/maps/search/{Uri.EscapeDataString(hotelName)}");
            chromiumWebBrowser1.Dock = DockStyle.Fill;
        }
        private void MapForm_Load(object sender, EventArgs e)
        {
        }
    }
}
