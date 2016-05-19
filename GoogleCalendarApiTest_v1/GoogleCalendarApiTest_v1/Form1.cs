using GoogleCalendarApiTest_v1.GoogleAPI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Calendar.v3.Data;


namespace GoogleCalendarApiTest_v1
{
    public partial class Form1 : Form
    {
        private GoogleApiConector calendarConect = new GoogleApiConector();
        private Dictionary<string, Event> eventList = new Dictionary<string, Event>();
        private Color eventColor;

        public Form1()
        {
            InitializeComponent();
            conectToCalendar();
            UpdateListBox();
        }
        public void conectToCalendar()
        {
            try
            {
                calendarConect.CreateCredential();

            }
            catch (Exception e)
            {

                MessageBox.Show(e.ToString());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                calendarConect.CreateEvent(String.Format("Urlop: {0}", textBox1.Text), dateTimePicker1.Value, dateTimePicker2.Value);
                MessageBox.Show("Utworzono wydarzenie", "Sukces", MessageBoxButtons.OK, MessageBoxIcon.Information);
                UpdateListBox();

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString());
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    DialogResult deleteWindow = MessageBox.Show(String.Format("Czy na pewno chcesz usunąć wydarzenie : {0}", listBox1.SelectedItem),
            //        "Ostrzeżenie", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            //    if (deleteWindow == DialogResult.Yes)
            //    {
            //        foreach (var item in eventList)
            //        {
            //            if (listBox1.SelectedItem.ToString() == item.Value.Summary)
            //            {
            //                calendarConect.DeleteEvent(item.Key);
            //                MessageBox.Show("Wydarzenie zostało usunięte!", "Sukces");
            //            }
            //        }
            //        UpdateListBox();
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.ToString());
            //}

        }

        private void UpdateListBox()
        {
            try
            {
                eventList = calendarConect.GetEventList();
                foreach (var item in eventList)
                {

                    dataGridView1.Rows.Add(item.Value.ToString());
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString());
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == WindowState)
            {
                Hide();
            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            WindowState = FormWindowState.Normal;
        }

        private void zakończToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void przywróćToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Show();
            WindowState = FormWindowState.Normal;
        }
    }
}
