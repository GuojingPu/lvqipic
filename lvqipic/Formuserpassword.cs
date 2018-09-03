using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace lvqipic
{
    public partial class Formuserpassword : Form
    {
        Form1 form1;
        public DataTable g_datausermeg = null;

        string REGISTERUSER = "Excel\\注册用户.xls";


        private bool mouseflag = false;//鼠标是否按下
        private Point FormLocation;     //form的location
        private Point mouseOffset;      //鼠标的按下位置

        public Formuserpassword(Form1 f1)
        {
            form1 = f1;

            InitializeComponent();
            g_datausermeg = ClassExcel.ExcelToDataTable(Application.StartupPath + "\\"+REGISTERUSER, null, true);
            if (g_datausermeg == null)
            {
                MessageBox.Show("用户注册信息表读取失败！" + Application.StartupPath + "\\"+ REGISTERUSER);
            }

          

        }
        private string userfound()
        {
            bool isfound = false;
            bool ispasscrect = false;
            int i = 0;

            if (g_datausermeg == null)
            {
                return null;
            }
            for (i = 0; i < g_datausermeg.Rows.Count; i++)
            {
                if (textBoxusername.Text == g_datausermeg.Rows[i]["用户名"].ToString())
                {
                    isfound = true;
                    if (textBoxpassword.Text == g_datausermeg.Rows[i]["密码"].ToString())
                    {
                        ispasscrect = true;
                        break;

                    }
                }
            }

            if (isfound && ispasscrect)
            {
                return g_datausermeg.Rows[i]["用户名"].ToString();         
            }

             return null;
        }

        private void buttonloginin_Click(object sender, EventArgs e)
        {
            string username = null;

            username = userfound();
            if (username == null)
            {
                MessageBox.Show("用户名或者密码不对");
            }
            else {

                form1.username = username;
                this.Close();

            }
            

        }

        private void buttonmini_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void buttonclose_Click(object sender, EventArgs e)
        {
            Application.Exit();
           // this.Close();

        }

        private void buttonregister_Click(object sender, EventArgs e)
        {
            MessageBox.Show("直接在"+ Application.StartupPath + "\\注册用户.xls添加信息即可");
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                mouseflag = true;
                FormLocation = this.Location;
                mouseOffset = Control.MousePosition;
            }
        }

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            int _x = 0;
            int _y = 0;
            if (mouseflag == true)
            {
                Point pt = Control.MousePosition;
                _x = mouseOffset.X - pt.X;
                _y = mouseOffset.Y - pt.Y;

                this.Location = new Point(FormLocation.X - _x, FormLocation.Y - _y);
            }
        }

        private void panel1_MouseUp(object sender, MouseEventArgs e)
        {
            mouseflag = false;
        }

        private void Formuserpassword_Load(object sender, EventArgs e)
        {
            int SW = Screen.PrimaryScreen.Bounds.Width;
            int SH = Screen.PrimaryScreen.Bounds.Height;

            this.Location = (Point)new Point((SW-this.Size.Width)/2, (SH-this.Size.Height)/2-80);
        }
    }
}
