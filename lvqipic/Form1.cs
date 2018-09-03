using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Drawing.Drawing2D;
using MsWord = Microsoft.Office.Interop.Word;
using MySql.Data.MySqlClient;
using NPOI;
using NPOI.Util;
using NPOI.OpenXml4Net;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Globalization;

namespace lvqipic
{
   
    public partial class Form1 : Form
    {
        string  STUDENTMEMER = "Excel\\会员信息.xls";
        string  CLASSSYSTEM = "Excel\\课程体系.xls";

        public string g_model_imag_filename = null;
        public string g_name_date = null;

        private bool mouseflag = false;//鼠标是否按下
        private Point FormLocation;     //form的location
        private Point mouseOffset;      //鼠标的按下位置


        Image g_imagback = null;
        Image g_image = null;
        Image g_insertImagenail = null;
        Image g_dealImage = null;

        Bitmap g_savebmp = null;


        List<string> list = new List<string>();//定义list变量，存放获取到的路径


        //excel
        string mesgexcelfileName = null; //文件名

        DataTable g_data = null;
        DataTable g_data_class_submit = null;

        DataTable g_autoData = null;

        IWorkbook workbook = null;
      
        bool disposed;

        bool b_totalstartbutton_flag = false;

        int star_attendance_num = 5;
        int star_discipline_num = 5;
        int star_learn_num = 5;
        int star_equipment_num = 5;

        int g_totalstarnum = 0;
        int g_star_totalsum_temp = 0;


        int g_thisStarNum = 0;


        int star_posion_x;
        int star_posion_y;

        int thisstarnum_posion_x;
        int thisstarnum_posion_y;

        int date_posion_x;
        int date_posion_y;

        int picture_posion_x;
        int picture_posion_y;

        int INSERTPICTUREWIDTH = 928;
        int INSERTPICTUREHEIGHT = 1236;
        int INSERTPICTURESTART_X = 1404;//开始x坐标
        int INSERTPICTURESTART_Y = 310;//开始y坐标

        public string username = null;


        List<int> comboxnameindex = new List<int>();//定义list变量，获取到的对应老师的孩子下标


        private bool isview = false;
        public Form1()
        {
            InitializeComponent();

           // Formuserpassword Formmakemodel = new Formuserpassword(this);
           // Formmakemodel.ShowDialog();

            if (username == null)
            {
                username = "所有老师";
            }

            labelusername.Text = "用户名："+username;
            
            Comboxinit();
            backimginit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (username == null)
            {
                Application.Exit();
            }    
        }

       
        private void pictureBox1_DragDrop(object sender, DragEventArgs e)
        {
            //从拖放的事件中取得需要的数据,注意要转换成字符串数组
            //然后再从字符串数组中取值
            string filename = (string)e.Data.GetData(DataFormats.FileDrop);
             
           
        }

        private void pictureBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false) == true)
            {
                // 允许拖放动作继续,此时鼠标会显示为+
                e.Effect = DragDropEffects.All;
            }
        }

   
        public Bitmap CombinImage(Image imgBack, Image img, int start_x, int start_y)
        {

            Bitmap bmp = new Bitmap(imgBack.Width, imgBack.Height);

            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Transparent);

                g.DrawImage(imgBack, 0, 0, imgBack.Width, imgBack.Height); //g.DrawImage(imgBack, 0, 0, 相框宽, 相框高);

                g.DrawImage(img, start_x, start_y, img.Width, img.Height);

                GC.Collect();
            }

              
            return bmp;
        }
        public Bitmap Brightness(Image img, int brightness)
        {
            if (img == null)
            {
                return null;
            }

            if (brightness  <= 0)
            {
                return null;
            }

            Bitmap bmp = new Bitmap(img.Width, img.Height);

            using (Graphics g = Graphics.FromImage(bmp))
            {

               g.Clear(Color.Transparent);
               g.DrawImage(img, 0, 0, img.Width, img.Height);
               bmp = Img_color_brightness(bmp, brightness);
                
                GC.Collect();

            }
            return bmp;

            }
        public Bitmap MakeThumbnail(Image img, int width, int height, int brightness,int mode=0)
        {
            int towidth;
            int toheight;
            if (mode == 0)//以宽为准
            {
                towidth = width;
                toheight = (width * img.Height) / img.Width;
            }
            else //以高为准
            {
                toheight = height;
                towidth = (img.Width * height) / img.Height;
            }
            
            Bitmap bmp = new Bitmap(towidth, toheight);

            using (Graphics g = Graphics.FromImage(bmp))
            {

                g.Clear(Color.Transparent);
                g.DrawImage(img, 0, 0, towidth, toheight); //g.DrawImage(imgBack, 0, 0, 相框宽, 相框高);

                if (brightness != -1)
                {
                    bmp = Img_color_brightness(bmp, brightness);
                }
                GC.Collect();

            }

            return bmp;
        }
        public Bitmap MakeThumbnail(Image img, int width, int height, int mode = 0)
        {
            int towidth;
            int toheight;
            if (mode == 0)//以宽为准
            {
                towidth = width;
                toheight = (width * img.Height) / img.Width;
            }
            else //以高为准
            {
                toheight = height;
                towidth = (img.Width * height) / img.Height;
            }

            Bitmap bmp = new Bitmap(towidth, toheight);

            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Transparent);
                g.DrawImage(img, 0, 0, towidth, toheight); //g.DrawImage(imgBack, 0, 0, 相框宽, 相框高);
                GC.Collect();
            }

            return bmp;
        }



        /*
        private List<string> GetStringRows(Graphics graphic, Font font, string text, int width)
        {
            int RowBeginIndex = 0;
            int rowEndIndex = 0;
            int textLength = text.Length;
            List<string> textRows = new List<string>();

            for (int index = 0; index < textLength; index++)
            {
                rowEndIndex = index;

                if (index == textLength - 1)
                {
                    textRows.Add(text.Substring(RowBeginIndex));
                }
                else if (rowEndIndex + 1 < text.Length && text.Substring(rowEndIndex, 2) == "\r\n")
                {
                    textRows.Add(text.Substring(RowBeginIndex, rowEndIndex - RowBeginIndex));
                    rowEndIndex = index += 2;
                    RowBeginIndex = rowEndIndex;
                }
                else if (graphic.MeasureString(text.Substring(RowBeginIndex, rowEndIndex - RowBeginIndex + 1), font).Width > width)
                {
                    textRows.Add(text.Substring(RowBeginIndex, rowEndIndex - RowBeginIndex));
                    RowBeginIndex = rowEndIndex;
                }
            }

            return textRows;
        }
        */
        private List<string> GetStringRows(Graphics graphic, Font font, string text, int width)
        {
            int RowBeginIndex = 0;
            int rowEndIndex = 0;
            int textLength = text.Length;
            List<string> textRows = new List<string>();

            for (int index = 0; index < textLength; index++)
            {
                rowEndIndex = index;

                if (index == textLength - 1)
                {
                    textRows.Add(text.Substring(RowBeginIndex));
                }
                else if (rowEndIndex + 1 < text.Length && text.Substring(rowEndIndex, 2) == "\r\n")
                {
                    // MessageBox.Show("换行");
                    textRows.Add(text.Substring(RowBeginIndex, rowEndIndex - RowBeginIndex));
                    rowEndIndex = index += 2;
                    RowBeginIndex = rowEndIndex;
                }
                else if (graphic.MeasureString(text.Substring(RowBeginIndex, rowEndIndex - RowBeginIndex + 1), font).Width > width)
                {
                    textRows.Add(text.Substring(RowBeginIndex, rowEndIndex - RowBeginIndex));
                    RowBeginIndex = rowEndIndex;
                }
            }

            return textRows;
        }
        public void Markword_2(Graphics g, string str, int x, int y, string font, int size, Color wordcolor, int mode = 0, int width = 0, int rowsapce = 0)
        {
            SolidBrush sbrush = new SolidBrush(wordcolor);

            if (mode == 0)//不换行
            {
                g.DrawString(str, new Font(font, size), sbrush, new PointF(x, y));
            }
            else if (mode == 1)//换行
            {
                List<string> textRows = new List<string>();
                int row_y = y;
                textRows = GetStringRows(g, new Font(font, size), str, width);

                for (int i = 0; i < textRows.Count; i++)
                {
                    g.DrawString(textRows[i], new Font(font, size), sbrush, new PointF(x, row_y + (i * (new Font(font, size).Height + rowsapce))));
                }
            }

            return;

        }
        public Bitmap Markword(Image img, string str, int x, int y, string font, int size, Color wordcolor,int mode=0,int width=0,int rowsapce=0)
        {
            Bitmap bmp = new Bitmap(img.Width, img.Height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Transparent);

                g.DrawImage(img, 0, 0, img.Width, img.Height);

                SolidBrush sbrush = new SolidBrush(wordcolor);

                if (mode == 0)//不换行
                {
                    g.DrawString(str, new Font(font, size), sbrush, new PointF(x, y));
                }
                else if (mode == 1)//换行
                {
                    List<string> textRows = new List<string>();
                    int row_y = y;
                    textRows = GetStringRows(g, new Font(font, size), str, width);

                    for (int i = 0; i < textRows.Count; i++)
                    {
                        g.DrawString(textRows[i], new Font(font, size), sbrush, new PointF(x, row_y + (i * (new Font(font, size).Height + rowsapce))));
                    }
                }
            }

            GC.Collect();

            return bmp;
        }
        public int star_num()
        {

            //课堂出勤
            if (radioButton_attendance_1.Checked)
            {
                star_attendance_num = 1;
            }
            else if(radioButton_attendance_2.Checked)
            {
                star_attendance_num = 2;
            }
            else if (radioButton_attendance_3.Checked)
            {
                star_attendance_num = 3;
            }
            else if (radioButton_attendance_4.Checked)
            {
                star_attendance_num = 4;
            }
            else if (radioButton_attendance_5.Checked)
            {
                star_attendance_num = 5;
            }

            //课堂纪律
            if (radioButton_discipline_1.Checked)
            {
                star_discipline_num = 1;
                
            }
            else if (radioButton_discipline_2.Checked)
            {
                star_discipline_num = 2;
            }
            else if (radioButton_discipline_3.Checked)
            {
                star_discipline_num = 3;
            }
            else if (radioButton_discipline_4.Checked)
            {
                star_discipline_num = 4;
            }
            else if (radioButton_discipline_5.Checked)
            {
                star_discipline_num = 5;
            }

            //知识掌握
            if (radioButton_learn_1.Checked)
            {
                star_learn_num = 1;

            }
            else if (radioButton_learn_2.Checked)
            {
                star_learn_num = 2;
            }
            else if (radioButton_learn_3.Checked)
            {
                star_learn_num = 3;
            }
            else if (radioButton_learn_4.Checked)
            {
                star_learn_num = 4;
            }
            else if (radioButton_learn_5.Checked)
            {
                star_learn_num = 5;
            }

            //器材整理
            if (radioButton_equipment_1.Checked)
            {
                star_equipment_num = 1;

            }
            else if (radioButton_equipment_2.Checked)
            {
                star_equipment_num = 2;
            }
            else if (radioButton__equipment_3.Checked)
            {
                star_equipment_num = 3;
            }
            else if (radioButton_equipment_4.Checked)
            {
                star_equipment_num = 4;
            }
            else if (radioButton__equipment_5.Checked)
            {
                star_equipment_num = 5;
            }

            return (star_attendance_num + star_discipline_num + star_learn_num + star_equipment_num);
        }
        public Bitmap DrawStar_2(Image img, Image startimg, int starnum, int start_x, int start_y, int height, int width, int space)
        {
            int x = start_x;
            int y = start_y;

            startimg = MakeThumbnail(startimg, height, width);

            Bitmap bmp = new Bitmap(img.Width, img.Height);

            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Transparent);
                g.DrawImage(img, 0, 0, img.Width, img.Height); //g.DrawImage(imgBack, 0, 0, 相框宽, 相框高);

                for (int index = 0; index < starnum; index++)
                {
                    g.DrawImage(startimg, x + (startimg.Width + space) * index, y, startimg.Width, startimg.Height);

                }
            }

            GC.Collect();

            return bmp;

        }
        public Bitmap DrawStar(Image img,int starnum,int start_x,int start_y,int height,int width,int space)
        {
            int x = start_x;
            int y = start_y;

            Image startimg = Properties.Resources.star1;
            startimg = MakeThumbnail(startimg, height, width);

            Bitmap bmp = new Bitmap(img.Width, img.Height);

            Graphics g = Graphics.FromImage(bmp);

            g.Clear(Color.Transparent);
            g.DrawImage(img, 0, 0, img.Width, img.Height); //g.DrawImage(imgBack, 0, 0, 相框宽, 相框高);

            for (int index = 0; index < starnum; index++)
            {
                if(index == 0)
                {
                    g.DrawImage(Properties.Resources.star1, x + (startimg.Width + space) * index, y, startimg.Width, startimg.Height);
                }
                else if (index == 1)
                {
                    g.DrawImage(Properties.Resources.start2, x + (startimg.Width + space) * index, y, startimg.Width, startimg.Height);
                }
                else if (index == 2)
                {
                    g.DrawImage(Properties.Resources.star3, x + (startimg.Width + space) * index, y, startimg.Width, startimg.Height);
                }
                else if (index == 3)
                {
                    g.DrawImage(Properties.Resources.start4, x + (startimg.Width + space) * index, y, startimg.Width, startimg.Height);
                }
                else if (index == 4)
                {
                    g.DrawImage(Properties.Resources.start5, x + (startimg.Width + space) * index, y, startimg.Width, startimg.Height);
                }

            }
           
            GC.Collect();

            return bmp;

        }
        public Bitmap star(Image imag)
        {
            Bitmap bmp;
            int sum_star = 0;
            string star_totalsum;

            int totalstar = 0;

            sum_star = star_num();//本次得星
            // g_star_totalsum = g_star_totalsum + sum_star;//累计得星
            star_totalsum = textBox_toalstar.Text;

            //画星星
            int start_x = 400;
            int start_y = 1200;


           
            bmp = DrawStar(imag, star_attendance_num, start_x, start_y + (80*0), 70, 70, 30);
            bmp = DrawStar(bmp, star_discipline_num, start_x, start_y + (80 * 1), 70, 70, 30);
            bmp = DrawStar(bmp, star_learn_num, start_x, start_y + (80 * 2), 70, 70, 30);
            bmp = DrawStar(bmp, star_equipment_num, start_x, start_y + (80 * 3), 70, 70, 30);
            /*
           bmp = DrawStar_2(imag, Properties.Resources.start5, star_attendance_num, start_x, start_y + (80 * 0), 70, 70, 30);
           bmp = DrawStar_2(bmp, Properties.Resources.start5, star_discipline_num, start_x, start_y + (80 * 1), 70, 70, 30);
           bmp = DrawStar_2(bmp, Properties.Resources.start5, star_learn_num, start_x, start_y + (80 * 2), 70, 70, 30);
           bmp = DrawStar_2(bmp, Properties.Resources.start5, star_equipment_num, start_x, start_y + (80 * 3), 70, 70, 30);
           */
            bmp = Markword(bmp, "课堂出勤：" , 160, start_y + (80 * 0) + 10, "微软雅黑", 35, Color.Black);
            bmp = Markword(bmp, "课堂纪律：" , 160, start_y + (80 * 1) + 10, "微软雅黑", 35, Color.Black);
            bmp = Markword(bmp, "知识掌握：" , 160, start_y + (80 * 2) + 10, "微软雅黑", 35, Color.Black);
            bmp = Markword(bmp, "器件整理：" , 160, start_y + (80 * 3) + 10, "微软雅黑", 35, Color.Black);


            int addstar;
            //本次得星
            if (textBox_addstar.Text.Length == 0)
            {
                addstar = 0;
            }else
            {
                addstar = Convert.ToInt32(textBox_addstar.Text);
            }

            start_y += ((80 * 4)+20);

            bmp = Markword(bmp, "本次得星：", 160, start_y+5, "楷体", 35, Color.Blue);

            if (addstar != 0)
            {
                bmp = Markword(bmp, sum_star.ToString() + "+", 410, start_y, "楷体", 40,Color.Red);
                bmp = Markword(bmp, textBox_addstar.Text, 490, start_y , "楷体", 40, Color.Red);
            }
            else {
                bmp = Markword(bmp, sum_star.ToString(), 410, start_y, "楷体", 40, Color.Red);
            }

            //累计得星
            bmp = Markword(bmp, "累计得星：", 160, start_y + 80+5, "楷体", 35, Color.Red);

            if (g_data != null && comboBox_name.SelectedIndex >= 0)
            {

                if (textBox_toalstar.Enabled == true)
                {
                     totalstar = Convert.ToInt32(textBox_toalstar.Text);
                }
                else {
                     totalstar = Convert.ToInt32(g_data.Rows[comboxnameindex[comboBox_name.SelectedIndex]]["累计得星"]);
                   
                }
               

               

                totalstar = totalstar + sum_star + addstar;

                //更新累计得星数据到g_data 
                if (g_data != null && comboBox_name.SelectedIndex >= 0)
                {
                   // MessageBox.Show("update");
                    textBox_toalstar.Text = totalstar.ToString();
                    label1stardate.Text = monthCalendar_date.SelectionStart.Year + "/" + monthCalendar_date.SelectionStart.Month + "/" + monthCalendar_date.SelectionStart.Day;
                    //g_data.Rows[comboBox_name.SelectedIndex]["累计得星"] = textBox_toalstar.Text;
                }

                bmp = Markword(bmp, totalstar.ToString(), 410, start_y + 80, "楷体", 40, Color.Red);
            }
             
          
            return bmp;
        }
        public Bitmap autostar(Image imag, int totalstar, int starnum1, int starnum2, int starnum3, int starnum4,int addstar)
        {
            Bitmap bmp;
            int sum_star = 0;
            
            int index = -1;
   
            //画星星
            int start_x = 400;
            int start_y = 1200;

            bmp = DrawStar(imag, starnum1, start_x, start_y + (80 * 0), 70, 70, 30);
            bmp = DrawStar(bmp, starnum2, start_x, start_y + (80 * 1), 70, 70, 30);
            bmp = DrawStar(bmp, starnum3, start_x, start_y + (80 * 2), 70, 70, 30);
            bmp = DrawStar(bmp, starnum4, start_x, start_y + (80 * 3), 70, 70, 30);
         
            bmp = Markword(bmp, "课堂出勤：", 160, start_y + (80 * 0) + 10, "微软雅黑", 35, Color.Black);
            bmp = Markword(bmp, "课堂纪律：", 160, start_y + (80 * 1) + 10, "微软雅黑", 35, Color.Black);
            bmp = Markword(bmp, "知识掌握：", 160, start_y + (80 * 2) + 10, "微软雅黑", 35, Color.Black);
            bmp = Markword(bmp, "器件整理：", 160, start_y + (80 * 3) + 10, "微软雅黑", 35, Color.Black);



            //本次得星

            sum_star = starnum1 + starnum2 + starnum3 + starnum4;
            start_y += ((80 * 4) + 20);

            bmp = Markword(bmp, "本次得星：", 160, start_y + 5, "楷体", 35, Color.Blue);

            if (addstar != 0)
            {
                bmp = Markword(bmp, sum_star.ToString() + "+", 410, start_y, "楷体", 40, Color.Red);
                bmp = Markword(bmp, addstar.ToString(), 490, start_y, "楷体", 40, Color.Red);
            }
            else
            {
                bmp = Markword(bmp, sum_star.ToString(), 410, start_y, "楷体", 40, Color.Red);
            }

            //累计得星
            bmp = Markword(bmp, "累计得星：", 160, start_y + 80 + 5, "楷体", 35, Color.Red);

            totalstar = totalstar + sum_star + addstar;

            bmp = Markword(bmp, totalstar.ToString(), 410, start_y + 80, "楷体", 40, Color.Red);


            g_data.Rows[index]["累计得星"] = totalstar;
            g_data.Rows[index]["更新时间"] = DateTime.Now.Year.ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Day.ToString();

            return bmp;
        }
        public Bitmap insertpPicture(Image imgback, Image img, int rangerstart_x, int rangestart_y, int rangewidth, int rangeheight)
        {
            int x = 0;
            int y = 0;
            Bitmap bmp;
            Image pictureimg;

            if (img.Width / img.Height > rangewidth / rangeheight)
            {
                pictureimg = MakeThumbnail(img, rangewidth, rangeheight, 0, trackBarbright.Value);//缩略图以宽为准
                
                x = rangerstart_x + (rangewidth - pictureimg.Width) / 2;
                y = rangestart_y + (rangeheight - pictureimg.Height) / 2;
               
            }
            else {

                pictureimg = MakeThumbnail(img, rangewidth, rangeheight, 1, trackBarbright.Value);////缩略图以高为准
                x = rangerstart_x + (rangewidth - pictureimg.Width) / 2;
                y = rangestart_y + (rangeheight - pictureimg.Height) / 2;
             
            }
            bmp = CombinImage(imgback, pictureimg, x, y);
            return bmp;
        }

        public static unsafe Bitmap Img_color_brightness(Bitmap src, int brightness)
        {
            int width = src.Width;
            int height = src.Height;
            Bitmap back = new Bitmap(width, height);
            Rectangle rect = new Rectangle(0, 0, width, height);
            //这种速度最快
            System.Drawing.Imaging.BitmapData bmpData = src.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadWrite, System.Drawing.Imaging.PixelFormat.Format24bppRgb);//24位rgb显示一个像素，即一个像素点3个字节，每个字节是BGR分量。Format32bppRgb是用4个字节表示一个像素
            byte* ptr = (byte*)(bmpData.Scan0);
            for (int j = 0; j < height; j++)
            {
                for (int i = 0; i < width; i++)
                {
                    //ptr[2]为r值，ptr[1]为g值，ptr[0]为b值
                    int red = ptr[2] + brightness; if (red > 255) red = 255; if (red < 0) red = 0;
                    int green = ptr[1] + brightness; if (green > 255) green = 255; if (green < 0) green = 0;
                    int blue = ptr[0] + brightness; if (blue > 255) blue = 255; if (blue < 0) blue = 0;
                    back.SetPixel(i, j, Color.FromArgb(red, green, blue));
                    ptr += 3; //Format24bppRgb格式每个像素占3字节
                }
                ptr += bmpData.Stride - bmpData.Width * 3;//每行读取到最后“有用”数据时，跳过未使用空间XX
            }
            src.UnlockBits(bmpData);
            return back;
        }
        private void backimginit()
        {
            string FileFullName;
            Random rd = new Random();
            int random = rd.Next(0, 4);

            /*
            if (random == 0)
            {
                FileFullName = Application.StartupPath + "\\主题\\课后总结主题默认.png";
            }
            else {

                FileFullName = Application.StartupPath + "\\主题\\课后总结主题"+ random.ToString()+ ".png";
            }
            if (!File.Exists(@FileFullName))
            {
                return;
            }
            */
            FileFullName = Application.StartupPath + "\\主题\\课后总结主题2.png";

            g_imagback = Image.FromFile(FileFullName);

            if ((g_imagback.Width != 2480) || (g_imagback.Height != 1748))
            {
                MessageBox.Show("图片尺寸有误（2480*1748）：" + g_imagback.Width.ToString() + " " + g_imagback.Height.ToString());
                return;
            }
            pictureBox1.Image = MakeThumbnail(g_imagback, pictureBox1.Width, pictureBox1.Height);
        }

        //选择课程总结模板
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.InitialDirectory =(Application.StartupPath+"\\主题");

            dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件夹";
            dialog.Filter = "图片|*.png;*.jpg;*.JPG;*JPEG";         

            if (DialogResult.OK == dialog.ShowDialog())
            {
                g_imagback = Image.FromFile(dialog.FileName);

                g_model_imag_filename = System.IO.Path.GetFileNameWithoutExtension(dialog.FileName);//文件名  “Default.aspx”

                if ((g_imagback.Width != 2480) || (g_imagback.Height != 1748))
                {
                    MessageBox.Show("图片尺寸有误（2480*1748）：" + g_imagback.Width.ToString() + " " + g_imagback.Height.ToString());
                    return;
                }
                pictureBox1.Image = MakeThumbnail(g_imagback, pictureBox1.Width, pictureBox1.Height);
            }
        }

        //插入会员图片
        private void button_doc_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();

            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);//默认打开桌面

            dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件夹";
            dialog.Filter = "图片|*.png;*.jpg;*.PNG;*.JPG;*.JPEG;*.jpeg";

            trackBarbright.Value = 0;//亮度默认为0

            if (DialogResult.OK == dialog.ShowDialog())
            {
               
                g_image = Image.FromFile(dialog.FileName);
       

                Bitmap bmp;
                if (g_imagback == null)
                {
                    MessageBox.Show("请先添加课后总结模板！");
                    return;
                }

                bmp = insertpPicture(g_imagback, g_image, 1400 + 5, 310 + 5, 928 - 10, 1236 - 10);//会员图片合成

                pictureBox1.Image = MakeThumbnail(bmp, pictureBox1.Width, pictureBox1.Height);

            }

        }
 
        //保存图片        
        private void button2_save_Click(object sender, EventArgs e)
        {
       
     
            string path = string.Empty;
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                path = fbd.SelectedPath;
                savepicture(path);        
            }
            return;

          
     
        }



        public void makePicture()
        {/*
            g_image = null;
            string classtitle = g_autoData.Rows[i]["课程主题"].ToString();
            string pictureYear = g_autoData.Rows[i]["年"].ToString();
            string pictureMonth = g_autoData.Rows[i]["月"].ToString();
            string pictureDay = g_autoData.Rows[i]["日"].ToString();
            string name = g_autoData.Rows[i]["会员姓名"].ToString();
            string picturfilename = g_autoData.Rows[i]["照片名称"].ToString();
            string teacherSummit = g_autoData.Rows[i]["课后评价"].ToString();
            int starNum_1 = Convert.ToInt32(g_autoData.Rows[i]["课堂出勤"].ToString());
            int starNum_2 = Convert.ToInt32(g_autoData.Rows[i]["课堂纪律"].ToString());
            int starNum_3 = Convert.ToInt32(g_autoData.Rows[i]["知识掌握"].ToString());
            int starNum_4 = Convert.ToInt32(g_autoData.Rows[i]["器件整理"].ToString());
            int starNum_add = Convert.ToInt32(g_autoData.Rows[i]["加扣星数"].ToString());

            Bitmap bmp = new Bitmap(g_imagback.Width, g_imagback.Height);

            using (Graphics g = Graphics.FromImage(bmp))
            {

                g.Clear(Color.Transparent);

                //画背景
                g.DrawImage(g_imagback, 0, 0, g_imagback.Width, g_imagback.Height); //g.DrawImage(imgBack, 0, 0, 相框宽, 相框高);

                //要插图片的起始位置和可以显示的范围
                int insertPictureRangeWidth = 928 - 10;
                int insertPictureRangeHeight = 1236 - 10;
                int insertPictureStart_x = 1400 + 5;
                int insertPictureStart_y = 310 + 5;

                //插入图片的起始位置和大小
                int start_x = -1;
                int start_y = -1;
                int towidth = -1;
                int toheight = -1;

                //根据图片的实际大小计算插入的位置和宽高
                if (g_image.Width / g_image.Height > insertPictureRangeWidth / insertPictureRangeHeight)
                {
                    towidth = insertPictureRangeWidth;
                    toheight = (insertPictureRangeWidth * g_image.Height) / g_image.Width;

                    start_x = insertPictureStart_x + (insertPictureRangeWidth - towidth) / 2;
                    start_y = insertPictureStart_y + (insertPictureRangeHeight - toheight) / 2;

                }
                else
                {
                    toheight = insertPictureRangeHeight;
                    towidth = (g_image.Width * toheight) / g_image.Height;
                    start_x = insertPictureStart_x + (insertPictureRangeWidth - towidth) / 2;
                    start_y = insertPictureStart_y + (insertPictureRangeHeight - toheight) / 2;
                }

                //画要插入的图片
                g.DrawImage(g_image, start_x, start_y, towidth, toheight);


                List<int> starnum = new List<int>();
                starnum.Add(starNum_1);
                starnum.Add(starNum_2);
                starnum.Add(starNum_3);
                starnum.Add(starNum_4);

                List<Bitmap> star = new List<Bitmap>();
                star.Add(Properties.Resources.star1);
                star.Add(Properties.Resources.start2);
                star.Add(Properties.Resources.star3);
                star.Add(Properties.Resources.start4);
                star.Add(Properties.Resources.start5);



                int StarWidth = 70;
                int StarHeight = 70;
                int StarSpace = 30;
                int LineSpace = 10;

                int starStart_X = 400;
                int starStart_Y = 1200;

                for (int j = 0; j < 4; j++)
                {
                    for (int k = 0; k < starnum[j]; k++)
                    {
                        g.DrawImage(star[k], starStart_X + (StarWidth + StarSpace) * k, starStart_Y + (StarHeight + LineSpace) * j, StarWidth, StarHeight);
                    }
                }

                Markword_2(g, "课堂出勤：", 160, starStart_Y + (80 * 0) + 10, "微软雅黑", 35, Color.Black);
                Markword_2(g, "课堂纪律：", 160, starStart_Y + (80 * 1) + 10, "微软雅黑", 35, Color.Black);
                Markword_2(g, "知识掌握：", 160, starStart_Y + (80 * 2) + 10, "微软雅黑", 35, Color.Black);
                Markword_2(g, "器件整理：", 160, starStart_Y + (80 * 3) + 10, "微软雅黑", 35, Color.Black);


                starStart_Y += ((80 * 4) + 20);

                int thisStarNum = starNum_1 + starNum_2 + starNum_3 + starNum_4 + starNum_add;

                Markword_2(g, thisStarNum.ToString(), starStart_X + 10, starStart_Y, "楷体", 40, Color.Red);


                int totalstar = lasttotalstar + thisStarNum;

                Markword_2(g, totalstar.ToString(), starStart_X + 10, starStart_Y + 80, "楷体", 40, Color.Red);// Markword_2(g,);



                Markword_2(g, "本次得星：", 160, starStart_Y + 5, "楷体", 35, Color.DeepSkyBlue);
                Markword_2(g, "累计得星：", 160, starStart_Y + 80 + 5, "楷体", 35, Color.DeepPink);


                Markword_2(g, "课程主题：" + classtitle + "\r\n" + "课程目标：\r\n" + classfors[i], 160, 340, "微软雅黑", 35, Color.Black, 1, 1100, 20);///课程目标

                if (g_autoData.Rows[i]["课后评价"].ToString().Length != 0)
                {
                    Markword_2(g, "课后评价：\r\n" + teacherSummit, 160, 870, "微软雅黑", 31, Color.Black, 1, 1100, 18);//课后评价
                }

                Markword_2(g, name, 2020, 1560, "微软雅黑", 31, Color.Red);//名字            
                Markword_2(g, pictureYear + "年" + pictureMonth + "月" + pictureDay + "日",
                    2020, 1610, "微软雅黑", 31, Color.Red);//合成日期


            }
            */
        }


        //预览图片
        private void buttonview_Click(object sender, EventArgs e)
        {   

            if (g_imagback == null)
            {
                MessageBox.Show("请添加背景！");
                return;
            }
            if (g_image == null || g_dealImage == null)
            {
                MessageBox.Show("请添加会员图片！");
                return;
            }
            if (comboBoxtitle.Text.Length == 0)
            {
                MessageBox.Show("请先选择课程主题！");
                return;
            }
            if (textBoxclassfor.Text.Length == 0)
            {
                MessageBox.Show("课程目标未正常加载，请检查！");
                return;

            }
          
            if (textBox_toalstar.Enabled == true)
            {
                MessageBox.Show("修改累计得星后，请先点击完成！");
                return;
            }

            if (comboBox_name.Text.Length == 0)
            {
                MessageBox.Show("请选择会员姓名！");
                return;
            }
            if (textBox_toalstar.Text.Length== 0)
            {
                MessageBox.Show("累计得星内容不能为空！");
                return;
            }
            if (textBox_date.Text.Length == 0)
            {
                MessageBox.Show("未选择日期！");
                return;
            }
            

            

            Bitmap bmp = new Bitmap(g_imagback.Width, g_imagback.Height);

            using (Graphics g = Graphics.FromImage(bmp))
            {

                g.Clear(Color.Transparent);

                //画背景
                g.DrawImage(g_imagback, 0, 0, g_imagback.Width, g_imagback.Height); //g.DrawImage(imgBack, 0, 0, 相框宽, 相框高);

                //要插图片的起始位置和可以显示的范围


                //插入图片的起始位置和大小
                int start_x = -1;
                int start_y = -1;
                int towidth = -1;
                int toheight = -1;

                start_x = INSERTPICTURESTART_X + (INSERTPICTUREWIDTH - g_dealImage.Width) / 2;
                start_y = INSERTPICTURESTART_Y + (INSERTPICTUREHEIGHT - g_dealImage.Height) / 2;

               /* MessageBox.Show(g_dealImage.Width.ToString()+" "+ g_dealImage.Height.ToString()+
                    " "+ start_x+" " +start_y);*/
                //画要插入的图片
                g.DrawImage(g_dealImage, start_x, start_y, g_dealImage.Width, g_dealImage.Height);

                int thisStarNum = star_num();

                g_thisStarNum = thisStarNum;

                List<int> starnum = new List<int>();
                starnum.Add(star_attendance_num);
                starnum.Add(star_discipline_num);
                starnum.Add(star_learn_num);
                starnum.Add(star_equipment_num);

                List<Bitmap> star = new List<Bitmap>();
                star.Add(Properties.Resources.star1);
                star.Add(Properties.Resources.start2);
                star.Add(Properties.Resources.star3);
                star.Add(Properties.Resources.start4);
                star.Add(Properties.Resources.start5);



                int StarWidth = 70;
                int StarHeight = 70;
                int StarSpace = 30;
                int LineSpace = 10;

                int starStart_X = 400;
                int starStart_Y = 1200;

                for (int j = 0; j < 4; j++)
                {
                    for (int k = 0; k < starnum[j]; k++)
                    {
                        g.DrawImage(star[k], starStart_X + (StarWidth + StarSpace) * k, starStart_Y + (StarHeight + LineSpace) * j, StarWidth, StarHeight);
                    }
                }

                Markword_2(g, "课堂出勤：", 160, starStart_Y + (80 * 0) + 10, "微软雅黑", 35, Color.Black);
                Markword_2(g, "课堂纪律：", 160, starStart_Y + (80 * 1) + 10, "微软雅黑", 35, Color.Black);
                Markword_2(g, "知识掌握：", 160, starStart_Y + (80 * 2) + 10, "微软雅黑", 35, Color.Black);
                Markword_2(g, "器件整理：", 160, starStart_Y + (80 * 3) + 10, "微软雅黑", 35, Color.Black);


                starStart_Y += ((80 * 4) + 20);


                //本次得星
                Markword_2(g, thisStarNum.ToString(), starStart_X + 10, starStart_Y, "楷体", 40, Color.Red);

                //加扣星星
                if (textBox_addstar.Text.Length > 0   && Convert.ToInt32(textBox_addstar.Text) > 0)//加扣星星
                {
                    string addstarstr = " +" + textBox_addstar.Text;

                    Markword_2(g, addstarstr, starStart_X + 50, starStart_Y, "楷体", 40, Color.Red);

                    thisStarNum += Convert.ToInt32(textBox_addstar.Text);
                }
                int lasttotalstar = -1;              
                lasttotalstar = Convert.ToInt32(g_data.Rows[comboxnameindex[comboBox_name.SelectedIndex]]["累计得星"]);
                g_totalstarnum = lasttotalstar + thisStarNum;
                //累计得星
                Markword_2(g, g_totalstarnum.ToString(), starStart_X + 10, starStart_Y + 80, "楷体", 40, Color.Red);// Markword_2(g,);



                Markword_2(g, "本次得星：", 160, starStart_Y + 5, "楷体", 35, Color.DeepSkyBlue);
                Markword_2(g, "累计得星：", 160, starStart_Y + 80 + 5, "楷体", 35, Color.DeepPink);


                Markword_2(g, "课程主题：" + comboBoxtitle.Text + "\r\n" + "课程目标：\r\n" + textBoxclassfor.Text, 160, 340, "微软雅黑", 35, Color.Black, 1, 1100, 20);///课程目标

                if (textBoxteacher.Text.Length != 0)
                {
                    Markword_2(g, "课后评价：\r\n" + textBoxteacher.Text, 160, 870, "微软雅黑", 31, Color.Black, 1, 1100, 18);//课后评价
                }

                Markword_2(g, comboBox_name.Text, 2010, 1560, "微软雅黑", 31, Color.Red);//名字            
                Markword_2(g, textBox_date.Text,2010, 1610, "微软雅黑", 31, Color.Red);//合成日期


                g_savebmp = bmp;

                pictureBox1.Image = MakeThumbnail(bmp, pictureBox1.Width, pictureBox1.Height);
                isview = true;

            }
              /*

           bmp = star(g_imagback);//会员积星合成

               bmp = insertpPicture(bmp, g_image,1400+5,310+5,928-10,1236-10);//会员图片合成


               bmp = Markword(bmp, comboBox_name.Text, 2020, 1560 ,"微软雅黑", 31, Color.Red);//名字

               bmp = Markword(bmp, textBox_date.Text, 2020, 1610, "微软雅黑", 31, Color.Red);//合成日期


               bmp = Markword(bmp, "课程主题："+comboBoxtitle.Text + "\r\n" +"课程目标：\r\n"+ textBoxclassfor.Text, 160, 340, "微软雅黑", 35, Color.Black,1,1100,20);///课程目标

               if (textBoxteacher.Text.Length != 0)
               {
                   bmp = Markword(bmp, "课后评价：\r\n" + textBoxteacher.Text, 160, 870, "微软雅黑", 31, Color.Black, 1, 1100, 18);//课后评价

               }


               g_savebmp = bmp;

               pictureBox1.Image = MakeThumbnail(bmp, pictureBox1.Width, pictureBox1.Height);

               isview = true;
               */
           }

       private void monthCalendar_date_DateSelected(object sender, DateRangeEventArgs e)
       {
           string month = null;
           string year = null;
           string day = null;
            year = monthCalendar_date.SelectionStart.Year.ToString();
           if (monthCalendar_date.SelectionStart.Month < 10)
           {
               month = "0" + monthCalendar_date.SelectionStart.Month.ToString();
           }
           else {
               month = monthCalendar_date.SelectionStart.Month.ToString();
           }

           if (monthCalendar_date.SelectionStart.Day < 10)
           {
               day = "0"+monthCalendar_date.SelectionStart.Day.ToString();
           }
           else {
               day = monthCalendar_date.SelectionStart.Day.ToString();
           }


           textBox_date.Text = year + "年" + month + "月" + day +"日";
           g_name_date = year + month  + day;//保存名字时使用

           monthCalendar_date.Visible = false;
       }

       private void textBox_addstar_KeyPress(object sender, KeyPressEventArgs e)
       {
           //如果输入的不是数字键，也不是回车键、Backspace键，则取消该输入
           if (!(Char.IsNumber(e.KeyChar)) && e.KeyChar != (char)13 && e.KeyChar != (char)8)
           {
               e.Handled = true;
           }
       }

       private void textBox_toalstar_KeyPress(object sender, KeyPressEventArgs e)
       {

           //如果输入的不是数字键，也不是回车键、Backspace键，则取消该输入
           if (!(Char.IsNumber(e.KeyChar)) && e.KeyChar != (char)13 && e.KeyChar != (char)8)
           {
               e.Handled = true;
           }
       }



       //调节亮度
       private void trackBarbright_MouseUp(object sender, MouseEventArgs e)
       {
           Bitmap bmp;
           if (g_imagback == null)
           {
               MessageBox.Show("请添加模板！");
               trackBarbright.Value = 0;
               return;
           }
           if (g_image == null)
           {
               MessageBox.Show("请添加图片！");
               trackBarbright.Value = 0;
               return;
           }
            if (trackBarbright.Value != -1)
            {
              //  Bitmap bmp = new Bitmap(towidth, toheight);
               // g_dealImage =  MakeThumbnail(bmp, pictureBox1.Width, pictureBox1.Height);
            }

            //bmp = insertpPicture(g_imagback, g_image, 1400 + 5, 310 + 5, 928 - 10, 1236 - 10);//会员图片合成
            if (trackBarbright.Value > 0)
            {
                g_dealImage = Brightness(g_insertImagenail, trackBarbright.Value);
            }
            else {
                g_dealImage = g_insertImagenail;
            }
            if (g_dealImage.Width / g_dealImage.Height > pictureBox_student.Width / pictureBox_student.Height)
            {
                pictureBox_student.Image = MakeThumbnail(g_dealImage, pictureBox_student.Width - 4, pictureBox_student.Height - 4, 0);
            }
            else
            {
                pictureBox_student.Image = MakeThumbnail(g_dealImage, pictureBox_student.Width - 4, pictureBox_student.Height - 4, 1);
            }



        }

     /*
       private void textBox_picture_paths_DragEnter(object sender, DragEventArgs e)
       {
           if (e.Data.GetDataPresent(DataFormats.FileDrop, false) == true)
           {
               // 允许拖放动作继续,此时鼠标会显示为+
               e.Effect = DragDropEffects.All;
           }
       }

       private void textBox_picture_paths_DragDrop(object sender, DragEventArgs e)
       {
           trackBarbright.Value = 0;//亮度默认为0

           string filename = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();

           if (System.IO.Path.GetExtension(filename).Equals(".jpg") || System.IO.Path.GetExtension(filename).Equals(".png"))
           {
               g_image = Image.FromFile(filename);
               textBox_picture_paths.Text = filename;

               Bitmap bmp;
               if (g_imagback == null)
               {
                   MessageBox.Show("请先添加课后总结模板！");
                   return;
               }
               bmp = insertpPicture(g_imagback, g_image, 1400 + 5, 310 + 5, 928 - 10, 1236 - 10);//会员图片合成

               pictureBox1.Image = MakeThumbnail(bmp, pictureBox1.Width, pictureBox1.Height);

           }
           else
           {
               MessageBox.Show("请选择图片png或jpg文件，本文件类型：" + System.IO.Path.GetExtension(filename));
           }


       }
       */
            private void buttonmakemodel_Click(object sender, EventArgs e)
        {
            Form2 Formmakemodel = new Form2();
           // Formmakemodel.ShowDialog();
            Formmakemodel.Show();


        }



        //处理excel
        /// <summary>
        /// 将DataTable数据导入到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <param name="sheetName">要导入的excel的sheet的名称</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public int DataTableToExcel(string fileName, DataTable data, string sheetName, bool isColumnWritten)
        {
            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;
            FileStream fs = null;

            try
            {
                fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                workbook = new XSSFWorkbook();
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                workbook = new HSSFWorkbook();


                if (workbook != null)
                {
                    sheet = workbook.CreateSheet(sheetName);
                }
                else
                {
                    return -1;
                }

                if (isColumnWritten == true) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                    }
                    count = 1;
                }
                else
                {
                    count = 0;
                }

                for (i = 0; i < data.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }
                workbook.Write(fs); //写入到excel
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                if (fs != null)
                {
                    fs.Close();
                    fs = null;
                }
                return -1;
            }

            if (fs != null)
            {
                fs.Close();
                fs = null;
            }
            return 0;
        }

        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public DataTable ExcelToDataTable(string fileName,string sheetName, bool isFirstRowColumn, string fliterecolumn = null, string fliterrow=null)
        {
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 0;
            FileStream fs = null;
            int fliterindex = -1;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fs == null)
                {
                    MessageBox.Show("打开文件失败："+ fileName);
                    return null;
                }
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                    workbook = new XSSFWorkbook(fs);
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook(fs);

                if (sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                }
                else
                {
                    sheet = workbook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

                    if (isFirstRowColumn)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue == fliterecolumn)//找到要过滤的列表下表
                                {
                                    fliterindex = i;
                                }
                                if (cellValue != null)
                                {
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }

                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　　　　　　

                        if (fliterindex != -1 && fliterrow != null)//按某一列（fliterindex）过滤行
                        {
                            if (row.GetCell(fliterindex).ToString() != fliterrow)
                            {
                                continue;
                            }
                        }
                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                                dataRow[j] = row.GetCell(j).ToString();
                        }
                        data.Rows.Add(dataRow);
                    }
                }
            }
            catch (Exception ex)
            {
                if (fs != null)
                {
                    fs.Close();
                    fs = null;
                }
                MessageBox.Show(ex.Message+"/"+fileName + "打开失败，请检查文件是否已经在其他程序打开");

                textBoxteacher.Text = ex.Message;
                Console.WriteLine("Exception: " + ex.Message);

                return null;
            }
           
            if (fs != null)
            {
                fs.Close();
                fs = null;
            }
          //  Dispose(fs);
            return data;


        }
        public void Dispose(FileStream fs)
        {
      /*
            
            Dispose(true);

            if (!this.disposed)
            {
     
                if (fs != null)
                    fs.Close();
         
                fs = null;
                disposed = true;
            }
           
          
           */
        }


        private void comboBox_name_TextChanged(object sender, EventArgs e)
        {


        }

        public void Comboxinit()
        {

            g_data = ExcelToDataTable(Application.StartupPath+ "\\"+STUDENTMEMER, null, true);

           //MessageBox.Show(g_data.Rows.Count.ToString()+" "+g_data.Columns.Count.ToString());
          
            if (g_data != null)
            {
                for (int i = 0; i < g_data.Rows.Count; i++)
                {
                    if (g_data.Rows[i]["授课老师"].ToString().Contains(","))
                    {
                        string[] teachers = g_data.Rows[i]["授课老师"].ToString().Split(',');
                        for (int j = 0; j < teachers.Length; j++)
                        {
                            if (teachers[j] == username)
                            {
                                comboBox_name.Items.Add(g_data.Rows[i]["姓名"].ToString());
                                comboxnameindex.Add(i);
                                break;
                            }
                        }
                    }
                    else if (g_data.Rows[i]["授课老师"].ToString() == username)
                    {
                        comboBox_name.Items.Add(g_data.Rows[i]["姓名"].ToString());
                        comboxnameindex.Add(i);
                    }           
                }
            }

            g_data_class_submit = ExcelToDataTable(Application.StartupPath + "\\"+CLASSSYSTEM, null, true);
            if (g_data_class_submit != null)
            {
              
                for (int i = 0; i < g_data_class_submit.Rows.Count; i++)
                {
                  
                    comboBoxtitle.Items.Add(g_data_class_submit.Rows[i]["课程主题"].ToString());
                    
                }
            }

        }


         public int   savedata()
        {
            int ret = -1;
            if(comboBox_name.SelectedIndex >=0  && textBox_toalstar.Text != null)
            {   
                g_data.Rows[comboxnameindex[comboBox_name.SelectedIndex]]["累计得星"] = g_totalstarnum;
                g_data.Rows[comboxnameindex[comboBox_name.SelectedIndex]]["更新时间"] = monthCalendar_date.SelectionStart.Year + "/" + monthCalendar_date.SelectionStart.Month + "/" + monthCalendar_date.SelectionStart.Day;
              
            }

            ret = DataTableToExcel(Application.StartupPath + "\\"+ STUDENTMEMER, g_data, "会员信息", true);

            if (ret != -1)
                return 0;
            else
                return -1;
        }

        private void comboBoxtitle_SelectedIndexChanged(object sender, EventArgs e)
        {
            string calssaim;
            if (g_data_class_submit != null)
            {
                calssaim = g_data_class_submit.Rows[comboBoxtitle.SelectedIndex]["课程目标"].ToString();

                calssaim = calssaim.Replace("。", "。\r\n");//在每个句号后面加换行

                textBoxclassfor.Text = calssaim;

            }
        }

        private void button_doc_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false) == true)
            {
                // 允许拖放动作继续,此时鼠标会显示为+
                e.Effect = DragDropEffects.All;
            }

        }

        private void button_doc_DragDrop(object sender, DragEventArgs e)
        {
            trackBarbright.Value = 0;//亮度默认为0


            string filename = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();

            if (System.IO.Path.GetExtension(filename).Equals(".jpg") || 
                System.IO.Path.GetExtension(filename).Equals(".JPG") ||
                 System.IO.Path.GetExtension(filename).Equals(".jpeg") ||
                 System.IO.Path.GetExtension(filename).Equals(".JPEG") ||
                  System.IO.Path.GetExtension(filename).Equals(".PNG") ||
                   System.IO.Path.GetExtension(filename).Equals(".png"))
            {
                g_image = Image.FromFile(filename);
            

                Bitmap bmp;
                if (g_imagback == null)
                {
                    MessageBox.Show("请先添加课后总结模板！");
                    return;
                }

                bmp = insertpPicture(g_imagback, g_image, 1400 + 5, 310 + 5, 928 - 10, 1236 - 10);//会员图片合成

                pictureBox1.Image = MakeThumbnail(bmp, pictureBox1.Width, pictureBox1.Height);

            }
            else
            {
                MessageBox.Show("请选择图片png、jpg、jpeg文件，本文件类型：" + System.IO.Path.GetExtension(filename));
            }
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false) == true)
            {
                // 允许拖放动作继续,此时鼠标会显示为+
                e.Effect = DragDropEffects.All;
            }
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            trackBarbright.Value = 0;//亮度默认为0


            string filename = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();

            if (System.IO.Path.GetExtension(filename).Equals(".jpg") ||
                System.IO.Path.GetExtension(filename).Equals(".JPG") ||
                 System.IO.Path.GetExtension(filename).Equals(".jpeg") ||
                 System.IO.Path.GetExtension(filename).Equals(".JPEG") ||
                  System.IO.Path.GetExtension(filename).Equals(".PNG") ||
                   System.IO.Path.GetExtension(filename).Equals(".png"))
            {
                g_image = Image.FromFile(filename);
                if (g_image == null)
                {
                    MessageBox.Show("图片加载失败！"+ filename);
                    return;
                }

                Bitmap bmp;
                if (g_imagback == null)
                {
                    MessageBox.Show("请先添加课后总结模板！");
                    return;
                }

                g_dealImage = g_image;

                if (g_dealImage.Width / g_dealImage.Height > pictureBox_student.Width / pictureBox_student.Height)
                {
                    pictureBox_student.Image = MakeThumbnail(g_dealImage, pictureBox_student.Width - 4, pictureBox_student.Height - 4, 0);
                }
                else
                {
                    pictureBox_student.Image = MakeThumbnail(g_dealImage, pictureBox_student.Width - 4, pictureBox_student.Height - 4, 1);
                }

                if (g_dealImage.Width / g_dealImage.Height > INSERTPICTUREWIDTH / INSERTPICTUREHEIGHT)
                {
                    g_dealImage = MakeThumbnail(g_dealImage, INSERTPICTUREWIDTH, INSERTPICTUREHEIGHT, 0);//以宽为准
                }
                else
                {
                    g_dealImage = MakeThumbnail(g_dealImage, INSERTPICTUREWIDTH, INSERTPICTUREHEIGHT, 1);//以高为准
                }

                g_insertImagenail = g_dealImage;//用于记录初始缩略状态


            }
            else
            {
                MessageBox.Show("请选择图片png、jpg、jpeg文件，本文件类型：" + System.IO.Path.GetExtension(filename));
            }
        }

        private void comboBox_name_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (g_data != null && comboBox_name.SelectedIndex >= 0)
            {
                textBox_toalstar.Text = g_data.Rows[comboxnameindex[comboBox_name.SelectedIndex]]["累计得星"].ToString();
                label1stardate.Text = g_data.Rows[comboxnameindex[comboBox_name.SelectedIndex]]["更新时间"].ToString();
                g_star_totalsum_temp = Convert.ToInt32(g_data.Rows[comboxnameindex[comboBox_name.SelectedIndex]]["累计得星"]);
            }
        }

        private void textBox_toalstar_TextChanged(object sender, EventArgs e)
        {
            if(textBox_toalstar.Text.Length >0)
            {
                g_star_totalsum_temp = Convert.ToInt32(textBox_toalstar.Text);
               // MessageBox.Show(g_star_totalsum_temp.ToString());
            }
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (b_totalstartbutton_flag == false)
            {
                textBox_toalstar.Enabled = true;
                b_totalstartbutton_flag = true;
                button2.Text = "完成";
            }
            else if (b_totalstartbutton_flag == true)
            {
                textBox_toalstar.Enabled = false;
                b_totalstartbutton_flag = false;
                button2.Text = "修改";

                if (g_data != null && textBox_toalstar.Text.Length != 0 )
                {
                    g_data.Rows[comboxnameindex[comboBox_name.SelectedIndex]]["累计得星"] = textBox_toalstar.Text;
                }
            }
           
        }

        public int savepicture(string path)
        {
            string savefilename = null;

            if (path == null)
            {
                MessageBox.Show("保存文件夹不存在！");
                return -1;
            }
            if (!isview)
            {
                MessageBox.Show("图片保存失败未生成图片，请先点击预览");
                return -1;
            }

            if (comboBox_name.Text.Length == 0 && comboBoxtitle.Text.Length == 0)
            {
                MessageBox.Show("保存图片失败，会员姓名或课程主题不能为空");
                return -1;
            }

            if (g_savebmp == null)
            {
                MessageBox.Show("图片保存失败未生成图片，请先点击预览");
                return -1;
            }

            string[] recordtime = g_data.Rows[comboxnameindex[comboBox_name.SelectedIndex]]["更新时间"].ToString().Split('/');
            string temptime = recordtime[0] + "-" + recordtime[1] + "-" + recordtime[2];
            DateTime dateTemp = DateTime.Parse(temptime);
          
            if (DateTime.Compare(dateTemp, monthCalendar_date.SelectionStart.Date) >= 0)
            {
                DialogResult result = MessageBox.Show("当前积星日期小于上次积星记录日期，是否确定保存？ 上次积星记录日期：" +
                    g_data.Rows[comboxnameindex[comboBox_name.SelectedIndex]]["更新时间"].ToString()+ ",当前积星日期：" +
                    monthCalendar_date.SelectionStart.Year.ToString() + "/" +
                    monthCalendar_date.SelectionStart.Month.ToString() + "/" +
                    monthCalendar_date.SelectionStart.Day.ToString(),"提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                if (result == DialogResult.No)
                {
                    if (g_imagback != null)
                    {
                        pictureBox1.Image = MakeThumbnail(g_imagback, pictureBox1.Width, pictureBox1.Height);
                    }
                    isview = false;
                    return -1;
                }
            }

            savefilename = path + "\\" + comboBox_name.Text + comboBoxtitle.Text + g_name_date + ".png";
            if (File.Exists(savefilename))
            {
                DialogResult result = MessageBox.Show(savefilename + "已经存在是否覆盖？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);
                if (result == DialogResult.Yes)
                {
                    g_savebmp.Save(savefilename);
                    MessageBox.Show("图片保存成功！ \n" + savefilename);

                    if (g_imagback != null)
                    {
                        pictureBox1.Image = MakeThumbnail(g_imagback, pictureBox1.Width, pictureBox1.Height);
                    }
                    isview = false;

                    //保存excel数据
                    int ret = savedata();
                    if (ret == -1)
                    {
                        MessageBox.Show("保存失败，数据保存失败！");
                        return -1;
                    }

                    writeLog();//写入日志
                }
                else
                {
                    if (g_imagback != null)
                    {
                        pictureBox1.Image = MakeThumbnail(g_imagback, pictureBox1.Width, pictureBox1.Height);
                    }
                    isview = false;
                    return 0;
                }
            }
            else
            {
                g_savebmp.Save(savefilename);
                MessageBox.Show("图片保存成功！ \n" + savefilename);

                if (g_imagback != null)
                {
                    pictureBox1.Image = MakeThumbnail(g_imagback, pictureBox1.Width, pictureBox1.Height);
                }
                isview = false;

                //保存excel数据
                int ret = savedata();
                if (ret == -1)
                {
                    MessageBox.Show("保存失败，数据保存失败！");
                    return -1;
                }

                writeLog();//写入日志

            }


            //刷线显示页面数据
            textBox_toalstar.Text = g_data.Rows[comboxnameindex[comboBox_name.SelectedIndex]]["累计得星"].ToString();
            label1stardate.Text = g_data.Rows[comboxnameindex[comboBox_name.SelectedIndex]]["更新时间"].ToString();
            textBox_addstar.Text = "0";

            return 0;

        }
        //图片按名字自动保存分类到桌面
        private void buttonsavepicture_Click(object sender, EventArgs e)
        {
      
            string path = string.Empty;

            path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\课后总结照片\\" + username + "\\" + comboBox_name.Text;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

           savepicture(path);
           

            return;

        }


        public Bitmap Splitpiecture(Image img1, Image img2)
        {
            if (img1 == null || img2 == null)
            {
                return null;
            }
            Bitmap bmp = new Bitmap(img1.Width, img1.Height+ img2.Height);
         
            Graphics g = Graphics.FromImage(bmp);

            g.Clear(Color.Transparent);

            img1.RotateFlip(RotateFlipType.Rotate180FlipNone);//旋转180

            g.DrawImage(img1, 0, 0, img1.Width, img1.Height); //g.DrawImage(imgBack, 0, 0, 相框宽, 相框高);

            g.DrawImage(img2, 0, img1.Height, img2.Width, img2.Height);

            GC.Collect();

            return bmp;
        }


        public void getPath(string path)
        {
            DirectoryInfo dir = new DirectoryInfo(path);
            FileInfo[] fil = dir.GetFiles();
            DirectoryInfo[] dii = dir.GetDirectories();
            foreach (FileInfo f in fil)
            {
                list.Add(f.FullName);//添加文件的路径到列表
            }
            //获取子文件夹内的文件列表，递归遍历
            foreach (DirectoryInfo d in dii)
            {
                getPath(d.FullName);
               // list.Add(d.FullName);//添加文件夹的路径到列表
            }
            return ;
        }

        private void buttonsplit_Click(object sender, EventArgs e)
        {

#if true
            string[] filenames = null;
            Bitmap bmp = null;
            Image image1 = null;
            Image image2 = null;
            OpenFileDialog dialog = new OpenFileDialog();
            string defalutpath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\课后总结照片";

            dialog.InitialDirectory = defalutpath;//默认打开桌面

            dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件夹";
            dialog.Filter = "图片|*.png;*.jpg;*.PNG;*.JPG;*.JPEG;*.jpeg";


            if (DialogResult.OK == dialog.ShowDialog())
            {

                filenames = dialog.FileNames;


                if (filenames.Length <= 0)
                {
                    return;
                }
               if (filenames.Length < 2)
               {
                   MessageBox.Show("请至少选择2张图片");
                    return;
                }

                    string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)  ;
                    progressBarsplit.Visible = true;
                    progressBarsplit.Maximum = filenames.Length;
                   
                    /*
                    //排序
                    int i, j; string temp;
                    string strj, strj2;
                    for (i = 0; i < filenames.Length - 1; i++)
                    {
                        for (j = 0; j < filenames.Length - 1 - i; j++)
                        {
                            strj = Path.GetFileNameWithoutExtension(filenames[j]);
                            strj2 = Path.GetFileNameWithoutExtension(filenames[j + 1]);
                            try
                            {
                                strj = strj.Substring(strj.Length - 8, 8);
                                strj2 = strj2.Substring(strj2.Length - 8, 8);
                                if (Convert.ToUInt32(strj) > Convert.ToUInt32(strj2))
                                {
                                    temp = filenames[j];
                                    filenames[j] = filenames[j + 1];
                                    filenames[j + 1] = temp;
                                }
                            }
                            catch {
                                continue;
                            }
                          
                        }
                        
                    }
                    */
                    for (int index = 0; index+1 < filenames.Length; index += 2)
                    {   

                    if (File.Exists(path + "\\" + Path.GetFileNameWithoutExtension(filenames[index]) +
                        Path.GetFileNameWithoutExtension(filenames[index + 1]) + ".png"))
                    {
                        progressBarsplit.Value = index;
                        continue;
                    }
                    image1 = Image.FromFile(filenames[index]);
                        image2 = Image.FromFile(filenames[index+1]);
                        bmp = Splitpiecture(image1, image2);
                        if (bmp != null)
                        {

                            progressBarsplit.Value = index;
                            bmp.Save(path + "\\" + Path.GetFileNameWithoutExtension(filenames[index]) +
                                Path.GetFileNameWithoutExtension(filenames[index + 1]) + ".png");
                        }
                    }

                    progressBarsplit.Visible = false;
                    MessageBox.Show("完成");
                
               
            }
#endif
        }

        private void buttonclose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void buttonmini_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
            //this.notifyIcon1.Visible = true;
        }

        private void panel2_MouseDown(object sender, MouseEventArgs e)
        {

      

            if (e.Button == MouseButtons.Left)
            {
                mouseflag = true;
                FormLocation = this.Location;
                mouseOffset = Control.MousePosition;
            }
        }

        private void panel2_MouseUp(object sender, MouseEventArgs e)
        {
      
                mouseflag = false;
           
        }

        private void panel2_MouseMove(object sender, MouseEventArgs e)
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

        private void button3_Click(object sender, EventArgs e)
        {
            monthCalendar_date.Visible = true;
        }

        private void pictureBox_student_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();

           // dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);//默认打开桌面

            dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件夹";
            dialog.Filter = "图片|*.png;*.jpg;*.PNG;*.JPG;*.JPEG;*.jpeg";

            trackBarbright.Value = 0;//亮度默认为0

            if (DialogResult.OK == dialog.ShowDialog())
            {
                g_image = Image.FromFile(dialog.FileName);

                Bitmap bmp;
                if (g_imagback == null)
                {
                    MessageBox.Show("请先添加课后总结模板！");
                    return;
                }
                if (g_image == null)
                {
                    MessageBox.Show("请先添加图片！");
                    return;
                }

                g_dealImage = g_image;
                
                if (g_dealImage.Width / g_dealImage.Height > pictureBox_student.Width / pictureBox_student.Height)
                {
                   
                    pictureBox_student.Image = MakeThumbnail(g_dealImage, pictureBox_student.Width - 4, pictureBox_student.Height - 4, 0);   
                }else{
                    pictureBox_student.Image = MakeThumbnail(g_dealImage, pictureBox_student.Width - 4, pictureBox_student.Height - 4, 1);          
                }
                 
                if (g_dealImage.Width / g_dealImage.Height > INSERTPICTUREWIDTH / INSERTPICTUREHEIGHT)
                {
                    g_dealImage = MakeThumbnail(g_dealImage, INSERTPICTUREWIDTH, INSERTPICTUREHEIGHT, 0);//以宽为准
                }
                else
                {
                    g_dealImage = MakeThumbnail(g_dealImage, INSERTPICTUREWIDTH, INSERTPICTUREHEIGHT, 1);//以高为准
                }
                g_insertImagenail = g_dealImage;//用于记录初始缩略状态

               // MessageBox.Show(g_dealImage.Width.ToString() + " " + g_dealImage.Height.ToString());
                //g_combineImage = CombinImage(g_imagback,g_image, 1400 + 5, 310 + 5);

                //pictureBox1.Image = MakeThumbnail(bmp, pictureBox1.Width, pictureBox1.Height);




            }
        }

        private void monthCalendar_date_DateChanged(object sender, DateRangeEventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void trackBarbright_Scroll(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            /*
         
            if (e.Button == MouseButtons.Right)
            {

                buttonsavepicture.Visible = true;
                buttonsavepicture.Location = new Point(Control.MousePosition.X - 50, Control.MousePosition.Y - 50);
              
            }
            else
            {
                buttonsavepicture.Visible = false;
            }
             */
        }

        private void buttonsplitall_Click(object sender, EventArgs e)
        {
            ;
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.Description = "请选择文件夹";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (string.IsNullOrEmpty(dialog.SelectedPath))
                {
                    MessageBox.Show(this, "文件夹路径不能为空", "提示");
                    return;
                }
                getPath(dialog.SelectedPath);
                MessageBox.Show(list.Count.ToString());
                for (int index = 0; index < list.Count; index++)
                {
                    textBoxteacher.Text += list[index] + "\r\n";
                }

                if (list.Count < 2)
                {
                    MessageBox.Show("不足2张图片");
                    return;
                }

                string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\课后总结拼接图片";
                progressBarsplit.Visible = true;
                progressBarsplit.Maximum = list.Count;

                if (!Directory.Exists(path))
                {
                    try
                    {
                        Directory.CreateDirectory(path);
                    }
                    catch
                    {
                        MessageBox.Show("创建文件夹失败");
                        path = Path.GetDirectoryName(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));

                    }
                }
                /*
                //排序
                int i, j; string temp;
                string strj, strj2;
                for (i = 0; i < list.Count - 1; i++)
                {
                    for (j = 0; j < list.Count - 1 - i; j++)
                    {
                        strj = Path.GetFileNameWithoutExtension(list[j]);
                        strj2 = Path.GetFileNameWithoutExtension(list[j + 1]);
                        try
                        {
                            strj = strj.Substring(strj.Length - 8, 8);
                            strj2 = strj2.Substring(strj2.Length - 8, 8);
                            if (Convert.ToUInt32(strj) > Convert.ToUInt32(strj2))
                            {
                                temp = list[j];
                                list[j] = list[j + 1];
                                list[j + 1] = temp;
                            }
                        }
                        catch
                        {
                            continue;
                        }

                    }

                }
                */
                Image image1 = null;
                Image image2 = null;
                Bitmap bmp = null;
                for (int index = 0; index + 1 < list.Count; index += 2)
                {

                    if (File.Exists(path + "\\" + Path.GetFileNameWithoutExtension(list[index]) +
                           Path.GetFileNameWithoutExtension(list[index + 1]) + ".png"))
                    {
                        continue;
                    }
                    image1 = Image.FromFile(list[index]);
                    image2 = Image.FromFile(list[index + 1]);
                    bmp = Splitpiecture(image1, image2);
                    if (bmp != null)
                    {
                       
                        progressBarsplit.Value = index;
                        bmp.Save(path + "\\" + Path.GetFileNameWithoutExtension(list[index]) +
                            Path.GetFileNameWithoutExtension(list[index + 1]) + ".png");
                    }  
                }

                progressBarsplit.Visible = false;
                MessageBox.Show("完成");
            }
        }
        private int getLastTotalStarFormData(string name, out int index)
        {
            int totalstar = -1;
            int tempindex = -1;
            bool isfound = false;

            if (g_data == null)
            {
                index = -1;
                return -1;
            }

            for (int i = 0; i < g_data.Rows.Count; i++)
            {
                if (name == g_data.Rows[i]["姓名"].ToString())
                {
                    totalstar = Convert.ToInt32(g_data.Rows[i]["累计得星"].ToString());
                    isfound = true;
                    tempindex = i;
                    break;

                }
            }
            if (isfound == false)
            {
                index = -1;
                return -1;
            }

            index = tempindex;
            return totalstar;
        }
        private string foundclassfor(string titlename)
        {
            string classfor = null;
            int index = -1;

            bool isfound = false;
            for (int i = 0; i < g_data_class_submit.Rows.Count; i++)
            {
                if (titlename == g_data_class_submit.Rows[i]["课程主题"].ToString())
                {
                    index = i;
                    classfor = g_data_class_submit.Rows[i]["课程目标"].ToString();
                    isfound = true;
                    break;
                }
            }
            if (isfound == false)
            {
                return null;
            }


            return  classfor;
        }
      
        private void buttonmultipic_Click(object sender, EventArgs e)
        {
            batchmakePicture();
        }

        public void batchmakePicture()
        {

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);//默认打开桌面

            dialog.Multiselect = false;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件";
            dialog.Filter = "Excel文件|*.xls;*xlsx";

            if (DialogResult.OK != dialog.ShowDialog())
            {
                return;
            }

          

            if (g_imagback == null)
            {
                MessageBox.Show("请先添加背景！");
                return;
            }

            g_autoData = ExcelToDataTable(dialog.FileName, null, true);
            if (g_autoData == null)
            {
                MessageBox.Show(dialog.FileName + "加载失败！");
                return;
            }


            List<string> logmsgclasstitle = new List<string>();
            List<string> logmsgname = new List<string>();
            List<string> logmsgthisstarnum = new List<string>();
            List<string> logmsgtotalstarnum = new List<string>();
            List<string> logmsgdate = new List<string>();

            string[] classfors = new string[g_autoData.Rows.Count];
            //int[] lastTotalstars = new int[g_autoData.Rows.Count];
            //int[] namegdataindex = new int[g_autoData.Rows.Count];

            progressBarsplit.Visible = true;
            progressBarsplit.Maximum = g_autoData.Rows.Count;
            progressBarsplit.Value = 1;

            for (int i = 0; i < g_autoData.Rows.Count; i++)
            {
                if (!File.Exists(Path.GetDirectoryName(dialog.FileName) + "\\" + g_autoData.Rows[i]["照片名称"].ToString()))
                {
                    MessageBox.Show(Path.GetDirectoryName(dialog.FileName) + "\\" + g_autoData.Rows[i]["照片名称"].ToString() + "不存在");
                    progressBarsplit.Visible = false;
                    return;
                }

                string classfor = foundclassfor(g_autoData.Rows[i]["课程主题"].ToString());
                if (classfor != null)
                {
                    classfor = classfor.Replace("。", "。\r\n");//在每个句号后面加换行
                    classfors[i] = classfor;
                }
                else
                {
                    MessageBox.Show("未找到课程主题" + g_autoData.Rows[i]["课程主题"].ToString());
                    progressBarsplit.Visible = false;
                    return;
                }

                if (g_autoData.Rows[i]["年"].ToString().Length == 0 || g_autoData.Rows[i]["月"].ToString().Length == 0 ||
                    g_autoData.Rows[i]["日"].ToString().Length == 0 || g_autoData.Rows[i]["课堂出勤"].ToString().Length == 0 ||
                    g_autoData.Rows[i]["课堂纪律"].ToString().Length == 0 || g_autoData.Rows[i]["知识掌握"].ToString().Length == 0 ||
                    g_autoData.Rows[i]["器件整理"].ToString().Length == 0 || g_autoData.Rows[i]["加扣星数"].ToString().Length == 0)
                {
                    MessageBox.Show("第"+(i+1).ToString()+"行数据有误，请检查日期或星星数是否为空");
                    progressBarsplit.Visible = false;
                    return;
                }

            }


            for (int i = 0; i < g_autoData.Rows.Count; i++)
            {
                g_image = null;
                string classtitle = g_autoData.Rows[i]["课程主题"].ToString();
                string pictureYear = g_autoData.Rows[i]["年"].ToString();
                string pictureMonth = g_autoData.Rows[i]["月"].ToString();
                string pictureDay = g_autoData.Rows[i]["日"].ToString();
                string name = g_autoData.Rows[i]["会员姓名"].ToString();
                string picturfilename = g_autoData.Rows[i]["照片名称"].ToString();
                string teacherSummit = g_autoData.Rows[i]["课后评价"].ToString();
                int starNum_1 = Convert.ToInt32(g_autoData.Rows[i]["课堂出勤"].ToString());
                int starNum_2 = Convert.ToInt32(g_autoData.Rows[i]["课堂纪律"].ToString());
                int starNum_3 = Convert.ToInt32(g_autoData.Rows[i]["知识掌握"].ToString());
                int starNum_4 = Convert.ToInt32(g_autoData.Rows[i]["器件整理"].ToString());
                int starNum_add = Convert.ToInt32(g_autoData.Rows[i]["加扣星数"].ToString());

           

                int namegdataindex;
                int lasttotalstar = getLastTotalStarFormData(g_autoData.Rows[i]["会员姓名"].ToString(), out namegdataindex);
                if (lasttotalstar == -1)
                {
                    MessageBox.Show("未找到" + g_autoData.Rows[i]["会员姓名"].ToString());
                    progressBarsplit.Visible = false;
                    return;
                }
            


                textBoxteacher.Text += (classtitle + pictureYear + pictureMonth + pictureDay + 
                    name + picturfilename + teacherSummit + starNum_1.ToString() + starNum_2.ToString() + 
                    starNum_3.ToString()+ starNum_4.ToString() + starNum_add.ToString()+ classfors[i]+ lasttotalstar.ToString());


                try { g_image = Image.FromFile(Path.GetDirectoryName(dialog.FileName) + "\\" + picturfilename); }
                catch (Exception E)
                {
                    MessageBox.Show(E.Message + picturfilename);
                    return;
                }


                Bitmap bmp = new Bitmap(g_imagback.Width, g_imagback.Height);

                using (Graphics g = Graphics.FromImage(bmp))
                {

                    g.Clear(Color.Transparent);

                    //画背景
                    g.DrawImage(g_imagback, 0, 0, g_imagback.Width, g_imagback.Height); //g.DrawImage(imgBack, 0, 0, 相框宽, 相框高);

                    //要插图片的起始位置和可以显示的范围
                    int insertPictureRangeWidth = 928 - 10;
                    int insertPictureRangeHeight = 1236 - 10;
                    int insertPictureStart_x = 1400 + 5;
                    int insertPictureStart_y = 310 + 5;

                    //插入图片的起始位置和大小
                    int start_x = -1;
                    int start_y = -1;
                    int towidth = -1;
                    int toheight = -1;

                    //根据图片的实际大小计算插入的位置和宽高
                    if (g_image.Width / g_image.Height > insertPictureRangeWidth / insertPictureRangeHeight)
                    {
                        towidth = insertPictureRangeWidth;
                        toheight = (insertPictureRangeWidth * g_image.Height) / g_image.Width;

                        start_x = insertPictureStart_x + (insertPictureRangeWidth - towidth) / 2;
                        start_y = insertPictureStart_y + (insertPictureRangeHeight - toheight) / 2;

                    }
                    else
                    {
                        toheight = insertPictureRangeHeight;
                        towidth = (g_image.Width * toheight) / g_image.Height;
                        start_x = insertPictureStart_x + (insertPictureRangeWidth - towidth) / 2;
                        start_y = insertPictureStart_y + (insertPictureRangeHeight - toheight) / 2;
                    }

                    //要插入的图片
                    g.DrawImage(g_image, start_x, start_y, towidth, toheight);

                    List<int> starnum = new List<int>();
                    starnum.Add(starNum_1);
                    starnum.Add(starNum_2);
                    starnum.Add(starNum_3);
                    starnum.Add(starNum_4);

                    List<Bitmap> star = new List<Bitmap>();
                    star.Add(Properties.Resources.star1);
                    star.Add(Properties.Resources.start2);
                    star.Add(Properties.Resources.star3);
                    star.Add(Properties.Resources.start4);
                    star.Add(Properties.Resources.start5);

                    int StarWidth = 70;
                    int StarHeight = 70;
                    int StarSpace = 30;
                    int LineSpace = 10;

                    int starStart_X = 400;
                    int starStart_Y = 1200;

                    for (int j = 0; j < 4; j++)
                    {
                        for (int k = 0; k < starnum[j]; k++)
                        {
                            g.DrawImage(star[k], starStart_X + (StarWidth + StarSpace) * k, starStart_Y + (StarHeight + LineSpace) * j, StarWidth, StarHeight);
                        }
                    }

                    Markword_2(g, "课堂出勤：", 160, starStart_Y + (80 * 0) + 10, "微软雅黑", 35, Color.Black);
                    Markword_2(g, "课堂纪律：", 160, starStart_Y + (80 * 1) + 10, "微软雅黑", 35, Color.Black);
                    Markword_2(g, "知识掌握：", 160, starStart_Y + (80 * 2) + 10, "微软雅黑", 35, Color.Black);
                    Markword_2(g, "器件整理：", 160, starStart_Y + (80 * 3) + 10, "微软雅黑", 35, Color.Black);


                    starStart_Y += ((80 * 4) + 20);

                    int thisStarNum = starNum_1 + starNum_2 + starNum_3 + starNum_4+ starNum_add;

                    Markword_2(g, thisStarNum.ToString(), starStart_X+10, starStart_Y, "楷体", 40, Color.Red);
           

                    int totalstar = lasttotalstar + thisStarNum;

                    Markword_2(g, totalstar.ToString(), starStart_X+10, starStart_Y + 80 , "楷体", 40, Color.Red);// Markword_2(g,);

                   

                    Markword_2(g, "本次得星：", 160, starStart_Y+5, "楷体", 35, Color.DeepSkyBlue);
                    Markword_2(g, "累计得星：", 160, starStart_Y +80 + 5, "楷体", 35, Color.DeepPink);


                    Markword_2(g, "课程主题：" + classtitle + "\r\n" + "课程目标：\r\n" + classfors[i], 160, 340, "微软雅黑", 35, Color.Black, 1, 1100, 20);///课程目标

                    if (g_autoData.Rows[i]["课后评价"].ToString().Length != 0)
                    {
                       Markword_2(g, "课后评价：\r\n" + teacherSummit, 160, 870, "微软雅黑", 31, Color.Black, 1, 1100, 18);//课后评价
                    }

                    Markword_2(g,name, 2020, 1560, "微软雅黑", 31, Color.Red);//名字            
                    Markword_2(g, pictureYear + "年" + pictureMonth + "月" + pictureDay + "日",
                        2020, 1610, "微软雅黑", 31, Color.Red);//合成日期


                    g_data.Rows[namegdataindex]["累计得星"] = totalstar.ToString();//更新数据到g_data  
                    g_data.Rows[namegdataindex]["更新时间"] = pictureYear + "/" + pictureMonth + "/" + pictureDay;//更新数据到g_data  



                    //日志信息
                    logmsgclasstitle.Add(classtitle);
                    logmsgname.Add(name);
                    logmsgthisstarnum.Add(thisStarNum.ToString());
                    logmsgtotalstarnum.Add(totalstar.ToString());
                    logmsgdate.Add(pictureYear + pictureMonth + pictureDay);


                }

                GC.Collect();
                if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "//批量生成照片"))
                {
                    Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "//批量生成照片");
                }

                bmp.Save(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + 
                    "//批量生成照片//"+classtitle+name+pictureYear.ToString() + pictureMonth.ToString() + pictureDay.ToString() + ".png");

                progressBarsplit.Value = i + 1;

               
            }


            int ret = DataTableToExcel(Application.StartupPath + "\\" + STUDENTMEMER, g_data, "会员信息", true);
            if (ret == -1)
            {
               string errlogPath = Environment.CurrentDirectory + "\\Log";
            if (!Directory.Exists(errlogPath))
            {
                Directory.CreateDirectory(errlogPath);
            }
            string errFilePath = Path.Combine(errlogPath, string.Format("errorlog{0}.log", DateTime.Now.Date.ToString("yyyyMMdd")));
                using (StreamWriter sw = new StreamWriter(errFilePath, true))
                {
                    sw.WriteLine("错误日志：");
                    sw.WriteLine(username + " " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  批量生成");
                    for (int i = 0; i < logmsgname.Count; i++)
                    {
                        string detailMsg = logmsgclasstitle[i] + " " + logmsgname[i] + " " + logmsgthisstarnum[i] + " " + logmsgtotalstarnum[i] + " " + logmsgdate[i];

                        sw.WriteLine(detailMsg);
                        sw.WriteLine();
                    }

                    sw.Close();
                }
                MessageBox.Show("数据保存失败！请重新生成，查看错误日志"+ errFilePath);

                progressBarsplit.Visible = false;
                return;
            }



            //写入Log信息
            string logPath = Environment.CurrentDirectory + "\\Log";
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }
            string logFilePath = Path.Combine(logPath, string.Format("{0}.log", DateTime.Now.Date.ToString("yyyyMMdd")));

            string usr = username + " ";
            string strMsg = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")+"  批量生成";
            
            using (StreamWriter sw = new StreamWriter(logFilePath, true))
            {
                sw.WriteLine(usr + strMsg);
                for (int i = 0; i < logmsgname.Count; i++)
                {
                    string detailMsg = logmsgclasstitle[i] + " " + logmsgname[i] + " " + logmsgthisstarnum[i] + " " + logmsgtotalstarnum[i]+" "+logmsgdate[i];
                  
                    sw.WriteLine(detailMsg);
                    sw.WriteLine();
                }
               
                sw.Close();
            }

           
            progressBarsplit.Visible = false;

            MessageBox.Show("完成");

        }
        private void pictureBox_student_MouseMove(object sender, MouseEventArgs e)
        {
            pictureBox_student.BackColor = Color.LimeGreen;
        }

        private void pictureBox_student_MouseLeave(object sender, EventArgs e)
        {
            pictureBox_student.BackColor = Color.Transparent;
        }

        private void pictureBox_student_MouseDown(object sender, MouseEventArgs e)
        {
            pictureBox_student.BackColor = Color.Green;
           
        }

        public void writeLog()
        {
            string logPath = Environment.CurrentDirectory + "\\Log";
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }

            string logFilePath = Path.Combine(logPath, string.Format("{0}.log", DateTime.Now.Date.ToString("yyyyMMdd")));
           
            string usr = username +" ";
            string strMsg = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string detailMsg = comboBoxtitle.Text + " " + comboBox_name.Text + " " + g_thisStarNum.ToString()+" "+textBox_addstar.Text+" "+
                g_data.Rows[comboxnameindex[comboBox_name.SelectedIndex]]["累计得星"].ToString() + " " +
                 g_data.Rows[comboxnameindex[comboBox_name.SelectedIndex]]["更新时间"].ToString();

            using (StreamWriter sw = new StreamWriter(logFilePath, true))
            {
                sw.WriteLine(usr+strMsg);
                sw.WriteLine(detailMsg);    
                sw.WriteLine();
                sw.Close();
            }
        }
    }
}
