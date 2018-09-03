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
    public partial class Form2 : Form
    {

        Image g_imagback = null;
        Color wordColor = Color.Black;
        Bitmap savebmp = null;

        public Form2()
        {
            InitializeComponent();
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

            Graphics g = Graphics.FromImage(bmp);

            g.Clear(Color.Transparent);
            g.DrawImage(img, 0, 0, towidth, toheight); //g.DrawImage(imgBack, 0, 0, 相框宽, 相框高);

            GC.Collect();

            return bmp;
        }

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
                else if (rowEndIndex + 1 < text.Length  &&  text.Substring(rowEndIndex, 2) == "\r\n" )
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


        public Bitmap Markword(Image img, string str, int x, int y, string font, int size, Color wordcolor, int mode = 0, int width = 0, int rowsapce = 0)
        {
            int row_y = y;
           // MessageBox.Show(str+ " "+ x.ToString()+" "+y.ToString()+" "+ font + " "+size.ToString() + " "+width.ToString() + " "+rowsapce.ToString());

            Bitmap bmp = new Bitmap(img.Width, img.Height);

            Graphics g = Graphics.FromImage(bmp);

            g.Clear(Color.Transparent);

            g.DrawImage(img, 0, 0, img.Width, img.Height);

            SolidBrush sbrush = new SolidBrush(wordcolor);

            if (mode == 0)
            {
                g.DrawString(str, new Font(font, size), sbrush, new PointF(x, y));
            }
            else if (mode == 1)
            {
                List<string> textRows = new List<string>();
                
                textRows = GetStringRows(g, new Font(font, size), str, width);
                for (int i = 0; i < textRows.Count; i++)
                {
                 
                    //g.DrawString(textRows[i], new Font(font, size), sbrush, new PointF(x, row_y + (i * (new Font(font, size).Height + rowsapce))));

                    g.DrawString(textRows[i], new Font(font, size), sbrush, new PointF(x, row_y ));
                    row_y += (new Font(font, size).Height + rowsapce);

                   // MessageBox.Show(row_y.ToString()+" "+textRows[i]);
                }
            }

            GC.Collect();

            return bmp;
        }
        private void buttonbackimg_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件夹";
            dialog.Filter = "图片|*.png;*.jpg";


            if (DialogResult.OK == dialog.ShowDialog())
            {
                g_imagback = Image.FromFile(dialog.FileName);

              
                /*
                if ((g_imagback.Width != 2480) || (g_imagback.Height != 1748))
                {
                    MessageBox.Show("图片尺寸有误（2480*1748）：" + g_imagback.Width.ToString() + " " + g_imagback.Height.ToString());
                    return;
                }

        */
                labelpicture_x.Text = g_imagback.Width.ToString();
                labelpicture_y.Text = g_imagback.Height.ToString();

                pictureBox1.Image = MakeThumbnail(g_imagback, pictureBox1.Width, pictureBox1.Height);
            }
        }

        private void buttonview_Click(object sender, EventArgs e)
        {
            Bitmap bmp;
            string text;

            //bmp = Markword(g_imagback, textBoxword.Text, 154, 435, "微软雅黑", 31, Color.Red, 1, 1035, 10);

            if (g_imagback == null)
            {
                MessageBox.Show("请先添加背景图片");
                return;
            }

            text = labelclasstitle.Text + textBoxclasstitle.Text + "\r\n" + labelclassfor.Text + "\r\n" + textBoxclassfor.Text;
            bmp = Markword(g_imagback, text, Convert.ToInt32(textBoxstarx.Text), Convert.ToInt32(textBoxstarty.Text), "微软雅黑", Convert.ToInt32(textBoxfontsize.Text), wordColor, 1, Convert.ToInt32(textBoxwordwidth.Text), Convert.ToInt32(textBoxrowsapce.Text));

            savebmp = bmp;

            pictureBox1.Image = MakeThumbnail(bmp, pictureBox1.Width, pictureBox1.Height);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ColorDialog ColorForm = new ColorDialog();
            if (ColorForm.ShowDialog() == DialogResult.OK)
            {
                Color GetColor = ColorForm.Color;
                //GetColor就是用户选择的颜色，接下来就可以使用该颜色了
                wordColor = GetColor;
            }
        }

        private void buttonsave_Click(object sender, EventArgs e)
        {

            if (savebmp == null)
            {
                MessageBox.Show("图片保存失败未生成图片，请先点击预览");
                return;
            }


            string path = string.Empty;
            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                path = fbd.SelectedPath;
               
                string savefilename = path + "\\" +textBoxclasstitle.Text +"课后总结模板"+ ".png";

               savebmp.Save(savefilename);

                MessageBox.Show("图片保存成功！ \n" + savefilename);
            }
            return;

        }
    }
}
