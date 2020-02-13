using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using ress;

namespace Mail
{
    public partial class Mail : Form
    {
        public Mail()
        {
            InitializeComponent();
        }

        public class Cargo
        {
            public string file { get; set; }
            public string fam { get; set; }
            public string im { get; set; }
            public string ot { get; set; }
            public string dr { get; set; }
            public string mo { get; set; }
            public string resp { get; set; }
            public string rn { get; set; }
            public string np { get; set; }
            public string ul { get; set; }
            public string dom { get; set; }
            public string kv { get; set; }
            public string korp { get; set; }
            public string indx { get; set; }

            public void piece(string line)
            {
                string[] parts = line.Split(';');  //Разделитель в CVS файле.
                file = parts[1];
                fam = parts[2];
                im = parts[3];
                ot = parts[4];
                dr = parts[5];
                mo = parts[6];
                resp = parts[7];
                //rn = parts[6];
                //np = parts[7];
                //ul = parts[8];
                //dom = parts[9];
                //kv = parts[10];
                //korp = parts[11];
                indx = parts[8];
            }

            public static List<Cargo> ReadFile(string filename)
            {
                List<Cargo> res = new List<Cargo>();
                using (StreamReader sr = new StreamReader(filename))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        Cargo p = new Cargo();
                        p.piece(line);
                        res.Add(p);
                    }
                }

                return res;
            }
        }
        private void saveSnils_Click(object sender, EventArgs e)
        {
            List<Cargo> CSV_Struct = new List<Cargo>();
            CSV_Struct = Cargo.ReadFile("test.csv");
            label1.Text = "Всего строк: " + Convert.ToString(CSV_Struct.Count);
            for (int i = 0; i <= CSV_Struct.Count - 1; i++)
            {
                Image image = Image.FromFile(CSV_Struct[i].file +".jpg"); //ress.Properties.Resources.img;
                using (Graphics g = Graphics.FromImage(image))
                {
                    Font fontind = new Font("Arial", 8, FontStyle.Italic);
                    SolidBrush Brush3 = new SolidBrush(Color.DarkGray);
                    g.DrawString((i + 1).ToString(), fontind, Brush3, new Point(10, 10));
                    //g.DrawString("Росгосстрах-Бурятия-Медицина", font, Brush2, new Point(159, 1168));
                    //g.DrawString("Бурятия, Улан-Удэ", font, Brush2, new Point(159, 1201));
                    //g.DrawString("Маяковского, 1А", font, Brush2, new Point(159, 1234));
                    RectangleF header3Rect = new RectangleF();
                    header3Rect.Width = 500;
                    header3Rect.Location = new Point(695, 1430);
                    Font font = new Font("Arial", 25, FontStyle.Italic);
                    SolidBrush Brush2 = new SolidBrush(Color.Black);
                    g.DrawString(CSV_Struct[i].im + " " + CSV_Struct[i].ot, font, Brush2, new Point(388, 375));
                    //g.DrawString("Росгосстрах-Бурятия-Медицина", font, Brush2, new Point(159, 1168));
                    //g.DrawString("Бурятия, Улан-Удэ", font, Brush2, new Point(159, 1201));
                    //g.DrawString("Маяковского, 1А", font, Brush2, new Point(159, 1234));
                    RectangleF header2Rect = new RectangleF();
                    header2Rect.Width = 500;
                    header2Rect.Location = new Point(695, 1430);
                    g.DrawString(CSV_Struct[i].fam + " " + CSV_Struct[i].im + " " + CSV_Struct[i].ot, font, Brush2, header2Rect);
                    //g.DrawString(CSV_Struct[i].im + " " + CSV_Struct[i].ot, font, Brush2, new Point(695, 1465));
                    RectangleF address2Rect = new RectangleF();
                    address2Rect.Width = 500;
                    address2Rect.Location = new Point(684, 1507);
                    g.DrawString(CSV_Struct[i].resp.Replace(',',' '), font, Brush2, address2Rect);
                    //g.DrawString(CSV_Struct[i].np, font, Brush2, new Point(643, 1542));
                    //g.DrawString(CSV_Struct[i].ul + ", " + CSV_Struct[i].dom, font, Brush2, new Point(643, 1579));
                    //g.DrawString(CSV_Struct[i].kv + ", " + CSV_Struct[i].korp, font, Brush2, new Point(831, 1622));
                    Font font1 = new Font("Arial", 33, FontStyle.Regular);
                    g.DrawString(CSV_Struct[i].indx, font1, Brush2, new Point(666, 1659));
                    Font fontmo = new Font("Arial", 24, FontStyle.Regular);
                    g.DrawString("Ваше прикрепление: "+CSV_Struct[i].mo.ToString(), fontmo, Brush2, new Point(110, 1130));
                    //g.DrawString("Росгосстрах-Бурятия-Медицина", font, Brush2, new Point(159, 1168));
                    //g.DrawString("Бурятия, Улан-Удэ", font, Brush2, new Point(159, 1201));
                    //g.DrawString("Маяковского, 1А", font, Brush2, new Point(159, 1234));
                    g.Dispose();
                    font.Dispose();
                    font1.Dispose();
                    Brush2.Dispose();
                }
                image.Save("img/" + (i + 1) + " " + CSV_Struct[i].file + " " + CSV_Struct[i].im + " " + CSV_Struct[i].ot + ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                image.Dispose();
            }
        }
     }
}
