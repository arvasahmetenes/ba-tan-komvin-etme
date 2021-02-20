using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;
using System.Windows.Forms.DataVisualization.Charting;
using offis = Microsoft.Office.Interop.Excel;// başvurular(referances) kısmına sağ tıklayıp excell i kaydediyoruz.

namespace baştan_komvin_etme
{
    public partial class Form1 : Form
    {
        string gelen = "0";

        DateTime yeni = DateTime.Now;
        int zaman = 0;
        int satir = 1;      //değişkenimizi satır no 1 den başlattık çünkü satır 0 da başlık yazıları mevcut.(saat tarih vb.)
        int sutun = 1;    //değişkenimiz 0 dan değil de 1 den başlamıştır sebebi malum
        int satirNo = 1;
        int k = 0;   //* string gelen den k = 0  a kadar olan kısım excell için olan kodlardır 
        string[] portlar = SerialPort.GetPortNames();
        string sonuc;
        long maks = 30, min = 0, i = 0;//string portlar dan long maks a kadar olan kısım grafik için olan kodlardır.

        public Form1()
        {
            InitializeComponent();
            serialPort1.PortName = "COM4";
            serialPort1.BaudRate = 9600;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            serialPort1.Close();
            timer1.Stop();
            //serialPort1.DiscardInBuffer();
            button1.Enabled = true;
            if (serialPort1.IsOpen == true)
            {
                serialPort1.Close();
                label2.Text = "Bağlantı Kapalı";
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application objExcel = new Microsoft.Office.Interop.Excel.Application();//excell adında bi uygulama oluşturduk,ve bunu yeniden tanımladık (new den sonra)
            objExcel.Visible = true;//(excelli görünür hale getirdik.)
            Microsoft.Office.Interop.Excel.Workbook objbook = objExcel.Workbooks.Add(System.Reflection.Missing.Value);//bir çalışma kitabı tanımlıyoruz.
            Microsoft.Office.Interop.Excel.Worksheet objSheet = (Microsoft.Office.Interop.Excel.Worksheet)objbook.Worksheets.get_Item(1);//çalışma sayfası ayarlıyoruz.

            for (int s = 0; s < dataGridView1.Columns.Count; s++)//datagridwievde kolomn lardan elde ettiğimiz bilgileri aşağıda microsoft office de seçtiğimiz alana ekliyoruz (range)
            {
                Microsoft.Office.Interop.Excel.Range myrange = (Microsoft.Office.Interop.Excel.Range)objSheet.Cells[1, s + 1];
                myrange.Value2 = dataGridView1.Columns[s].HeaderText;

            }
            for (int s = 0; s < dataGridView1.Columns.Count; s++)//satır ve sutunları eklemek için kullanılan döngü.
            {
                for (int j = 0; j < dataGridView1.Rows.Count; j++)
                {
                    Microsoft.Office.Interop.Excel.Range myrange = (Microsoft.Office.Interop.Excel.Range)objSheet.Cells[j + 2, s + 1];
                    myrange.Value2 = dataGridView1[s, j].Value;
                }
            }
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            gelen = serialPort1.ReadLine();
            string[] pot = gelen.Split('*');
            label1.Text = gelen + "";
            // this.chart1.Series[0].Points.AddXY(zaman, gelen);
            zaman = (zaman + 1);

            satir = dataGridView1.Rows.Add();  //timer da satır diyerek yeni bir satır ekledik

            dataGridView1.Rows[satir].Cells[0].Value = satirNo;  //1. satırın 0. hücresine satırNo yu ekledik
            //dataGridView1.Rows[satir].Cells[1].Value = gelen;    //1. satırın 1. hücresine gelen veriyi ekledik yani seri porttan okunan dağer.
            //dataGridView1.Rows[satir].Cells[1].Value = pot[0];
            for (int i = 0; i < pot.Length; i++)
            {
                dataGridView1.Rows[satir].Cells[i + 1].Value = pot[i];                    //YENİ DENEME


            }

            dataGridView1.Rows[satir].Cells[2].Value = yeni.ToLongTimeString();  // 1. satırın 2. hücresine saati  ekledik
            dataGridView1.Rows[satir].Cells[3].Value = yeni.ToShortDateString(); //1. satırın 3. hücresine tarihi ekledik
            satir++;  //sonra da satırı ve satır noyu birebirer artırdık.
            satirNo++;

            label1.Text = gelen;  //gelen veriyi labelde gösterdik.
                                  //timer 1 başlangıcından buraya kadar olan kısım excell kodlarıdır.try dan sonrası ise grafik kodlarıdır.

            string sonuc = serialPort1.ReadLine();//burada readexisting yapsaydık seri portta aşağıdaki komutları da okurdu o needenle readLine yaptık
            pot = sonuc.Split('*');//**yukarıda pot değerini yerel değişken olarak tanımlamıştık bu satırı şimdilk kapattım.(deneyip bakacam)
            label1.Text = sonuc + "";

            textBox2.Text = pot[0];
            if (pot.Length>=2)
            {
                textBox3.Text = pot[1];

            }
            if (pot.Length>=3)
            {
                textBox4.Text = pot[2];

            }
            if (pot.Length>=4)
            {
                textBox5.Text = pot[3];

            }

            serialPort1.DiscardInBuffer();


            timer1.Stop();
            //throw;

            //   chart1.ChartAreas[0].AxisX.Minimum = min;
            //   chart1.ChartAreas[0].AxisX.Maximum = maks;

            //   chart1.ChartAreas[0].AxisY.Minimum = 0;
            //   chart1.ChartAreas[0].AxisY.Maximum = 300;

            //   chart1.ChartAreas[0].AxisX.ScaleView.Zoom(min, maks);


            sonuc = textBox2.Text;//önceden burada serial portu oku demiştik.
            if (sonuc != null)
            {
                label3.Text = sonuc + "";
                label3.Text = textBox2.Text;
                // label3.Text = sonuc + "";
                this.chart1.Series[0].Points.AddXY(i, sonuc);
                //   maks++;
                //   min++;
            }
            serialPort1.DiscardInBuffer();

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            serialPort1.Close();
            this.chart1.Titles.Add("mesafe ölçüm");
            DateTime yeni = DateTime.Now;//serialPort1 close dan bu satıra kadar olan kısım excell denemesinin
            foreach (string port in portlar)
            {
                comboBox1.Items.Add(port);
                comboBox1.SelectedIndex = 0;
            }
            comboBox2.Items.Add("4800");
            comboBox2.Items.Add("9600");
            comboBox2.SelectedIndex = 1;
            label2.Text = "Bağlantı Kapalı";//foreach string port in portlar dan bu satıra kadar olan kısım grafik denemesinin



        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (serialPort1.IsOpen == true)
                serialPort1.Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            System.Threading.Thread.Sleep(3000);
            serialPort1.Open();
            timer1.Start();
            button1.Enabled = false;
            if (serialPort1.IsOpen == false)
            {
                if (comboBox1.Text == "")
                    return;
                serialPort1.PortName = comboBox1.Text;
                serialPort1.BaudRate = Convert.ToInt16(comboBox2.Text);
                try
                {
                    serialPort1.Open();
                    label2.Text = "Bağlantı Açık";

                }
                catch (Exception hata)
                {
                    MessageBox.Show("hata : " + hata.Message);
                }
            }
            else
            { label2.Text = "Bağlantı Kuruldu"; }
        }
    }
}
