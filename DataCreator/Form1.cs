using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DataCreator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {            
            if (comboBox3.SelectedIndex == 0) //Veritabanı DB2 seçilmişse
            {

                if (radioButton1.Checked == true)//Dev
                {
                    TabloBilgi();
                    DataOlustur();
                    DosyaOlustur();
                }

                else if (radioButton2.Checked == true)//Test
                {
                    TabloBilgi();
                    DataOlustur();
                    DosyaOlustur();
                }
            }


            if (comboBox3.SelectedIndex == 1) //Veritabanı SQL seçilmişse
            {

                if (radioButton1.Checked == true)//Dev
                {
                    TabloBilgi();
                    DataOlustur();
                    DosyaOlustur();
                }

                else if (radioButton2.Checked == true)//Test
                {
                    TabloBilgi();
                    DataOlustur();
                    DosyaOlustur();

                }
            }
        }


        public void TabloBilgi()//Tablo metadatasını çeker
        {

            //Metadata bir dizide tutulacak


            /*
---BU KOD DATA LENGTH HARİÇ DİĞER VERİLERİ ÇEKİYOR DATA LENGTH İ NASIL ALACAZ???

SELECT  object_name(c.id)    AS table_name, 
        c.name               AS column_name,
        t.name               AS data_type
FROM  syscolumns AS c 
INNER JOIN systypes   AS t  ON c.xtype = t.xtype
WHERE c.id = object_id( 'TabloAdı' )



*/



        }


        public void DataOlustur()//Random data oluşturur
        {

            //Random data diziye basılacak
        }


        public void DosyaOlustur()//Excel ya da txt formatında dosya oluşturur
        {
                if (comboBox1.SelectedIndex == 0)
                {
                    Excel.Application xlApp = new Excel.Application();
                    xlApp.Visible = true;
                    xlApp.DisplayAlerts = true;

                    Excel.Workbook wb = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    Excel.Worksheet ws = (Excel.Worksheet)wb.Sheets[1];


                    //örnek veri
                    ws.Cells[1, 1] = "Merhaba - Hello";
                }
                else if (comboBox1.SelectedIndex == 1)
                {
                    //txt dosya çıkılacak
                }
            }
   

        }
    }

