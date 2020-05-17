using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace TextReadactor
{
    public partial class ChartStaticstic : Form
    {
        public static int Numbers, LatinLtr, KirilLtr;
        public ChartStaticstic()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button2.Enabled = false;
            Readactor readactor = new Readactor();
            readactor.Excel_Create(Program.Redact.Name, LatinLtr, KirilLtr, Numbers);
            button2.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        public void ChartStaticstic_Load(object sender, EventArgs e)
        {
            Text = "Символьная статистика " + Program.Redact.Text; 
            LatinLtr = 0;
            KirilLtr = 0;
            Numbers = 0;
            foreach (char ch in Program.RedactTB.Text)
            {
                if ((ch>='a')&&(ch<='z')||(ch>='A')&&(ch<='Z'))
                {
                    LatinLtr++;
                }
                if ((ch >= 'а') && (ch <= 'я') || (ch >= 'А') && (ch <= 'Я'))
                {
                    KirilLtr++;
                }
                if ((ch >= '0') && (ch <= '9'))
                {
                    Numbers++;
                }
            }
            chart1.Series["SymbStatic"].Points.AddXY("Латинские буквы "+LatinLtr.ToString(),LatinLtr);
            chart1.Series["SymbStatic"].Points.AddXY("Русские буквы "+KirilLtr.ToString(),KirilLtr);
            chart1.Series["SymbStatic"].Points.AddXY("Числа "+Numbers.ToString(),Numbers);
        }
    }
}
