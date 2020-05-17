using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace TextReadactor
{
    public partial class InterfaceOption : Form
    {
        public Color MainCLR = new Color();
        public Color ContainerCLR = new Color();
        public Color ManupCLR = new Color();
        public Color TextCLR = new Color();
        public InterfaceOption()
        {
            InitializeComponent();
        }

        public void button3_Click(object sender, EventArgs e)
        {
            ((MainForm)Owner).BackColor = MainCLR;
            ((MainForm)Owner).ForeColor = TextCLR;
            ((MainForm)Owner).menuStrip1.BackColor = ContainerCLR;
            ((MainForm)Owner).menuStrip1.ForeColor = TextCLR;
            ((MainForm)Owner).statusStrip1.BackColor = ContainerCLR;
            ((MainForm)Owner).statusStrip1.ForeColor = TextCLR;
            ((MainForm)Owner).toolStrip1.BackColor = ManupCLR;
            ((MainForm)Owner).toolStrip1.ForeColor = TextCLR;
            foreach (ToolStripMenuItem mi 
                in ((MainForm)Owner).menuStrip1.Items.
                OfType<ToolStripMenuItem>())
            {
                mi.BackColor = ContainerCLR;
                mi.ForeColor = TextCLR;
                foreach (ToolStripItem ddi in 
                    mi.DropDownItems.OfType<ToolStripItem>())
                {
                    ddi.BackColor = ContainerCLR;
                    ddi.ForeColor = TextCLR;
                }
                foreach (ToolStripSeparator ssi 
                    in mi.DropDownItems.OfType<ToolStripSeparator>())
                {
                    ssi.BackColor = ContainerCLR;
                    ssi.ForeColor = TextCLR;
                }
            }
            foreach (ToolStripComboBox micb 
                in ((MainForm)Owner).toolStrip1.Items.
                OfType<ToolStripComboBox>())
            {
                micb.BackColor = ManupCLR;
                micb.ForeColor = TextCLR;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox1.SelectedIndex)
            {
                case (0):
                    MainCLR = panel2.BackColor;
                    ContainerCLR = panel6.BackColor;
                    ManupCLR = panel10.BackColor;
                    TextCLR = label1.ForeColor;
                    break;
                case (1):
                    MainCLR = panel3.BackColor;
                    ContainerCLR = panel7.BackColor;
                    ManupCLR = panel11.BackColor;
                    TextCLR = label2.ForeColor;
                    break;
                case (2):
                    MainCLR = panel4.BackColor;
                    ContainerCLR = panel8.BackColor;
                    ManupCLR = panel12.BackColor;
                    TextCLR = label3.ForeColor;
                    break;
                case (3):
                    MainCLR = panel5.BackColor;
                    ContainerCLR = panel9.BackColor;
                    ManupCLR = panel13.BackColor;
                    TextCLR = label13.ForeColor;
                    break;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            RegistryKey txtRedOption = Registry.CurrentUser;
            RegistryKey Interface = txtRedOption.CreateSubKey("Interface");
            Interface.SetValue("MC", MainCLR.Name);
            Interface.SetValue("CC", ContainerCLR.Name);
            Interface.SetValue("MPC", ManupCLR.Name);
            Interface.SetValue("TC", TextCLR.Name);
            button3_Click(sender, e);
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
