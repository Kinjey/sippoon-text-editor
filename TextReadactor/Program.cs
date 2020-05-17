using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;

namespace TextReadactor
{
    static class Program
    {
        public static int New_Form_Count = 0;
        public static string Font_Name;
        public static float Font_Size;
        public static Form Redact;
        public static RichTextBox RedactTB;
        public static FontStyle style;
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
