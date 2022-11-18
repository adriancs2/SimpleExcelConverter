using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ExcelDocConverter
{
    public partial class Form1 : Form
    {
        public Form1(string message, string title)
        {
            InitializeComponent();
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<html>\r\n<head>\r\n<style>\r\nbody { font-family: Arial; line-height: 150%; }\r\n</style>\r\n</head>");
            sb.AppendLine("<body>");
            sb.AppendLine(EscapeStringSequence(message));
            sb.AppendLine("</body>");
            sb.AppendLine("</html>");
            webBrowser1.DocumentText = sb.ToString();
            this.Text = title;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        
        string EscapeStringSequence(string input)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            foreach (char c in input)
            {
                switch (c)
                {
                    case '&':
                        sb.AppendFormat("&amp;");
                        break;
                    case '"':
                        sb.AppendFormat("&quot;");
                        break;
                    case '\'':
                        sb.AppendFormat("&#39;");
                        break;
                    case '<':
                        sb.AppendFormat("&lt;");
                        break;
                    case '>':
                        sb.AppendFormat("&gt;");
                        break;
                    default:
                        sb.Append(c);
                        break;
                }
            }
            sb.Replace("&#39;r&#39;n", "<br />");
            sb.Replace("\r\n", "<br />");
            return sb.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }    

    }
}
