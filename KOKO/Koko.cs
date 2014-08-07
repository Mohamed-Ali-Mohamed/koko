using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using System.Windows.Documents;

namespace Koko
{
    public partial class Koko : Form
    {
        private mshtml.IHTMLDocument2 doc;
        public Koko()
        {
            InitializeComponent();
        }

        private String[] SplitInput(String Input)
        {
            String E;
            String[] All = Input.Split('\n');
            for (int i = 0; i < All.Length; i++)
            {
                E = "";
                for (int j = 0; j < All[i].Length; j++)
                {
                    if (All[i][j] != ' ')
                    {
                        E += All[i][j];
                    }
                }
                All[i] = E;
            }
            return All;
        }
        private void Attach(ref MailMessage mail, string path)
        {
            System.Net.Mail.Attachment attachment1;
            attachment1 = new System.Net.Mail.Attachment(path);
            mail.Attachments.Add(attachment1);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            RtfToHtmlConverter RTHC = new RtfToHtmlConverter();
            //-------------------------------------------------
            string FromEmail = Email.Text + ServerName.Text, Output = "", Send = "Your message has been sent to :\n",
                NotSend = "Your message \"has not\" been sent to :\n",
            FromName = User_Name.Text,
            subject = Message_Subject.Text,
            fromPassword = User_Password.Text,
            body = RTHC.ConvertRtfToHtml(Message_Body.Rtf),
            host ;
            int port,NumSend=1,NumNotSend=1;
            bool EnableSsl;
            //-------------------------------------------------
            switch (ServerName.Text)
            {
                case "@gmail.com":
                    host = "smtp.gmail.com";
                    port = 587;
                    EnableSsl = true;
                    break;
                case "@msn.com":
                case "@live.com":
                case "@hotmail.com":
                case "@outlook.com":
                    host = "smtp.live.com";
                    port = 587;
                    EnableSsl = true;
                    break;
                case "@aol.com":
                    host = "smtp.aol.com";
                    port = 587;
                    EnableSsl = true;
                    break;
                case "@yahoo.com":
                case "@ymail.com":
                case "@rocketmail.com":
                case "@mail.yahoo.com":
                    host = "smtp.mail.yahoo.com";
                    port = 465;
                    EnableSsl = false;
                    break;
                default:
                    OutputResult.Text="Error in your Email";
                    return;
            }
            String[] AllEmails = SplitInput(ToEmails.Text);
            for (int i = 0; i < AllEmails.Length; i++)
            {
                if (AllEmails[i] == "")
                    continue;
                try
                {
                    MailMessage mail = new MailMessage();
                    SmtpClient SmtpServer = new SmtpClient(host,port);
                    mail.From = new MailAddress(FromEmail, FromName);
                    mail.To.Add(AllEmails[i]);
                    mail.Subject = subject;
                    mail.Body = body;
                    mail.IsBodyHtml = true;
                    //-------------------------------------------------
                    if (Attach1.Text != "")
                    {
                        Attach(ref mail, Attach1.Text);
                    }
                    if (Attach2.Text != "")
                    {
                        Attach(ref mail, Attach2.Text);
                    }
                    if (Attach3.Text != "")
                    {
                        Attach(ref mail, Attach3.Text);
                    }
                    //
                    SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                    SmtpServer.UseDefaultCredentials = false;
                    SmtpServer.Credentials = new NetworkCredential(mail.From.Address, fromPassword);
                    SmtpServer.EnableSsl = EnableSsl;
                    //
                    SmtpServer.Send(mail);
                    Send += NumSend+++")\""+ AllEmails[i]+"\"\n";
                    //
                }
                catch
                {
                    NotSend += NumNotSend+++")\""+AllEmails[i]+"\"\n";
                }
            }
            Output = Send + "\n\n" + NotSend;
            OutputResult.Text = Output;
            MessageBox.Show("Look To Output");
        }

        private void fontDialog1_Apply(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
               Attach3.Text = openFileDialog1.FileName.ToString();
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                Attach2.Text = openFileDialog2.FileName.ToString();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (openFileDialog3.ShowDialog() == DialogResult.OK)
            {
                Attach1.Text = openFileDialog3.FileName.ToString();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Attach1.Text = "";
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Attach2.Text = "";
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Attach3.Text = "";
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                fontDialog1.ShowColor = true;
                if (fontDialog1.ShowDialog() == DialogResult.OK && !String.IsNullOrEmpty(Message_Body.Text))
                {
                    Message_Body.SelectionFont = fontDialog1.Font;
                    Message_Body.SelectionColor = fontDialog1.Color;
                }
            }
            catch
            {
                MessageBox.Show("Error");
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void ShowPassword_CheckedChanged(object sender, EventArgs e)
        {
            if (ShowPassword.Checked == true)
            {
                User_Password.UseSystemPasswordChar = false;
            }
            else
            {
                User_Password.UseSystemPasswordChar = true;
            }
        }

        private void newTS_Click(object sender, EventArgs e)
        {
            Message_Body.Text = "";
        }

        private void tsCut_Click(object sender, EventArgs e)
        {
            try
            {
                Message_Body.Cut();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Couldn't Cut\n\r" + ex.Message, "Erro Executing Cut Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tsCopy_Click(object sender, EventArgs e)
        {
            try
            {
                Message_Body.Copy();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Couldn't Copy\n\r" + ex.Message, "Erro Executing Copy Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tsPaste_Click(object sender, EventArgs e)
        {
            try
            {
                Message_Body.Paste();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Couldn't Paste\n\r" + ex.Message, "Erro Executing Paste Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tsBold_Click(object sender, EventArgs e)
        {
            try
            {
                if (Message_Body.SelectionFont.Bold)
                    Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily,Message_Body.SelectionFont.Size, FontStyle.Regular);
                else
                    Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, Message_Body.SelectionFont.Size, FontStyle.Bold);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Couldn't Bold\n\r" + ex.Message, "Erro Executing Bold Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tsUnderline_Click(object sender, EventArgs e)
        {
            try
            {
                if (Message_Body.SelectionFont.Underline)
                    Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, Message_Body.SelectionFont.Size, FontStyle.Regular);
                else
                    Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, Message_Body.SelectionFont.Size, FontStyle.Underline);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Couldn't Underline\n\r" + ex.Message, "Erro Executing Underline Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tsItalics_Click(object sender, EventArgs e)
        {
            try
            {
                if (Message_Body.SelectionFont.Italic)
                    Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, Message_Body.SelectionFont.Size, FontStyle.Regular);
                else
                    Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, Message_Body.SelectionFont.Size, FontStyle.Italic);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Couldn't Italic\n\r" + ex.Message, "Erro Executing Italic Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tsLeft_Click(object sender, EventArgs e)
        {
            try
            {
                Message_Body.SelectionAlignment = HorizontalAlignment.Left;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Couldn't Left\n\r" + ex.Message, "Erro Executing Left Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tsCenter_Click(object sender, EventArgs e)
        {
            try
            {
                Message_Body.SelectionAlignment = HorizontalAlignment.Center;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Couldn't Center\n\r" + ex.Message, "Erro Executing Center Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tsRight_Click(object sender, EventArgs e)
        {
            try
            {
                Message_Body.SelectionAlignment = HorizontalAlignment.Right;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Couldn't Right\n\r" + ex.Message, "Erro Executing Right Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tsIndent_Click(object sender, EventArgs e)
        {
            try
            {
                Message_Body.SelectionIndent += 30;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Couldn't RightIndent\n\r" + ex.Message, "Erro Executing RightIndent Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        private void tsOutdent_Click(object sender, EventArgs e)
        {
            try
            {
                Message_Body.SelectionIndent -= 30;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Couldn't LeftIndent\n\r" + ex.Message, "Erro Executing LeftIndent Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tsBullets_Click(object sender, EventArgs e)
        {
            try
            {
                if(Message_Body.SelectionBullet ==true)
                    Message_Body.SelectionBullet =false;
                else
                    Message_Body.SelectionBullet = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Couldn't SelectionBullet\n\r" + ex.Message, "Erro Executing SelectionBullet Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tsNumeric_Click(object sender, EventArgs e)
        {
            try
            {
                if (Message_Body.SelectionBullet == true)
                    Message_Body.SelectionBullet = false;
                else
                    Message_Body.SelectionBullet = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Couldn't SelectionNumber\n\r" + ex.Message, "Erro Executing SelectionNumber Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tsBackColor_Click(object sender, EventArgs e)
        {
            try
            {
                ColorDialog colorDialog = new ColorDialog();
                if (colorDialog.ShowDialog() != DialogResult.Cancel)
                {
                    Message_Body.SelectionBackColor = colorDialog.Color;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Couldn't BackColor\n\r" + ex.Message, "Erro Executing BackColor Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tsTextColor_Click(object sender, EventArgs e)
        {
            try
            {
                ColorDialog colorDialog = new ColorDialog();
                if (colorDialog.ShowDialog() != DialogResult.Cancel)
                {
                    Message_Body.SelectionColor = colorDialog.Color;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Couldn't Color\n\r" + ex.Message, "Erro Executing Color Command", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void tsFontSize_ButtonClick(object sender, EventArgs e)
        {
           
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, 8, Message_Body.SelectionFont.Style); 
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, 9, Message_Body.SelectionFont.Style);
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, 10, Message_Body.SelectionFont.Style);
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, 11, Message_Body.SelectionFont.Style);
        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, 12, Message_Body.SelectionFont.Style);
        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, 14, Message_Body.SelectionFont.Style);
        }

        private void toolStripMenuItem8_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, 16, Message_Body.SelectionFont.Style);
        }

        private void toolStripMenuItem9_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, 18, Message_Body.SelectionFont.Style);
        }

        private void toolStripMenuItem10_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, 20, Message_Body.SelectionFont.Style);
        }

        private void toolStripMenuItem11_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, 22, Message_Body.SelectionFont.Style);
        }

        private void toolStripMenuItem12_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, 24, Message_Body.SelectionFont.Style);
        }

        private void toolStripMenuItem13_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, 26, Message_Body.SelectionFont.Style);
        }

        private void toolStripMenuItem14_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, 28, Message_Body.SelectionFont.Style);
        }

        private void toolStripMenuItem15_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, 36, Message_Body.SelectionFont.Style);
        }

        private void toolStripMenuItem16_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, 48, Message_Body.SelectionFont.Style);
        }

        private void toolStripMenuItem17_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font(Message_Body.SelectionFont.FontFamily, 72, Message_Body.SelectionFont.Style);
        }

        private void verdanaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font("Verdana", Message_Body.SelectionFont.Size, Message_Body.SelectionFont.Style);
        }

        private void ariaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font("Arial", Message_Body.SelectionFont.Size, Message_Body.SelectionFont.Style);
        }

        private void timesNewRomanToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font("Times New Roman", Message_Body.SelectionFont.Size, Message_Body.SelectionFont.Style);
        }

        private void currierToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font("Currier New", Message_Body.SelectionFont.Size, Message_Body.SelectionFont.Style);
        }

        private void comicSansToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font("Cambria", Message_Body.SelectionFont.Size, Message_Body.SelectionFont.Style);
        }

        private void helveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font("Tahoma", Message_Body.SelectionFont.Size, Message_Body.SelectionFont.Style);
        }

        private void bookAntiquaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Message_Body.SelectionFont = new Font("Book Antiqua", Message_Body.SelectionFont.Size, Message_Body.SelectionFont.Style);
        }


        

        private void tsRemoveLink_Click(object sender, EventArgs e)
        {

        }
    }

}



