using System;
using System.DirectoryServices.AccountManagement;
using System.Windows.Forms;


namespace Groups
{
    public partial class Form1 : Form
    {
        String hoje = DateTime.Now.ToString("dd.MM.yyyy");
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PrincipalContext context2 = new PrincipalContext(ContextType.Domain, "gbl.ad.hedani.net", "OU=BR,OU=CS,DC=gbl,DC=ad,DC=hedani,DC=net");
            PrincipalContext context = new PrincipalContext(ContextType.Domain, "csfb.cs-group.com", "OU=GroupsLocal,OU=SAO,OU=FAO,DC=csfb,DC=cs-group,DC=com");
            

            foreach (var linha in richTextBox3.Lines)
            {
                UserPrincipal user = UserPrincipal.FindByIdentity(context2, richTextBox3.Text);
                foreach (var line in richTextBox1.Lines)
                {
                    if (line != "")
                    {

                        try
                        {
                            GroupPrincipal group = GroupPrincipal.FindByIdentity(context, line);
                            group.Members.Add(user);
                            group.Save();
                            richTextBox2.AppendText("Usuario: " + user + " adicionado ao grupo" + line + "\n");

                        }
                        catch (Exception x)
                        {
                            MessageBox.Show(x.Message);
                        }

                    }
                    Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
                    Microsoft.Office.Interop.Outlook.MailItem mailItem = app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                    mailItem.To = "list.csbg-usr-support@credit-suisse.com";
                    mailItem.Subject = "AD Log - " + Environment.UserName + " - " + hoje ;
                    mailItem.Body = richTextBox2.Text;
                    mailItem.Send();
                    
                }
            }

             

        }

        
        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}

