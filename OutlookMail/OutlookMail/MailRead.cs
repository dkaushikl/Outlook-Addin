using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMail
{
    public partial class MailRead
    {
        private void MailRead_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ReadEmail_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("test");
        }
    }
}
