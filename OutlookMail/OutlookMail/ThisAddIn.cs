using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

namespace OutlookMail
{
    public partial class ThisAddIn
    {
        private Office.CommandBar menuBar;
        private Office.CommandBarPopup newMenuBar;
        private Office.CommandBarButton buttonOne;
        private string menuTag = "WorldAddIn";
        //private object applicationObject;
        //private object addInInstance;
        public Outlook.Application OutlookApplication;
        public Outlook.Inspectors OutlookInspectors;
        public Outlook.Inspector OutlookInspector;
        public Outlook.MailItem OutlookMailItem;

        #region Add Button from Outlook's Menu
        private void AddMenuBar()
        {
            try
            {
                //Define the existent Menu Bar
                menuBar = this.Application.ActiveExplorer().CommandBars.ActiveMenuBar;
                //Define the new Menu Bar into the old menu bar
                newMenuBar = (Office.CommandBarPopup)menuBar.Controls.Add(
                    Office.MsoControlType.msoControlPopup, missing,
                    missing, missing, false);
                //If I dont find the newMenuBar, I add it
                if (newMenuBar != null)
                {
                    newMenuBar.Caption = "World Tour Agency";
                    newMenuBar.Tag = menuTag;
                    buttonOne = (Office.CommandBarButton)newMenuBar.Controls.
                    Add(Office.MsoControlType.msoControlButton, missing,
                        missing, 1, true);
                    buttonOne.Style = Office.MsoButtonStyle.
                        msoButtonIconAndCaption;
                    buttonOne.Caption = "Send Confirmation Email";
                    buttonOne.Click += new Office._CommandBarButtonEvents_ClickEventHandler(buttonOne_Click);
                    //This is the Icon near the Text
                    buttonOne.FaceId = 610;
                    buttonOne.Tag = "c123";
                    //Insert Here the Button1.Click event    
                    newMenuBar.Visible = true;
                }
            }
            catch (Exception ex)
            {
                //This MessageBox is visible if there is an error
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message.ToString(), "Error Message Box", MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation);
            }
        }
        #endregion

        #region Remove Button from Outlook's Menu
        private void RemoveMenubar()
        {
            // If the menu already exists, remove it.
            try
            {
                Office.CommandBarPopup foundMenu = (Office.CommandBarPopup)
                    this.Application.ActiveExplorer().CommandBars.ActiveMenuBar.
                    FindControl(Office.MsoControlType.msoControlPopup,
                    missing, menuTag, true, true);
                if (foundMenu != null)
                {
                    foundMenu.Delete(true);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
        #endregion

        //Outlook Compose Email

        //private void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        //{
        //    Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
        //    if (mailItem != null)
        //    {
        //        if (mailItem.EntryID == null)
        //        {
        //            mailItem.Subject = "This text was added by using code";
        //            mailItem.Body = "This text was added by using code";
        //        }

        //    }
        //}

        void OutlookInspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            OutlookInspector = (Outlook.Inspector)Inspector;
            if (Inspector.CurrentItem is Outlook.MailItem)
            {
                OutlookMailItem = (Outlook.MailItem)Inspector.CurrentItem;
            }

        }

        void OutlookApplication_ItemSend(object Item, ref bool Cancel)
        {
           
            //string strchkTo = OutlookMailItem.To;
            //string strchk = "hello Welcome to c#";
            //MessageBox.Show(strchk + "\r\n" + strchkTo);
        }


        //void OutlookApplication_ItemSends(object Item, ref bool Cancel)
        //{
        //    //string strchkTo = OutlookMailItem.To;
        //    //string strchk = "hello Welcome to c#";
        //    //MessageBox.Show(strchk + "\r\n" + strchkTo);
        //}

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            OutlookApplication = Application as Outlook.Application;
            OutlookInspectors = OutlookApplication.Inspectors;
            OutlookInspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(OutlookInspectors_NewInspector);
            OutlookApplication.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(OutlookApplication_ItemSend);
            //OutlookApplication.ItemSend += new Outlook.ApplicationEvents_ItemSendEventHandler(OutlookApplication_ItemSend);

            //Outlook.Inspectors inspectors;
            //inspectors = this.Application.Inspectors;
            //inspectors.NewInspector += new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);

            AddMenuBar();
            RemoveMenubar();


            // Disable Ribbon 
            //Type type = typeof(MailRead);
            //MailRead ribbon = Globals.Ribbons.GetRibbon(type) as MailRead;
            //ribbon.ReadEmail.Enabled = false;

            // Visible false Ribbon 
            //Type type = typeof(MailRead);
            //MailRead ribbon = Globals.Ribbons.GetRibbon(type) as MailRead;
            //ribbon.ReadEmail.Visible = false;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void buttonOne_Click(Office.CommandBarButton ctrl, ref bool cancel)
        {

        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
