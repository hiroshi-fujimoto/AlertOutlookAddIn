using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace AlertOutlookAddIn
{
    public partial class ThisAddIn
    {

        Outlook.Application application;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            application = this.Application;
            //すべての送信処理を検知するイベントを追加
            (application as Outlook.ApplicationEvents_11_Event).ItemSend += OnItemSend;

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        private void OnItemSend(object item, ref bool cancel)
        {
            if (item is Outlook.MailItem)
            {
                var mailItem = item as Outlook.MailItem;

                Boolean appendfile = Properties.Settings.Default.appendfile;

                if (mailItem.Body.Contains("添付") == true &&
                  mailItem.Attachments.Count == 0 && appendfile == true)
                {
                    if (MessageBox.Show("添付ファイルがりませんが送信しますか？", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.Cancel)
                    {
                        cancel = true;
                        return;
                    }
                }
                
                string strMsg = "";
                strMsg = "件名：" + mailItem.Subject + "\n\n";
                const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";






                Boolean msgboxflg = false;
                string domains = Properties.Settings.Default.domain;
                string[] domainList = domains.Split(',');

                string mailadressTO = "";
                string mailadressCC = "";
                string mailadressBCC = "";
                int cnt = 0;
                
                foreach (Outlook.Recipient recip in mailItem.Recipients) 
                {
                    Outlook.PropertyAccessor pa = recip.PropertyAccessor;
                    
                    string smtpAddress = pa.GetProperty(PR_SMTP_ADDRESS).ToString();
                    foreach (string dm in domainList)
                    {
                        if (smtpAddress.IndexOf(dm) < 0)
                        {
                            msgboxflg = true;
                            smtpAddress = smtpAddress + " 【注意】 ";
                            cnt++;
                                                               }
                    }


                    if (recip.Type == 1) { 
                        mailadressTO = mailadressTO  + smtpAddress + "\n";
                    }
                    else if (recip.Type == 2)
                    {
                        mailadressCC = mailadressCC  + smtpAddress + "\n";

                    }else{
                        mailadressBCC = mailadressBCC  + smtpAddress + "\n";

                    }

                }

                strMsg = strMsg + "[　TO　] --------------------------------------\n\n";
                strMsg = strMsg +  mailadressTO + "\n";

                if (mailadressCC.Length > 0) 
                {
                    strMsg = strMsg + "[　CC　] --------------------------------------\n\n";
                    strMsg = strMsg + mailadressCC + "\n";
                }
                if (mailadressBCC.Length > 0)
                {
                    strMsg = strMsg + "[ BCC　] --------------------------------------\n\n";
                    strMsg = strMsg + mailadressBCC + "\n";
                }
                strMsg = strMsg +     "--------------------------------------------------\n\n";


                strMsg = strMsg + "上記の宛先に、メールを送信してもよろしいですか?\n";

                if (msgboxflg == true) strMsg = "【注意】確認が必要なメールアドレスが" + cnt + "つあります。\n\n\n" + strMsg;


                 if (MessageBox.Show(strMsg, "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.Cancel)
                {
                    cancel = true;
                }

            }
        }


        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
