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

                if (mailItem.Body.Contains("添付") == true &&
                  mailItem.Attachments.Count == 0)
                {
                    if (MessageBox.Show("添付ファイルがりませんが送信しますか？", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.Cancel)
                    {
                        cancel = true;
                        return;
                    }
                }
                
                string strMsg = "";
                strMsg = "件名：" + mailItem.Subject + "\n";
                strMsg = strMsg + "宛先:\n";
                const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

                Boolean msgboxflg = false;


                string mailadress = "";
                
                foreach (Outlook.Recipient recip in mailItem.Recipients) 
                {
                    Outlook.PropertyAccessor pa = recip.PropertyAccessor;
                            string smtpAddress = pa.GetProperty(PR_SMTP_ADDRESS).ToString();
                            if(smtpAddress.IndexOf("@lac.co.jp") < 0)
                            {
                                msgboxflg = true;
                            }

                            mailadress = mailadress + "           " + smtpAddress + "\n";
                }

                strMsg = strMsg + mailadress + "\n";

                 strMsg = strMsg + "上記の宛先に、メールを送信してもよろしいですか?\n";

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
