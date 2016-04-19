using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;

namespace AlertOutlookAddIn
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            this.MaximumSize = this.Size;
            this.MinimumSize = this.Size;
            // 最大化ボタンを無効にする
            this.MaximizeBox = false;
            // フォームのコンストラクタ内で設定
            this.StartPosition = FormStartPosition.CenterScreen;

            // サイズ変更不可の直線ウィンドウに変更する
            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            button1.Click += new EventHandler(this.Button1_Click);
            button2.Click += new EventHandler(this.Button2_Click);
            button3.Click += new EventHandler(this.Button3_Click);
            button4.Click += new EventHandler(this.Button4_Click);
            button5.Click += new EventHandler(this.Button5_Click);


            string domain = Properties.Settings.Default.domain;

            if(domain != null){
                // カンマ区切りで分割して配列を格納する
                listBox1.Items.AddRange(domain.Split(','));
            }

            checkBox1.Checked = Properties.Settings.Default.appendfile;


        }

        //追加
        private void Button1_Click(object sender, EventArgs e)
        {
            Form2 cForm2 = new Form2();
            cForm2.ShowDialog();
            if (cForm2.strParam != null)
            {
                listBox1.Items.Add(cForm2.strParam);
            }
            cForm2.Dispose();
        }

        //編集
        private void Button2_Click(object sender, EventArgs e)
        {
            string item = "";
            // 選択されている項目のインデックスを取得
            int index = listBox1.SelectedIndex;

            // 選択されていればインデックスに０以上の値が入ります
            if (index >= 0)
            {
                // 選択されている項目を表示（Itemsはobjectなのでキャストしています）
                item = (string)listBox1.Items[index];

            }
            else
            {
                return;

            }



            Form2 cForm2 = new Form2(item);
            cForm2.ShowDialog();
            if (cForm2.strParam != null)
            {
                listBox1.Items[index] = cForm2.strParam;
            }
            cForm2.Dispose();
        }

        //削除
        private void Button3_Click(object sender, EventArgs e)
        {
            int index = listBox1.SelectedIndex;
            listBox1.Items.RemoveAt(index);
        }



        //保存
        private void Button4_Click(object sender, EventArgs e)
        {
            string domain = "";
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                if (i == 0)
                {
                    domain = (string)listBox1.Items[i];
                }
                else
                {
                   domain = domain + "," + (string)listBox1.Items[i];

                }
            }
            Properties.Settings.Default.domain = domain;
            Properties.Settings.Default.appendfile = checkBox1.Checked;
            Properties.Settings.Default.Save();

            // 自身のフォームを閉じる
            this.Close();
        }

        
        //キャンセル
        private void Button5_Click(object sender, EventArgs e)
        {
            // 自身のフォームを閉じる
            this.Close();
        }


    }
}
