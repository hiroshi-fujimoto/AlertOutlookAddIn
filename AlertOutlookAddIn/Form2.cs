using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AlertOutlookAddIn
{
    public partial class Form2 : Form
    {

        private String _strParam = null;


        public Form2()
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

        }


        public Form2(params string[] argumentValues)
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

            this.textBox1.Text = argumentValues[0];
        }


        //保存
        private void Button1_Click(object sender, EventArgs e)
        {

            _strParam = this.textBox1.Text;
            // 自身のフォームを閉じる
            this.Close();
        }


        //キャンセル
        private void Button2_Click(object sender, EventArgs e)
        {
            _strParam = null;
            // 自身のフォームを閉じる
            this.Close();
        }

        public String strParam
        {
            get
            {
                return _strParam;
            }
            set
            {
                _strParam = value;
            }
        }
   
    }
}
