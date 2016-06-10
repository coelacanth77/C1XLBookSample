using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace C1XLBookSample
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // エクセルファイルを開く
            this.c1XLBook1.Load("C:\\sample.xlsx");

            // 一番目のシートの一列一行目の文字列を取得
            this.label1.Text = c1XLBook1.Sheets[0][0, 0].Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // 一列一行目の値を書き換える
            c1XLBook1.Sheets[0][0, 0].Value = this.textBox1.Text;

            this.label1.Text = this.textBox1.Text;

            // エクセルの数式を使って計算する
            c1XLBook1.Sheets[0][3, 0].Value = 5;
            c1XLBook1.Sheets[0][4, 0].Value = 2;

            c1XLBook1.Sheets[0][5, 0].Formula = "=SUM(A4:A5)";

            // エクセルファイルを保存する
            c1XLBook1.Save("C:\\sample.xlsx");
        }
    }
}
