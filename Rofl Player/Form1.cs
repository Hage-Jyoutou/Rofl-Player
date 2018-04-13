using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Rofl_Player
{
    public partial class Form1 : Form
    {
        string ClientPath;
        string ReplayPath;



        public Form1()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {

            //コントロールをすべて取得して、ドラッグドロップのイベントを設定//
            List<Control> controls = GetAllControls<Control>(this);
            foreach (Control c in Controls)
            {
                controls.Add(c);
            }

            for(int i =0; i < controls.Count(); i++)
            {
                controls[i].AllowDrop = true;
                controls[i].DragEnter += new DragEventHandler(DragEnter);
                controls[i].DragDrop += new DragEventHandler(Drop);


            }


            //設定ファイルからクライアントのファイルパスがあれば読み込んで設定する//
            if(Properties.Settings.Default.ClientDir != "")
            {
                if (System.IO.File.Exists(Properties.Settings.Default.ClientDir))
                {
                    ClientPath = Properties.Settings.Default.ClientDir;
                    textBox1.Text = ClientPath;
                    label3.Text = "◎";
                }
            }
            this.MaximumSize = this.Size;
            this.MinimumSize = this.Size;
        }

        //フォーム上のコントロールを取得//
        public static List<T> GetAllControls<T>(Control top) where T : Control
        {
            List<T> buf = new List<T>();
            foreach (Control ctrl in top.Controls)
            {
                if (ctrl is T) buf.Add((T)ctrl);
                buf.AddRange(GetAllControls<T>(ctrl));
            }
            return buf;
        }


        //LOLクライアント実行ファイルを開く
        private void button1_Click(object sender, EventArgs e)
        {
            //ファイルを開くダイアログを表示//
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = Properties.Settings.Default.releases;

            string tmpPath = "null";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //OKボタンがクリックされたとき、tmpPathにファイルパスを一時格納//
                tmpPath = ofd.FileName;

                //ファイルがLOLClientであればClientPathにファイルパスを格納//
                //違うのであればnullを格納し、警告文を表示//
                if (System.IO.Path.GetFileName(tmpPath) == @"League of Legends.exe")
                {
                    ClientPath = tmpPath;
                    textBox1.Text = tmpPath;
                    label3.Text = "◎";

                    string tmp = tmpPath.Substring(0, tmpPath.LastIndexOf("releases"));

                    Properties.Settings.Default.releases = tmp;
                    Properties.Settings.Default.ClientDir = tmpPath;
                    Properties.Settings.Default.Save();
                }
                else
                {
                    ClientPath = "null";
                    textBox1.Text = "LOLクライアントを選択してください";
                    label3.Text = "✕";
                }
            }
        }


        //roflファイルを開く
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = Properties.Settings.Default.ReplayDir;
            string tmpPath = "null";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //OKボタンがクリックされたとき、tmpPathにファイルパスを一時格納//
                tmpPath = ofd.FileName;

                //roflファイルであればReplayPathにファイルパスを格納//
                //違うのであればnullを格納し、警告文を表示//
                if (System.IO.Path.GetExtension(tmpPath) == @".rofl")
                {
                    ReplayPath = tmpPath;
                    label2.Text = System.IO.Path.GetFileName(tmpPath);
                    label4.Text = "◎";

                    string tmp = tmpPath.Substring(0, tmpPath.LastIndexOf(@"\"));
                    Properties.Settings.Default.ReplayDir = tmp;
                    Properties.Settings.Default.Save();
                }
                else
                {
                    ReplayPath = "null";
                    label2.Text = "roflファイルを選択してね";
                    label4.Text = "✕";
                }
            }
        }

        void button3_Click(object sender, EventArgs e)
        {

            if (label4.Text == "◎" & label3.Text == "◎")
            {
                //ショートカットを作成する//

                string shortcutPath = System.IO.Directory.GetCurrentDirectory() + "\\" + "test.lnk";
                IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
                IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(shortcutPath); //ショートカットのパスを指定して、WshShortcutを作成

                shortcut.TargetPath = ClientPath;               //リンク先
                shortcut.Arguments = "\"" + ReplayPath + "\"";  //コマンドパラメータ を設定
                shortcut.WorkingDirectory = System.IO.Path.GetDirectoryName(ClientPath);    //作業ディレクトリを設定
                shortcut.Save();                                                            //ショートカットを作成
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shortcut);     //後処理
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shell);

                var proc = new Process(); //新しいプロセスを作成
                proc.StartInfo.FileName = shortcutPath; // ショートカットパスをセット
                proc.Start(); //　プロセスをスタート
                proc.WaitForExit(); //プロセスが終了するまで待機

                //ショートカットを消去する//
                File.Delete(shortcutPath);
            }
        }


        //ドラッグを受け入れる//
        private new void DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        //ドロップされたものがroflファイルならReplayPathにファイルパスを格納//
        private void Drop(object sender,System.Windows.Forms.DragEventArgs e)
        {
            string[] tmpPath = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            if (System.IO.Path.GetExtension(tmpPath[0]) == @".rofl")
            {
                ReplayPath = tmpPath[0];
                label2.Text = System.IO.Path.GetFileName(tmpPath[0]);
                label4.Text = "◎";
            }
            else
            {
                ReplayPath = "null";
                label2.Text = "roflファイルを選択してね";
                label4.Text = "✕";
            }
        }
    }
}
