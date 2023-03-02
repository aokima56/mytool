using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Runtime.InteropServices;





namespace AnalyzeForFukuhon1stUpload
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Folderを選択させる
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            //fbd.RootFolder = Environment.SpecialFolder.Desktop;
            
            //初期表示フォルダはカレントフォルダ
            string strNowDirectory =  Directory.GetCurrentDirectory();
            fbd.SelectedPath = strNowDirectory;


            //fbd.RootFolder = Environment.SpecialFolder("C:\\Users\aokima\\Desktop\\仕事\\20170222 宿題\\Excel");


            if (fbd.ShowDialog(this) == DialogResult.OK) 

            {
                textBox1.Text = fbd.SelectedPath;
            }


        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Excel格納パスを取得する。
            string strExcelFolder = string.Empty;

            if (this.textBox1.Text == "")
            {
                MessageBox.Show("まずはファイル格納パスを指定しろ。話はそれからだ。");
                return;

            }

            strExcelFolder = this.textBox1.Text;


            DialogResult result = MessageBox.Show("過去データ消していいよね？", "確認", MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes)
            {

                //まずはデータを消す
                EntityClass MyDb = new EntityClass();

                if (MyDb.TruncateTableTokukoData())
                {
                    MessageBox.Show("過去データを消去しました。\r\n 続けてデータを読み込みます。");
                }
                else
                {
                    MessageBox.Show("過去データの消去に失敗したよ。処理中断します。");
                    return;


                }



                //指定パスのすべてのExcelファイルをリスト化する
                string strFileList = string.Empty;

                //全処理件数格納用
                int ExecCnt = new int();


                //ファイルがある分、処理を継続
                foreach(string stFilePath in Directory.GetFiles(strExcelFolder, "*.xlsx"))
                {
                    //カーソル変更
                    Cursor.Current = Cursors.WaitCursor;

                    //New ExcelObject
                    Excel.Application oXls; // Excelオブジェクト
                    oXls = new Excel.Application();
                    Excel.Workbook oWBook; // workbookオブジェクト

                    try
                    {

                        //oXls.Visible = true; // 確認のためExcelのウィンドウを表示する
                        oWBook = (Excel.Workbook)(oXls.Workbooks.Open(stFilePath, Type.Missing, true));  // オープンするExcelファイル名

                        //シートの指定
                        // 与えられたワークシート名から、Worksheetオブジェクトを得る
                        string sheetName = "整理シート";
                        Excel.Worksheet oSheet; // Worksheetオブジェクト
                        oSheet = (Excel.Worksheet)oWBook.Sheets[getSheetIndex(sheetName, oWBook.Sheets)];

                        //OpenFile


                        //指定箇所の読み取り

                        Excel.Range rng; // Rangeオブジェクト

                        /*----------------------------------------------
                         * 当社で複本登録する対象の特個ファイルの件数とサイズを
                         * 開いた別紙1のxlsxからRanegで取得し、LongでArrayListにぶっこむ、
                         * DBにInsertする
                         ---------------------------------------------*/

                        //Entityクラスに読み込んだ件数、サイズを渡すためのArrayList
                        ArrayList alDataCnt = new ArrayList();
                        ArrayList alDataSize = new ArrayList();
                        ArrayList alTokukoKanriNo = new ArrayList();



                        /*---------------------------------------------
                    　    管理番号1 住民記録 10行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("1");

                        //件数
                        rng = (Excel.Range)oSheet.Cells[10, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[10, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号2 税情報 12行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("2");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[12, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[12, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号3 児童手当 14行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("3");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[14, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[14, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号4 介護 16行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("4");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[16, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[16, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号8 障害者福祉 27行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("8");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[27, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[27, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号9 障害者福祉 37行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("9");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[37, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[37, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号10 障害者福祉 40行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("10");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[40, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[40, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号16 児童扶養手当 51行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("16");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[51, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[51, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号31 国保 103行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("31");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[103, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[103, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号33 介護 141行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("33");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[141, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[141, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号34 介護 153行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("34");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[153, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[153, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号36 介護 165行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("36");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[165, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[165, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号37 国保介護 173行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("37");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[173, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[173, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号38 国保 187行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("38");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[187, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[187, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号39 国保 196行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("39");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[196, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[196, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号43 国保 205行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("43");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[205, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[205, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号44 国保 222行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("44");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[222, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[222, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号46 国保介護 237行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("46");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[237, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[237, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号47 国保介護 251行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("47");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[251, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[251, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号50 国保介護 278行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("50");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[278, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[278, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号80 健康管理 310行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("80");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[310, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[310, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号81 国保 312行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("81");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[312, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[312, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号83 国保介護 324行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("83");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[324, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[324, 16];
                        alDataSize.Add((long)rng.Value);

                        /*---------------------------------------------
                    　    管理番号84 健康管理 340行目
                         --------------------------------------------*/
                        alTokukoKanriNo.Add("84");
                        //件数
                        rng = (Excel.Range)oSheet.Cells[340, 15];
                        alDataCnt.Add((long)rng.Value);

                        //データサイズ
                        rng = (Excel.Range)oSheet.Cells[340, 16];
                        alDataSize.Add((long)rng.Value);




                        /*---------------------------------------------
                    　
                         --------------------------------------------*/



                        //解放
                        //oWBook.Close(Type.Missing, Type.Missing, Type.Missing);

                        //ファイル名から市町村コード取得
                        string strCityCd = string.Empty;
                        strCityCd = oWBook.Name.Substring(0,5);


                        oWBook.Close(false);

                        oXls.Quit();
                    
 

                        /*---------------------------------------------
                    　    EntityClassにArrayList渡して、中でぐるぐるInsert
                         --------------------------------------------*/
                        //EntityClass MyDb = new EntityClass();

                        MyDb.InsertDataCntAndDataSize(strCityCd, alTokukoKanriNo,alDataCnt, alDataSize);


                        //全処理件数を加算
                        ExecCnt += alDataCnt.Count;

                        ////forDebug ラベルに書き出し
                        ////strFileList += stFilePath + System.Environment.NewLine;

                        //foreach (long DataCnt in alDataCnt)
                        //{
                        //    strFileList += DataCnt.ToString() + Environment.NewLine;
                        
                        //}

                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    finally
                    {

                    }

                }

                //カーソル変更
                Cursor.Current = Cursors.Default;

                //
                //this.lblFileList.Text = strFileList;
                MessageBox.Show("データを"+ExecCnt.ToString()+"件保存したよ。");
            }

        }

        //

        private void btnTrncate_Click(object sender, EventArgs e)
        {

            DialogResult result = MessageBox.Show("過去データ消していいよね？","確認", MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes)
            {

                EntityClass MyDb = new EntityClass();

                if (MyDb.TruncateTableTokukoData())
                {
                    MessageBox.Show("過去データを消去しました。");
                }
                else
                {
                    MessageBox.Show("過去データの消去に失敗したよ");

                }
            }
        }

        // 指定されたワークシート名のインデックスを返すメソッド
        private int getSheetIndex(string sheetName, Excel.Sheets shs)
        {
            int i = 0;
            foreach (Excel.Worksheet sh in shs)
            {
                if (sheetName == sh.Name)
                {
                    return i + 1;
                }
                i += 1;
            }
            return 0;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            frmMenteUploadCombi objfrmMenteUploadCombi = new frmMenteUploadCombi();
            objfrmMenteUploadCombi.ShowDialog(this);


        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnSumUploadDataSizeByCity_Click(object sender, EventArgs e)
        {

            //File保存場所指定させる

            // FileStreamオープン




            //SaveFileDialogクラスのインスタンスを作成
            SaveFileDialog sfd = new SaveFileDialog();

            //はじめのファイル名を指定する
            //はじめに「ファイル名」で表示される文字列を指定する
            sfd.FileName = "団体ごと書き出しサイズ.csv";
            //[ファイルの種類]に表示される選択肢を指定する
            sfd.Filter = "CSV(*.csv)|*.csv|すべてのファイル(*.*)|*.*";
            //[ファイルの種類]ではじめに選択されるものを指定する
            //2番目の「すべてのファイル」が選択されているようにする
            sfd.FilterIndex = 1;
            //タイトルを設定する
            sfd.Title = "保存先のファイルを選択してください";
            //ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
            sfd.RestoreDirectory = true;

            //ダイアログを表示する
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                //カーソル変更
                Cursor.Current = Cursors.WaitCursor;

                //OKボタンがクリックされたとき、
                //選択された名前で新しいファイルを作成し、
                //読み書きアクセス許可でそのファイルを開く。
                //既存のファイルが選択されたときはデータが消える恐れあり。
                System.IO.Stream stream;
                stream = sfd.OpenFile();
                if (stream != null)
                {
                    //ファイルに書き込む
                    System.IO.StreamWriter sw = new System.IO.StreamWriter(stream);

                    //集計クラスの生成（メモリ上でdrから読みだして、文字列作成
                    EntityClass MyDb = new EntityClass();
                    string strDev = string.Empty;
                    sw.Write(MyDb.SumUploadDataSizeByCity());

                    //StreamWriterとFileStrem閉じる
                    sw.Close();
                    stream.Close();

                    Cursor.Current = Cursors.Default;

                    MessageBox.Show("CSVに出力しました。");
                }
            }

        }

        private void btnUploadDataByCity_Click(object sender, EventArgs e)
        {
            //Datasetを利用して、Excelに団体ごとのシートを作成
            /*
             *調べること 
             * ①DataSet To ExcelWorkSheet
             * ②
             * 
             */


        }
    }
}
