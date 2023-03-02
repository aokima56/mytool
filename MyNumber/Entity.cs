using System;
using System.IO;
using System.Collections;
using System.Data.SqlClient;



public class EntityClass
 /*
  * DB周りをここで
     
 */
{
	public EntityClass()
	{
	}


    public Boolean TruncateTableTokukoData()
    {
        // 接続文字列を生成する
        string stConnectionString = string.Empty;

        // SqlConnection の新しいインスタンスを生成する (接続文字列を指定)
        SqlConnection cSqlConnection = new SqlConnection(SetConnetionStrings());

        string strTrucateTabel = string.Empty;
        var sqlCmdTruncate = cSqlConnection.CreateCommand();

        try
        {
            // データベース接続を開く
            cSqlConnection.Open();

            //まずは対象をTruncate
            sqlCmdTruncate.CommandText = @"Truncate table [dbo].[E_TokukoRecordCount]";
            sqlCmdTruncate.ExecuteNonQuery();

            cSqlConnection.Close();
            cSqlConnection.Dispose();

            return true;
        }
        catch (Exception ex)
        {  
            return false;
            throw ex;
        }


    }

    public Boolean InsertDataCntAndDataSize(string CityCd,ArrayList alTokukoKanriNo,ArrayList alDataCnt, ArrayList alDataSize)
    {


        try
        {
            //DB ConnectionObject

            // 接続文字列を生成する
            string stConnectionString = string.Empty;

            // SqlConnection の新しいインスタンスを生成する (接続文字列を指定)
            SqlConnection cSqlConnection = new SqlConnection(SetConnetionStrings());

            string strTrucateTabel = string.Empty;
            var sqlCmdTruncate = cSqlConnection.CreateCommand();
            var sqlCmdInsDataCntAndSize = cSqlConnection.CreateCommand();



            // データベース接続を開く
            cSqlConnection.Open();


            //最大行数を取得（idの値を決定する）
            sqlCmdInsDataCntAndSize.CommandText = @"SELECT MAX(id) FROM [dbo].[E_TokukoRecordCount]";
            SqlDataReader dr = sqlCmdInsDataCntAndSize.ExecuteReader();
            int idMax = new int();
            //確実に1件しかないから
            while (dr.Read())
            {
                if (dr[0].ToString() != "")
                {
                    idMax = int.Parse(dr[0].ToString());
                }

            }
            dr.Close();

            // CreateInsertStatement
            for (int i = 0; i < alDataCnt.Count; i++)
            {
                //idを++
                idMax++;

                sqlCmdInsDataCntAndSize.CommandText = @"";
                //sqlCmdInsDataCntAndSize.CommandText = @"INSERT INTO  [dbo].[E_TokukoRecordCount](id,DataCount,DataSizeByte) VALUES (" + idMax.ToString() + "," + alDataCnt[i].ToString() + "," + alDataSize[i].ToString() + ")";
                sqlCmdInsDataCntAndSize.CommandText = @"INSERT INTO  [dbo].[E_TokukoRecordCount] VALUES (" 
                                                        + idMax.ToString() + ",'" + CityCd + "','"+ alTokukoKanriNo[i]+ "',"+ alDataCnt[i].ToString() + "," + alDataSize[i].ToString() + ")";
                sqlCmdInsDataCntAndSize.ExecuteNonQuery();

            }


            // データベース接続を閉じる (正しくは オブジェクトの破棄を保証する を参照)
            cSqlConnection.Close();
            cSqlConnection.Dispose();

            return true;

        }
        catch (Exception ex)
        {
            ex.ToString();

            return false;
            throw;
        }
    }

    private string SetConnetionStrings()
    {
        var builder = new SqlConnectionStringBuilder()
        {
            //ホントならapp.configくらい使え

            //DataSource = "IPアドレス指定してね",
            DataSource = "",
            IntegratedSecurity = false,
            UserID = "",
            Password = "",
            InitialCatalog ="anything"

        };

        return builder.ToString();
    }
       
    public string SumUploadDataSizeByCity()
    {

        //Conn.Open
        // SqlConnection の新しいインスタンスを生成する (接続文字列を指定)
        SqlConnection cSqlConnection = new SqlConnection(SetConnetionStrings());

        string strTrucateTabel = string.Empty;
        var sqlCmd = cSqlConnection.CreateCommand();

        //Retun用のStringを用意
        string strReturn = string.Empty;



        // データベース接続を開く
        cSqlConnection.Open();


        //集計用SQL
        /*
SELECT B.CityCd, A.CombiCd, (SUM(B.DataSizeByte)/1024/1024)
FROM M_Combi AS A, E_TokukoRecordCount AS B
WHERE B.TokukoFileKanriNo = A.TokukoFileKanriNo
GROUP BY A.CombiCd,B.CityCd 
ORDER BY B.CityCd,A.CombiCd; 

SELECT B.CityCd, C.CityName,A.CombiCd,  SUM(B.DataCount), (SUM(B.DataSizeByte)/1024/1024)
FROM M_Combi AS A, E_TokukoRecordCount AS B,M_Dantai AS C
WHERE B.TokukoFileKanriNo = A.TokukoFileKanriNo AND C.CityCd= B.CityCd
GROUP BY B.CityCd, C.CityName,A.CombiCd
ORDER BY B.CityCd,A.CombiCd; 

 * 
 *          */

        sqlCmd.CommandText = @"SELECT B.CityCd, C.CityName, A.CombiCd,  SUM(B.DataCount), (SUM(B.DataSizeByte)/1024/1024)" +
            " FROM M_Combi AS A, E_TokukoRecordCount AS B, M_Dantai AS C" +
            " WHERE B.TokukoFileKanriNo = A.TokukoFileKanriNo AND C.CityCd= B.CityCd" +
            " GROUP BY A.CombiCd,B.CityCd,CityName " +
            " ORDER BY B.CityCd,A.CombiCd";

        SqlDataReader dr = sqlCmd.ExecuteReader();

        //Loop内 変数
        string strCityCd = string.Empty;
        int i = new int();

        //読み込み
        while (dr.Read())
        {   
            //読み込みが同じ市町村ならCombiCdとサイズをカンマ区切りで
            if (i == 0)
            {
                //最初の市町村の書き出しを実施
                strReturn += "市町村コード,市町村名,コマ,データ件数,データサイズ(MB),コマ,データ件数,データサイズ(MB),税コマ,データ件数,データサイズ(MB)r\n";
                strReturn += dr[0] + "," + dr[1] + "," + dr[2] + "," + dr[3] + "," + dr[4];
            }
            else
            {
                //2行目以降の処理
                if (dr[0].ToString() == strCityCd)
                {
                    //同じ市町村は前回行にAppendする
                    strReturn += "," + dr[2] + "," + dr[3] + "," + dr[4];
                }
                else
                {
                    //市町村が変わっていたら改行して書き出し
                    strReturn += "\r\n" + dr[0] + "," + dr[1] + "," + dr[2] + "," + dr[3] + "," + dr[4];
                }
            }

            //市町村判定用
            strCityCd = "";
            strCityCd = dr[0].ToString();

            i++;

        }
        dr.Close();



        return strReturn;

    }

}
