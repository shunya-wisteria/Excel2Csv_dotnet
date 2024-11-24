//----------------------------------------------------------
// Excel-CSV変換
//
//  概要：
//    xlsxファイルパスをinputに、対象ファイルを展開し、
//    シートごとに出力フォルダパスで指定された場所にcsv出力する
//
//  引数：
//    ・入力xlsxファイルパス
//    ・出力フォルダパス
// 
//  リターンコード：
//    ・81：引数不足
//    ・82：入力ファイルパス不正
//    ・83：出力フォルダ作成失敗
//    ・98：ファイル入出力エラー
//    ・99：その他例外
//----------------------------------------------------------
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Excel2CSv;

class Program
{
  // 入力ファイルパス
  private static string inFile = "";
  // 出力フォルダパス
  private static string outPath = "";

  /// <summary>
  /// メイン関数
  /// </summary>
  /// <param name="args">第1引数：入力xlsxファイルパス、第2引数：出力フォルダパス</param>
  /// <returns>リターンコード</returns>
  static int Main(string[] args)
  {
    int check = 0;

    // 引数チェック
    check = checkImportArgs(args);
    if(check != 0)
    {
      return check;
    }

    // 事前処理
    check = preProc();
    if(check != 0)
    {
      return check;
    }

    // 開始メッセージ
    Console.WriteLine("Excel-CSV変換処理を開始します。");
    Console.WriteLine("入力ファイル：" + inFile);

    Console.WriteLine("...");

    // 変換処理
    List<string> outFiles;
    check = excel2CSv(out outFiles);
    if(check != 0)
    {
      return check;
    }

    //  終了メッセージ
    Console.WriteLine("Excel-CSV変換処理が完了しました。\n出力ファイル：");
    foreach(string outFile in outFiles)
    {
      Console.WriteLine("・" + outFile);
    }

    return 0;
  }

  /// <summary>
  /// 引数チェック・引数取込み
  /// </summary>
  /// <param name="args">引数配列</param>
  /// <returns>リターンコード、81：引数不足、82：入力ファイルパス不正</returns>
  static private int checkImportArgs(string[] args)
  {
    // 引数不足
    if(args.Length < 2)
    {
      Console.WriteLine("引数が不足しています、第1引数：入力xlsxファイル、第2引数：出力フォルダパス");
      return 81;
    }

    // 入力ァイル存在チェック
    if(!File.Exists(args[0]))
    {
      Console.WriteLine("入力xlsxファイルが存在しません。");
      return 82;
    }

    // 引数をフィールド変数に取り込み
    inFile = args[0];
    outPath = args[1];

    return 0;
  }

  /// <summary>
  /// 事前処理
  /// </summary>
  /// <returns>リターンコード、83：フォルダ作成失敗</returns>
  static private int preProc()
  {
    try{
      // 出力先フォルダが存在しない場合、作成する
      if(!Directory.Exists(outPath))
      {
        Directory.CreateDirectory(outPath);
      }
    }
    catch(Exception e)
    {
      Console.WriteLine("出力フォルダの作成に失敗しました。");
      Console.WriteLine(e.Message);
      return 83;
    }
    return 0;
  }

  /// <summary>
  /// エクセル-CSV変換
  /// </summary>
  /// <param name="OutFiles">出力ファイルパス</param>
  /// <returns>リターンコード、98：ファイル入出力エラー、99：その他例外</returns>
  static private int excel2CSv(out List<string> OutFiles)
  {
    // 出力ファイルパスListの初期化
    OutFiles= new List<string>();

    try{
      // エクセルファイル読み込み
      using(FileStream fs = new FileStream(inFile, FileMode.Open, FileAccess.Read))
      {
        // エクセルブックオブジェクト
        IWorkbook excelBook = new XSSFWorkbook(fs);

        // エクセルファイル名prefix
        string prefix = System.IO.Path.GetFileNameWithoutExtension(inFile);

        // エクセルブックに含まれるシートを走査
        foreach(ISheet sheet in excelBook)
        {
          // 出力csvファイルパス
          string outFile = outPath + "\\" + prefix + "_" + sheet.SheetName + ".csv";

          using(StreamWriter sw = new StreamWriter(outFile, false, Encoding.UTF8))
          {
            // 行を走査
            foreach(IRow row in sheet)
            {
              StringBuilder lineBuilder = new StringBuilder();
              // 列（セル）を走査
              foreach(ICell cell in row)
              {
                lineBuilder.Append('"').Append(cell.ToString()).Append('"').Append(',');
              }

              // ファイル出力
              sw.WriteLine(lineBuilder.ToString().TrimEnd(','));
            }
          }
          // 出力先を格納
          OutFiles.Add(outFile);
        }
      }
    }
    catch(IOException ex)
    {
      Console.WriteLine("ファイル入出力に失敗しました");
      Console.WriteLine(ex.Message);
      return 98;
    }
    catch(Exception ex)
    {
      Console.WriteLine("Excel-CSV変換処理で例外が発生しました");
      Console.WriteLine(ex.Message);
      return 99;
    }

    return 0;
  }
}
