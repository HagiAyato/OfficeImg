using System;
using System.IO;
using System.IO.Compression;

/// <summary>
/// 
/// </summary>
namespace OfficeImg
{
    /// <summary>
    /// 
    /// </summary>
    class Program
    {
        /// <summary>
        /// メイン関数
        /// </summary>
        /// <param name="args">コマンドライン引数</param>
        static void Main(string[] args)
        {
            try
            {
                // 引数確認
                if (args.Length < 1)
                {
                    Console.WriteLine("画像を抜き出したいファイルを、本実行ファイルにドラッグ&ドロップしてください。");
                    Console.WriteLine("対象のファイル：word,excel,power pointのうち、Office2007以降のファイル");
                    Console.WriteLine("　　　　　　　　Open Document形式のファイル");
                    return;
                }
                foreach (string fname in args)
                {
                    Console.WriteLine("読み込みファイル：{0}", fname);
                    // ファイルの存在チェック
                    if (!File.Exists(fname))
                    {
                        Console.WriteLine("ファイルが存在しません。");
                        continue;
                    }
                    // ファイルの拡張子チェック
                    string type = "", extension = "";
                    extension = Path.GetExtension(fname);
                    switch (extension)
                    {
                        case ".docx":
                        case ".docm":
                        case ".dotx":
                        case ".dotm":
                            type = "word";
                            break;
                        case ".pptx":
                        case ".pptm":
                        case ".ppsx":
                        case ".ppsm":
                        case ".potx":
                        case ".potm":
                            type = "ppt";
                            break;
                        case ".xlsx":
                        case ".xlsm":
                        case ".xltx":
                        case ".xltm":
                            type = "xl";
                            break;
                        case ".ods":
                        case ".odt":
                        case ".odp":
                            type = "";
                            break;
                        default:
                            Console.WriteLine("対応のファイルではありません。");
                            continue;
                    }
                    Console.WriteLine("読み込みファイルの拡張子：{0}", extension);
                    // ファイルを展開
                    FileInfo info = new System.IO.FileInfo(fname);
                    if (info.Length <= 0)
                    {
                        Console.WriteLine("ファイルサイズ0Byte、このため中に画像はありません。");
                        continue;
                    }
                    // ディレクトリ名
                    string directoryName = Path.GetDirectoryName(fname);
                    // ファイル名拡張子なし
                    string fileName = Path.GetFileNameWithoutExtension(fname);
                    string unzipDir = directoryName + "/" + fileName + "_" + type;
                    // 既に同名ディレクトリがあれば削除
                    if (Directory.Exists(unzipDir)) Directory.Delete(unzipDir, true);
                    ZipFile.ExtractToDirectory(fname, unzipDir);
                    // 解凍したフォルダから画像ファイルを取り出す
                    string mediaDir = unzipDir + "/" + type + "/media";
                    // 画像なしなら処理打ち切り
                    if (!Directory.Exists(mediaDir))
                    {
                        Console.WriteLine("ファイル内に画像はありません。");
                        Directory.Delete(unzipDir, true);
                        continue;
                    }
                    DirectoryUtil.Copy(mediaDir, directoryName + "/" + fileName + extension + "_media");
                    // zipファイルと解凍したフォルダ削除
                    Directory.Delete(unzipDir, true);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("例外発生：{0}", e.ToString());
            }
            finally
            {

            }
        }
    }
}
