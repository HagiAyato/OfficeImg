using System.IO;
/// <summary>
/// ディレクトリ操作自作クラス
/// 参考：https://kan-kikuchi.hatenablog.com/entry/DirectoryProcessor
/// </summary>
public static class DirectoryUtil
{
    /// <summary>
    /// ディレクトリとその中身を、一括削除
    /// </summary>
    /// <param name="directory">処理対象ディレクトリ</param>
    public static void Delete(string directory)
    {
        // ディレクトリ存在チェック
        if (!Directory.Exists(directory)) return;
        // ディレクトリ内のファイル全削除
        foreach (string filePath in Directory.GetFiles(directory))
        {
            File.SetAttributes(filePath, FileAttributes.Normal);
            File.Delete(filePath);
        }
        // ディレクトリ内のフォルダも削除(再帰)
        foreach (string dirPath in Directory.GetDirectories(directory))
        {
            Delete(dirPath);
        }
        // 最後に、ディレクトリ自身を削除
        Directory.Delete(directory, false);
    }

    /// <summary>
    /// ディレクトリとその中身をコピー
    /// </summary>
    /// <param name="sourcePath">コピー元ディレクトリ</param>
    /// <param name="copyPath">コピー先ディレクトリ</param>
    /// <param name="enableOverWrite">上書きを許可するか</param>
    public static void Copy(string sourcePath, string copyPath, bool enableOverWrite = true)
    {
        // 既にディレクトリがあるかチェック
        if (Directory.Exists(copyPath))
        {
            if (!enableOverWrite)
            {
                throw new IOException("ディレクトリが既に存在します。");
            }
            // ディレクトリを一度消す
            Delete(copyPath);
            // Directory.Delete(copyPath, true);
        }
        // コピー先ディレクトリ作成
        Directory.CreateDirectory(copyPath);
        // ディレクトリ内のファイルをコピー
        foreach (string filePath in Directory.GetFiles(sourcePath))
        {
            File.Copy(filePath, Path.Combine(copyPath, Path.GetFileName(filePath)));
        }
        // ディレクトリ内のフォルダもコピー(再帰)
        foreach (string dirPath in Directory.GetDirectories(sourcePath))
        {
            Copy(dirPath, Path.Combine(copyPath, Path.GetFileName(dirPath)));
        }
    }
}