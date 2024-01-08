using System.Collections.Generic;
using System.Drawing;
using Microsoft.Office.Interop.Excel;

namespace Program.Utils;

public class SheetDataWriter
{
    #region 色の設定

    public void ClearColor(Range range)
    {
        range.Interior.ColorIndex = 0;
    }

    #endregion

    #region 値の書き込み

    /// <summary>
    ///     データをワークシートに書き込む
    ///     startY, startXは0-indexed
    /// </summary>
    public void WriteMatrix(Worksheet ws, int startY, int startX, object[,] data)
    {
        var h = data.GetLength(0);
        var w = data.GetLength(1);
        ws.Range["A1"].Offset[startY, startX].Resize[h, w].Value = data;
    }

    public void WriteRange(Range range, int startY, int startX, object[,] data)
    {
        // TODO: セルの操作中に処理が実行されると例外が発生する
        var h = data.GetLength(0);
        var w = data.GetLength(1);
        range.Offset[startY, startX].Resize[h, w].Value = data;
    }

    #endregion

    #region データのクリア

    /// <summary>
    ///     データをワークシートからクリアする
    /// </summary>
    public void ClearSheet(Worksheet ws, int h = 10000, int w = 256)
    {
        ws.Range["A1"].Resize[h, w].ClearContents();
    }

    /// <summary>
    ///     範囲をクリアする
    /// </summary>
    public void ClearRange(Range range)
    {
        range.ClearContents();
    }

    #endregion

    #region 色の設定

    /// <summary>
    ///     セルの色をワークシートに設定する
    /// </summary>
    public void SetColor(Worksheet ws, List<(int y, int x, Color color)> list)
    {
        foreach (var (y, x, color) in list)
        {
            var range = ws.Cells[1 + y, 1 + x];
            range.Interior.Color = color;
        }
    }

    /// <summary>
    ///     範囲の色を修正する
    /// </summary>
    public void SetColor(Range range, List<(int y, int x, Color color)> list)
    {
        foreach (var (y, x, color) in list) range[1 + y, 1 + x].Interior.Color = color;
    }

    /// <summary>
    ///     範囲の色を修正する
    /// </summary>
    public void SetColor(Range range, Color color)
    {
        range.Interior.Color = color;
    }

    #endregion
}