using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace Program.Utils;

public class SheetDataLoader
{
    /// <summary>
    ///     シートを各列がintであるデータフレームとして返す
    /// </summary>
    public List<(string Header, int[] Values)> ReadIntDataFrame(Worksheet ws)
    {
        // TODO: 例外処理は省略
        var matRaw = ReadDataAsMatrix(ws);

        var h = matRaw.GetLength(0);
        var w = matRaw.GetLength(1);

        var ret = new List<(string, int[])>();

        for (var x = 0; x < w; x++)
        {
            var header = matRaw[0, x];
            var values = Enumerable.Range(1, h - 1).Select(y => int.Parse(matRaw[y, x])).ToArray();
            ret.Add((header, values));
        }

        return ret;
    }

    /// <summary>
    ///     シートの存在する範囲の行・列のうち、最も外側の行を除き、Integerとして取得する
    /// </summary>
    public int[,] ReadIntMatrix(Worksheet ws)
    {
        // TODO: 例外処理は省略
        var matRaw = ReadDataAsMatrix(ws);

        var hRaw = matRaw.GetLength(0);
        var wRaw = matRaw.GetLength(1);

        // 最も外側の行を除く
        var h = hRaw - 1;
        var w = wRaw - 1;

        var mat = new int[h, w];
        for (var y = 0; y < hRaw; y++)
        for (var x = 0; x < wRaw; x++)
        {
            if (y == 0 || x == 0)
                continue;
            mat[y - 1, x - 1] = int.Parse(matRaw[y, x]);
        }

        return mat;
    }

    /// <summary>
    ///     存在する値までの行もしくは列の長さを取得する
    /// </summary>
    private int GetRowOrColumnCount(Worksheet ws, int max, bool isVertical)
    {
        for (var i = 0; i < max; i++)
        {
            var y = isVertical ? i + 1 : 1;
            var x = isVertical ? 1 : i + 1;
            var v = ws.Cells[y, x].Value2();
            if (i != 0 && v == null)
                return i;
        }

        return max;
    }

    /// <summary>
    ///     シートの存在する範囲の行・列をstringとして取得する
    /// </summary>
    public string[,] ReadDataAsMatrix(Worksheet ws)
    {
        // 列・行の上限を限定している
        const int yMax = 10000;
        const int xMax = 256;

        // 値が存在する場所までの高さ・幅を取る
        var h = GetRowOrColumnCount(ws, yMax, true);
        var w = GetRowOrColumnCount(ws, xMax, false);
        object[,] matLoad = ws.Range["A1"].Resize[h, w].Value;

        var mat = new string[h, w];
        for (var y = 0; y < h; y++)
        for (var x = 0; x < w; x++)
        {
            // なぜか1-indexedの配列で取得できる
            var v = matLoad[y + 1, x + 1];
            if (v != null)
                mat[y, x] = v.ToString();
            else
                mat[y, x] = string.Empty;
        }

        return mat;
    }

    public string[,] ReadDataAsMatrix(Range range)
    {
        object[,] matRaw = range.Value;

        var hRaw = matRaw.GetLength(0);
        var wRaw = matRaw.GetLength(1);
        var h = hRaw;
        var w = wRaw;

        var mat = new string[h, w];
        for (var y = 0; y < h; y++)
        for (var x = 0; x < w; x++)
        {
            // なぜか1-indexedの配列で取得できる
            var v = matRaw[y + 1, x + 1];
            if (v != null)
                mat[y, x] = v.ToString();
            else
                mat[y, x] = string.Empty;
        }

        return mat;
    }
}