using System;
using System.Collections.Generic;

namespace Program.Utils;

public static class Extension
{
    public static int Sum2DArray(this int[,] array)
    {
        // 2次元配列の合計を求める
        var sum = 0;

        var h = array.GetLength(0);
        var w = array.GetLength(1);
        for (var y = 0; y < h; y++)
        for (var x = 0; x < w; x++)
            sum += array[y, x];

        return sum;
    }

    public static List<T> Flatten2DArray<T>(this T[,] array)
    {
        var ret = new List<T>();
        var h = array.GetLength(0);
        var w = array.GetLength(1);
        for (var y = 0; y < h; y++)
        for (var x = 0; x < w; x++)
            ret.Add(array[y, x]);

        return ret;
    }

    public static TTo[,] Select2DArray<TFrom, TTo>(this TFrom[,] array, Func<TFrom, TTo> source)
    {
        var h = array.GetLength(0);
        var w = array.GetLength(1);
        var ret = new TTo[h, w];
        for (var y = 0; y < h; y++)
        for (var x = 0; x < w; x++)
            ret[y, x] = source(array[y, x]);

        return ret;
    }

    public static object[,] ToRowMatrix<T>(this IList<T> array)
    {
        var h = 1;
        var w = array.Count;
        var ret = new object[h, w];
        for (var i = 0; i < array.Count; i++)
        {
            var y = 0;
            var x = i;
            ret[y, x] = array[i];
        }

        return ret;
    }

    public static object[,] ToColumnMatrix<T>(this IList<T> array)
    {
        var h = array.Count;
        var w = 1;
        var ret = new object[h, w];
        for (var i = 0; i < array.Count; i++)
        {
            var y = i;
            var x = 0;
            ret[y, x] = array[i];
        }

        return ret;
    }
}