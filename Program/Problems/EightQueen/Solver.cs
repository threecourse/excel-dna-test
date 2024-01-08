using System;
using System.Collections.Generic;
using System.Linq;
using Program.Utils;

namespace Program.Problems.EightQueen;

internal class Solver
{
    public OutputData Solve(InputData input)
    {
        var hasQueen = input.HasQueen;

        var count = hasQueen.Flatten2DArray().Sum(x => x ? 1 : 0);
        var h = hasQueen.GetLength(0);
        var w = hasQueen.GetLength(1);

        var errorQueens = new List<(int y, int x)>();
        for (var y = 0; y < h; y++)
        for (var x = 0; x < w; x++)
            if (hasQueen[y, x])
                if (HasOtherQueen(y, x, hasQueen))
                    errorQueens.Add((y, x));

        var success = count == 8 && errorQueens.Count == 0;

        var messages = new List<string>();
        if (count < 8)
            messages.Add("クイーンの数が8個より少ないです");
        if (count > 8)
            messages.Add("クイーンの数が8個より多いです");
        if (errorQueens.Count > 0)
            messages.Add("条件を満たさないクイーンがあります");

        return new OutputData
        {
            IsSolved = success,
            Messages = messages,
            ErrorQueens = errorQueens
        };
    }

    private bool HasOtherQueen(int y0, int x0, bool[,] hasQueen)
    {
        var h = hasQueen.GetLength(0);
        var w = hasQueen.GetLength(1);

        // 縦
        for (var y = 0; y < h; y++)
        {
            if (y == y0) continue;
            if (hasQueen[y, x0]) return true;
        }

        // 横
        for (var x = 0; x < w; x++)
        {
            if (x == x0) continue;
            if (hasQueen[y0, x]) return true;
        }

        // 斜め
        for (var y = 0; y < h; y++)
        {
            if (y == y0) continue;
            var dy = Math.Abs(y - y0);

            var xl = x0 - dy;
            var xr = x0 + dy;
            if (xl >= 0 && xl < w && hasQueen[y, xl]) return true;
            if (xr >= 0 && xr < w && hasQueen[y, xr]) return true;
        }

        return false;
    }
}