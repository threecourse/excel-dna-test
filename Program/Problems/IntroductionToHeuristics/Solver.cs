using System;

namespace Program.Problems.IntroductionToHeuristics;

internal class Solver
{
    public OutputData Solve(InputData input)
    {
        // NOTE: Solverはめちゃくちゃ適当。ランダムに答えを設定するのみ
        var rnd = new Random();

        var days = input.D.GetLength(0);
        var cs = input.C.Length;
        var score = 999999;

        var selected = new int[days];
        for (var d = 0; d < days; d++) selected[d] = rnd.Next(cs);

        return new OutputData
        {
            Score = score,
            Selected = selected
        };
    }
}