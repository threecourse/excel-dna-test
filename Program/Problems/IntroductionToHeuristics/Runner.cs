using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using NLog;
using Program.Utils;

namespace Program.Problems.IntroductionToHeuristics;

public class Runner
{
    private static readonly Logger logger = LogManager.GetCurrentClassLogger();

    public void Run()
    {
        var xlApp = (Application)ExcelDnaUtil.Application;
        var wb = xlApp.ActiveWorkbook;

        var inputData = LoadInputDataFromSheet(wb);
        var outputData = new Solver().Solve(inputData);
        WriteResultToSheet(wb, inputData, outputData);
    }

    private InputData LoadInputDataFromSheet(Workbook wb)
    {
        var wsInputC = wb.Sheets["Input-C"] as Worksheet;
        var wsInputD = wb.Sheets["Input-D"] as Worksheet;

        var manager = new SheetDataLoader();
        var D = manager.ReadIntMatrix(wsInputD);
        var C = manager.ReadIntDataFrame(wsInputC).First(r => r.Header == "C").Values;

        logger.Info($"L:[{D.GetLength(0)}, {D.GetLength(1)}], sum:{D.Sum2DArray()}");
        logger.Info($"L:{C.Length}, sum:{C.Sum()}");

        var inputData = new InputData
        {
            D = D,
            C = C
        };

        return inputData;
    }

    private void WriteResultToSheet(Workbook wb, InputData input, OutputData output)
    {
        var wsResultMatrix = wb.Sheets["Result-Matrix"] as Worksheet;
        var wsResultScore = wb.Sheets["Result-Score"] as Worksheet;

        var days = input.D.GetLength(0);
        var w = input.D.GetLength(1);

        // NOTE: object型で良いのかは不明

        // Result-Matrixシートのヘッダ
        var matrixSheetHeader = new object[w];
        for (var x = 0; x < w; x++) matrixSheetHeader[x] = $"C{x + 1}";

        // Result-Matrixシートのインデックス
        var matrixSheetIndex = new object[days];
        for (var y = 0; y < days; y++) matrixSheetIndex[y] = $"D{y}";

        // Result-Matrixシートの入力値
        var matrixSheetData = new object[days, w];
        for (var y = 0; y < days; y++)
        for (var x = 0; x < w; x++)
            matrixSheetData[y, x] = input.D[y, x];
        // Result-Matrixシートの色のセット
        var color = Color.FromArgb(255, 91, 132);
        var colorList = new List<(int y, int x, Color color)>();
        for (var d = 0; d < days; d++)
        {
            // ヘッダの分ずらす
            var y = d + 1;
            var x = output.Selected[d] + 1;
            colorList.Add((y, x, color));
        }

        // Result-Scoreシートの入力値
        var scoreSheetDataList = new List<(string Key, object Value)>();
        scoreSheetDataList.Add(("Score", output.Score));
        for (var d = 0; d < days; d++) scoreSheetDataList.Add(($"D{d}", output.Selected[d]));

        var scoreSheetData = new object[scoreSheetDataList.Count, 2];
        for (var i = 0; i < scoreSheetDataList.Count; i++)
        {
            scoreSheetData[i, 0] = scoreSheetDataList[i].Key;
            scoreSheetData[i, 1] = scoreSheetDataList[i].Value;
        }

        // 書き込み
        var writer = new SheetDataWriter();
        writer.ClearSheet(wsResultMatrix);
        writer.WriteMatrix(wsResultMatrix, 0, 1, matrixSheetHeader.ToRowMatrix());
        writer.WriteMatrix(wsResultMatrix, 1, 0, matrixSheetIndex.ToColumnMatrix());
        writer.WriteMatrix(wsResultMatrix, 1, 1, matrixSheetData);
        writer.SetColor(wsResultMatrix, colorList);
        writer.WriteMatrix(wsResultScore, 0, 0, scoreSheetData);
    }
}