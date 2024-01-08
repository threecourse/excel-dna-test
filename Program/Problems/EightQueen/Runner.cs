using System.Drawing;
using System.Linq;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using NLog;
using Program.Utils;

namespace Program.Problems.EightQueen;

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
        var ws = wb.Sheets["8Queen"] as Worksheet;

        var manager = new SheetDataLoader();
        var boardString = manager.ReadDataAsMatrix(ws.Range["Board"]);
        var hasQueen = boardString.Select2DArray(v => v == "Q");

        var inputData = new InputData
        {
            HasQueen = hasQueen
        };

        return inputData;
    }

    private void WriteResultToSheet(Workbook wb, InputData input, OutputData output)
    {
        var ws = wb.Sheets["8Queen"] as Worksheet;

        var errorQueens = output.ErrorQueens;

        // 色
        var colorError = Color.FromArgb(255, 91, 132);
        var colorSuccess = Color.FromArgb(64, 255, 64);
        var colorErrorList = errorQueens.Select(tpl => (tpl.y, tpl.x, color: colorError)).ToList();

        // メッセージ
        var messages = new string[5];
        for (var i = 0; i < messages.Length; i++)
            if (i < output.Messages.Count)
                messages[i] = output.Messages[i];
            else
                messages[i] = "";

        // ステータス
        var statuses = new string[1];
        statuses[0] = output.IsSolved ? "成功" : "失敗";

        // シートへの書き込み
        var writer = new SheetDataWriter();

        // 色のセット
        writer.ClearColor(ws.Range["Board"]);
        writer.SetColor(ws.Range["Board"], colorErrorList);
        if (output.IsSolved)
            writer.SetColor(ws.Range["Board"], colorSuccess);

        // メッセージのセット
        writer.ClearRange(ws.Range["Messages"]);
        writer.WriteRange(ws.Range["Messages"], 0, 0, messages.ToColumnMatrix());

        // ステータスのセット
        writer.ClearRange(ws.Range["Status"]);
        writer.WriteRange(ws.Range["Status"], 0, 0, statuses.ToColumnMatrix());
    }
}