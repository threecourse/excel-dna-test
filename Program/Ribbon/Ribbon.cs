using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using Program.Problems.IntroductionToHeuristics;

namespace Program.Ribbon;

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    public override string GetCustomUI(string RibbonID)
    {
        return @"
      <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
      <ribbon>
        <tabs>
          <tab id='t1' label='ExcelDnaTest'>
            <group id='g1' label='ExcelDnaTest'>
              <button id='b1' label='Run Intro-Heuristics' onAction='OnButtonPressedIntroHeuristics'/>
              <button id='b2' label='Run 8Queen' onAction='OnButtonPressed8Queen'/>
            </group >
          </tab>
        </tabs>
      </ribbon>
    </customUI>";
    }

    public void OnButtonPressedIntroHeuristics(IRibbonControl control)
    {
        var runner = new Runner();
        runner.Run();
    }

    public void OnButtonPressed8Queen(IRibbonControl control)
    {
        var runner = new Problems.EightQueen.Runner();
        runner.Run();
    }
}