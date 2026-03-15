using System;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;

namespace Testing.ExcelDnaAddIn
{
    public class RibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='TestingTab' label='Testing DNA'>
        <group id='TestingGroup' label='Test Actions'>
          <button id='ShowJsPaneButton'
                  label='Show JS Pane'
                  size='large'
                  imageMso='HappyFace'
                  onAction='OnShowJsPaneClicked' />
          <button id='HideJsPaneButton'
                  label='Hide JS Pane'
                  size='large'
                  imageMso='SadFace'
                  onAction='OnHideJsPaneClicked' />
          <button id='DumpUiAutomationButton'
                  label='Dump UIA Tree'
                  size='large'
                  imageMso='ReviewingPane'
                  onAction='OnDumpUiAutomationClicked' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnShowJsPaneClicked(object control)
        {
            TryRunAutomation(
                OfficeJsPaneAutomation.ShowPane,
                "The Office JS pane show command was sent via UI Automation.");
        }

        public void OnHideJsPaneClicked(object control)
        {
            TryRunAutomation(
                OfficeJsPaneAutomation.HidePane,
                "The Office JS pane hide command was sent via UI Automation.");
        }

        public void OnDumpUiAutomationClicked(object control)
        {
            TryRunAutomation(() =>
            {
                var outputPath = OfficeJsPaneAutomation.DumpTreeToTextFile();
                OfficeJsPaneAutomation.DumpTreeToWorksheet(outputPath);
                MessageBox.Show(
                    $"UI Automation tree dumped to:{Environment.NewLine}{outputPath}{Environment.NewLine}{Environment.NewLine}A worksheet was also added to the active workbook.",
                    "Testing Excel-DNA",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            });
        }

        private static void TryRunAutomation(Action action, string successMessage = null)
        {
            try
            {
                action();
                if (!string.IsNullOrEmpty(successMessage))
                {
                    MessageBox.Show(successMessage, "Testing Excel-DNA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Testing Excel-DNA", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
