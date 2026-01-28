using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;

namespace ClearFileAssoc
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.AllFiles)]
    [Guid("F1E2D3C4-B5A6-7890-1234-FEDCBA098765")]
    public class ClearAssocHandler : SharpContextMenu
    {
        protected override bool CanShowMenu()
        {
            // Show menu for any file that has an extension
            return SelectedItemPaths.Any(path => !string.IsNullOrEmpty(Path.GetExtension(path)));
        }

        protected override ContextMenuStrip CreateMenu()
        {
            var menu = new ContextMenuStrip();

            // Get the extension of the first selected file
            string firstFile = SelectedItemPaths.FirstOrDefault();
            if (string.IsNullOrEmpty(firstFile))
                return menu;

            string extension = Path.GetExtension(firstFile);
            string currentAssoc = FileAssociationHelper.GetCurrentAssociation(extension);

            string menuText = string.IsNullOrEmpty(currentAssoc)
                ? $"Clear {extension} association"
                : $"Clear {extension} association ({currentAssoc})";

            var clearItem = new ToolStripMenuItem(menuText);
            clearItem.Click += ClearItem_Click;

            menu.Items.Add(clearItem);

            return menu;
        }

        private void ClearItem_Click(object sender, EventArgs e)
        {
            string firstFile = SelectedItemPaths.FirstOrDefault();
            if (string.IsNullOrEmpty(firstFile))
                return;

            string extension = Path.GetExtension(firstFile);

            var result = FileAssociationHelper.ClearAssociation(extension);

            if (result.Success)
            {
                MessageBox.Show(
                    $"{result.Message}\n\nThe file type will now use the system default or prompt you to choose an app.",
                    "Association Cleared",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(
                    result.Message,
                    "Clear Association Failed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
