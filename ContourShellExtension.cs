using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.IO.Compression;

namespace ContourShellExtension
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.ClassOfExtension, ".pbpp")]
    public class ContourShellExtension : SharpContextMenu
    {
        /// <summary>
        /// Determines whether this instance can a shell context show menu, given the specified selected file list.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this instance should show a shell context menu for the specified file list; otherwise, <c>false</c>.
        /// </returns>
        protected override bool CanShowMenu()
        {
            //  We always show the menu.
            return true;
        }

        /// <summary>
        /// Creates the context menu. This can be a single menu item or a tree of them.
        /// </summary>
        /// <returns>
        /// The context menu for the shell context menu.
        /// </returns>
        protected override ContextMenuStrip CreateMenu()
        {
            //  Create the menu strip.
            var menu = new ContextMenuStrip();

            //  Create a 'count lines' item.
            var itemCountLines = new ToolStripMenuItem
            {
                Text = "SAFAN PBPP Rename",
                Image = Properties.Resources.Rename
            };

            //  When we click, we'll count the lines.
            itemCountLines.Click += (sender, args) => this.RenamePbppFile();

            //  Add the item to the context menu.
            menu.Items.Add(itemCountLines);

            //  Return the menu.
            return menu;
        }

        /// <summary>
        /// Counts the lines in the selected files.
        /// </summary>
        private void RenamePbppFile()
        {
            if (this.SelectedItemPaths.Count() != 1)
            {
                MessageBox.Show("Maximaal en minimaal 1 bestand selecteren", "Fout", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                var oldFilePath = this.SelectedItemPaths.First();
                string newBaseFileName = Interaction.InputBox("Nieuwe bestandsnaam (zonder extensie)", "Invoerveld", Path.GetFileNameWithoutExtension(oldFilePath), -1, -1);
                if (!string.IsNullOrWhiteSpace(newBaseFileName))
                {
                    // Bestandsnaam hernoemen
                    var path = Path.GetDirectoryName(oldFilePath);
                    var ext = Path.GetExtension(oldFilePath);
                    var newFilePath = path + @"\" + newBaseFileName + ext;

                    if (File.Exists(newFilePath))
                    {
                        MessageBox.Show($"Fout bij hernoemen van {Path.GetFileName(oldFilePath)} naar {Path.GetFileName(newFilePath)} {Environment.NewLine}Nieuwe bestandsnaam bestaat al.");
                    }
                    else
                    {

                        try
                        {
                            File.Copy(oldFilePath, newFilePath);
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show($"Fout bij kopieren van {Path.GetFileName(oldFilePath)} naar  {Path.GetFileName(newFilePath)} {Environment.NewLine}Fout: {e.Message}");
                            throw;
                        }

                        // Inhoud hernoemen
                        try
                        {
                            using (var file = File.Open(newFilePath, FileMode.Open))
                            using (var zip = new ZipArchive(file, ZipArchiveMode.Update))
                            {
                                zip.RenameEntries(Path.GetFileNameWithoutExtension(oldFilePath), newBaseFileName);
                            }
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show($"Fout bij wijzigen inhoud PBPP file {Path.GetFileName(oldFilePath)}{Environment.NewLine}Fout: {e.Message}");
                            throw;
                        }

                        File.Delete(oldFilePath);
                        MessageBox.Show("Bestandsnaam van PBPP, en bestandsnamen in inhoud met zelfde bestandsnaam zijn hernoemd.");

                    }
                }
            }
        }
    }

    public static class ExtensionClasses
    {
        public static void RenameEntries(this ZipArchive archive, string archiveBaseFilename, string newBaseFileName)
        {
            foreach (var oldEntry in archive.Entries.ToList())
            {
                var entryPath = Path.GetDirectoryName(oldEntry.FullName);
                if (!string.IsNullOrWhiteSpace(entryPath))
                {
                    entryPath += @"\";
                }
                var entryExt = Path.GetExtension(oldEntry.Name);
                var entryBaseFilename = Path.GetFileNameWithoutExtension(oldEntry.Name);

                if (entryBaseFilename.Equals(archiveBaseFilename, StringComparison.OrdinalIgnoreCase) || entryExt.Equals(".aif", StringComparison.OrdinalIgnoreCase) || entryExt.Equals(".agx", StringComparison.OrdinalIgnoreCase) || entryExt.Equals(".pbp", StringComparison.OrdinalIgnoreCase))
                {
                    var newEntry = archive.CreateEntry(entryPath + newBaseFileName + entryExt);

                    using (Stream oldStream = oldEntry.Open())
                    using (Stream newStream = newEntry.Open())
                    {
                        oldStream.CopyTo(newStream);
                    }

                    oldEntry.Delete();
                }
            }
        }
    }
}