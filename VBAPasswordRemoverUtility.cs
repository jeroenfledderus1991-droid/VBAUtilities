using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;

namespace VBEAddIn
{
    /// <summary>
    /// Verwijdert VBA project wachtwoorden door de DPB-hash in vbaProject.bin te patchen.
    /// Het bestand moet GESLOTEN zijn in Excel voordat dit werkt.
    /// Na heropenen vraagt Excel om het project te resetten (klik Ja).
    /// </summary>
    public static class VBAPasswordRemoverUtility
    {
        public static void Execute(VBE vbe)
        {
            string suggestedPath = "";
            try
            {
                if (vbe?.ActiveVBProject != null)
                    suggestedPath = vbe.ActiveVBProject.FileName;
            }
            catch { }

            var warn = MessageBox.Show(
                "LET OP: Het bestand moet GESLOTEN zijn in Excel voordat het wachtwoord verwijderd kan worden.\n\n" +
                (string.IsNullOrEmpty(suggestedPath) ? "" : "Actief project:\n" + suggestedPath + "\n\n") +
                "Sluit het bestand in Excel, klik daarna OK en selecteer het bestand.",
                "VBA Wachtwoord Verwijderen",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Warning);

            if (warn != DialogResult.OK) return;

            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Title = "Selecteer het Excel bestand";
                dlg.Filter = "Excel bestanden (*.xlsm;*.xlam;*.xltm;*.xlsb;*.xls;*.xla)|*.xlsm;*.xlam;*.xltm;*.xlsb;*.xls;*.xla|Alle bestanden (*.*)|*.*";

                if (!string.IsNullOrEmpty(suggestedPath) && File.Exists(suggestedPath))
                {
                    dlg.InitialDirectory = Path.GetDirectoryName(suggestedPath);
                    dlg.FileName = Path.GetFileName(suggestedPath);
                }

                if (dlg.ShowDialog() != DialogResult.OK) return;

                RemovePassword(dlg.FileName);
            }
        }

        private static void RemovePassword(string filePath)
        {
            try
            {
                string ext = Path.GetExtension(filePath).ToLower();
                string backupPath = filePath + ".bak";
                File.Copy(filePath, backupPath, true);

                if (ext == ".xlsm" || ext == ".xlam" || ext == ".xltm" || ext == ".xlsb")
                    RemovePasswordZip(filePath);
                else
                    RemovePasswordRawBinary(filePath);

                MessageBox.Show(
                    "Wachtwoord succesvol verwijderd!\n\n" +
                    "Backup opgeslagen als:\n" + backupPath + "\n\n" +
                    "Open het bestand nu weer in Excel.\n" +
                    "Als Excel vraagt of het VBA project gereset moet worden, klik dan op 'Ja'.",
                    "Geslaagd",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Fout bij verwijderen wachtwoord:\n\n" + ex.Message,
                    "Fout",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Voor ZIP-gebaseerde bestanden (.xlsm, .xlam, .xlsb): patch vbaProject.bin in het archief.
        /// </summary>
        private static void RemovePasswordZip(string filePath)
        {
            string tempPath = filePath + ".tmp";
            try
            {
                using (var inputStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                using (var inputArchive = new ZipArchive(inputStream, ZipArchiveMode.Read))
                using (var outputStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
                using (var outputArchive = new ZipArchive(outputStream, ZipArchiveMode.Create))
                {
                    foreach (ZipArchiveEntry entry in inputArchive.Entries)
                    {
                        byte[] data;
                        using (var entryStream = entry.Open())
                        using (var ms = new MemoryStream())
                        {
                            entryStream.CopyTo(ms);
                            data = ms.ToArray();
                        }

                        if (entry.FullName == "xl/vbaProject.bin")
                            data = PatchDpb(data);

                        ZipArchiveEntry newEntry = outputArchive.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                        newEntry.LastWriteTime = entry.LastWriteTime;
                        using (var newEntryStream = newEntry.Open())
                            newEntryStream.Write(data, 0, data.Length);
                    }
                }

                File.Delete(filePath);
                File.Move(tempPath, filePath);
            }
            catch
            {
                if (File.Exists(tempPath)) File.Delete(tempPath);
                throw;
            }
        }

        /// <summary>
        /// Voor binaire .xls/.xla bestanden: zoek DPB= direct in de bestandsbytes.
        /// </summary>
        private static void RemovePasswordRawBinary(string filePath)
        {
            byte[] data = File.ReadAllBytes(filePath);
            data = PatchDpb(data);
            File.WriteAllBytes(filePath, data);
        }

        /// <summary>
        /// Vervangt "DPB=" door "DPx=" zodat Excel de wachtwoord-hash negeert.
        /// </summary>
        private static byte[] PatchDpb(byte[] data)
        {
            // Search for: DPB="
            byte[] search = Encoding.ASCII.GetBytes("DPB=\"");

            for (int i = 0; i <= data.Length - search.Length; i++)
            {
                bool match = true;
                for (int j = 0; j < search.Length; j++)
                {
                    if (data[i + j] != search[j]) { match = false; break; }
                }
                if (match)
                {
                    data[i + 2] = (byte)'x'; // DPB= → DPx=
                    return data;
                }
            }

            throw new Exception(
                "Geen DPB-wachtwoordhash gevonden in het bestand.\n\n" +
                "Mogelijke oorzaken:\n" +
                "• Het bestand is niet VBA-wachtwoord beveiligd\n" +
                "• Het bestandsformaat wordt niet ondersteund");
        }
    }
}
