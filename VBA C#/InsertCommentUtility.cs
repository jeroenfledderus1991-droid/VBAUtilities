using System;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace VBEAddIn
{
    /// <summary>
    /// Utility voor het invoegen van commentaren met gebruikersinfo en timestamp
    /// </summary>
    public static class InsertCommentUtility
    {
        [DllImport("user32.dll")]
        private static extern short GetKeyState(int nVirtKey);

        private const int VK_CONTROL = 0x11;
        private const int VK_SHIFT = 0x10;

        /// <summary>
        /// Voeg commentaar toe aan de huidige regel
        /// 
        /// Gebruik:
        /// - Normaal: voegt commentaar toe aan einde van regel: '20260212-1430 Gebruiker - 
        /// - SHIFT: voegt commentaar toe met asterisks: '20260212-1430 Gebruiker ***
        /// - CTRL: voegt START/END block toe rond geselecteerde regels
        /// 
        /// Als er al een commentaar met de gebruikersnaam op de regel staat, wordt deze verwijderd
        /// </summary>
        public static void Execute(VBE vbe)
        {
            try
            {
                if (vbe == null || vbe.ActiveCodePane == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Open eerst een code module in de VBA Editor.",
                        "Geen actieve code",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                // Check of gebruikersnaam is ingesteld
                if (string.IsNullOrWhiteSpace(FormatterSettings.CommentUserName))
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Stel eerst uw naam in via Utilities > Instellingen > Comments tab.",
                        "Gebruikersnaam niet ingesteld",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                var codeMod = vbe.ActiveCodePane.CodeModule;
                
                int startLine, startCol, endLine, endCol;
                vbe.ActiveCodePane.GetSelection(out startLine, out startCol, out endLine, out endCol);

                // Zorg ervoor dat we een geldige regel hebben
                if (startLine < 1) startLine = 1;
                if (startLine > codeMod.CountOfLines) startLine = codeMod.CountOfLines;
                
                string orgCodeLine = "";
                try
                {
                    orgCodeLine = codeMod.get_Lines(startLine, 1);
                }
                catch
                {
                    // Als het lezen faalt, gebruik lege string
                    orgCodeLine = "";
                }
                
                int currentCol = startCol;

                string userName = FormatterSettings.CommentUserName;
                string timestamp = DateTime.Now.ToString("yyyyMMdd-HHmm");

                // Check of er al een commentaar van deze gebruiker op de regel staat
                if (orgCodeLine.Contains("'") && orgCodeLine.Contains(userName))
                {
                    // Verwijder alles vanaf de eerste apostrof
                    int commentStart = orgCodeLine.IndexOf("'");
                    string newLine = orgCodeLine.Substring(0, commentStart).TrimEnd();
                    codeMod.ReplaceLine(startLine, newLine);
                    
                    // Cursor op veilige positie zetten (niet verder dan nieuwe regel lengte)
                    int safeCol = Math.Min(currentCol, newLine.Length + 1);
                    if (safeCol < 1) safeCol = 1;
                    vbe.ActiveCodePane.SetSelection(startLine, safeCol, startLine, safeCol);
                    return;
                }

                // Check welke modifier key is ingedrukt
                bool ctrlPressed = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
                bool shiftPressed = (GetKeyState(VK_SHIFT) & 0x8000) != 0;

                if (ctrlPressed)
                {
                    // CTRL: Voeg START/END block toe
                    InsertStartEndBlock(codeMod, vbe, startLine, endLine, userName, timestamp);
                }
                else if (shiftPressed)
                {
                    // SHIFT: Voeg commentaar met asterisks toe
                    InsertCommentWithAsterisks(codeMod, vbe, startLine, orgCodeLine, userName, timestamp);
                }
                else
                {
                    // Normaal: Voeg simpel commentaar toe
                    InsertSimpleComment(codeMod, vbe, startLine, orgCodeLine, userName, timestamp);
                }

                // Focus blijft in VBE (SetSelection heeft dit al gedaan)
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Fout bij invoegen commentaar: " + ex.Message,
                    "Fout",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Voeg START/END commentaar block toe (CTRL mode)
        /// </summary>
        private static void InsertStartEndBlock(CodeModule codeMod, VBE vbe, int startLine, int endLine, string userName, string timestamp)
        {
            string template = FormatterSettings.CommentTemplate;
            int lineLen = FormatterSettings.CommentLineLength;

            // Vervang placeholders in template voor START regel
            string startTemplate = template.Replace("{TIMESTAMP}", timestamp)
                                          .Replace("{USERNAME}", userName)
                                          .Replace("{TYPE}", "START");

            // Bereken aantal asterisks
            int asteriskCount = lineLen - startTemplate.Length;
            if (asteriskCount < 0) asteriskCount = 0;

            string startComment = "\t" + startTemplate.Replace("{FILLER}", new string('*', asteriskCount));

            // Vervang placeholders voor END regel
            string endTemplate = template.Replace("{TIMESTAMP}", timestamp)
                                        .Replace("{USERNAME}", userName)
                                        .Replace("{TYPE}", "END  ");

            asteriskCount = lineLen - endTemplate.Length;
            if (asteriskCount < 0) asteriskCount = 0;

            string endComment = "\t" + endTemplate.Replace("{FILLER}", new string('*', asteriskCount));

            // Voeg regels toe
            codeMod.InsertLines(startLine, startComment);
            codeMod.InsertLines(endLine + 2, endComment);

            // Positioneer cursor na de filler in de START regel
            int pipePos = startComment.IndexOf('|');
            if (pipePos > 0 && pipePos + 2 <= startComment.Length)
            {
                int cursorCol = Math.Min(pipePos + 3, startComment.Length + 1);
                vbe.ActiveCodePane.SetSelection(startLine, cursorCol, startLine, cursorCol);
            }
            else
            {
                vbe.ActiveCodePane.SetSelection(startLine, 1, startLine, 1);
            }
        }

        /// <summary>
        /// Voeg commentaar met asterisks toe (SHIFT mode)
        /// </summary>
        private static void InsertCommentWithAsterisks(CodeModule codeMod, VBE vbe, int startLine, string orgLine, string userName, string timestamp)
        {
            string template = FormatterSettings.CommentTemplateShift;
            int lineLen = FormatterSettings.CommentLineLength;

            // Vervang placeholders
            string comment = template.Replace("{TIMESTAMP}", timestamp)
                                    .Replace("{USERNAME}", userName);

            // Bereken aantal asterisks
            int totalLen = orgLine.Length + comment.Length;
            int asteriskCount = lineLen - totalLen;
            if (asteriskCount < 0) asteriskCount = 0;

            comment = comment.Replace("{FILLER}", new string('*', asteriskCount));

            string newLine = orgLine + "\t" + comment;
            codeMod.ReplaceLine(startLine, newLine);

            // Positioneer cursor na de username (let op: tab = 1 char, +2 voor spaties)
            int usernameStartInComment = comment.IndexOf(userName);
            if (usernameStartInComment >= 0)
            {
                // orgLine.Length + 1 (tab) + positie van username in comment + length van username + 2 (spatie na username)
                int cursorPos = orgLine.Length + 1 + usernameStartInComment + userName.Length + 2;
                cursorPos = Math.Min(cursorPos, newLine.Length + 1); // Zorg dat het binnen bereik blijft
                if (cursorPos > 1) cursorPos += 3; // 2 spaties extra naar voren
                vbe.ActiveCodePane.SetSelection(startLine, cursorPos, startLine, cursorPos);
            }
            else
            {
                // Fallback: einde van regel
                vbe.ActiveCodePane.SetSelection(startLine, newLine.Length + 1, startLine, newLine.Length + 1);
            }
        }

        /// <summary>
        /// Voeg simpel commentaar toe (normale mode)
        /// </summary>
        private static void InsertSimpleComment(CodeModule codeMod, VBE vbe, int startLine, string orgLine, string userName, string timestamp)
        {
            string template = FormatterSettings.CommentTemplateNormal;

            // Vervang placeholders
            string comment = template.Replace("{TIMESTAMP}", timestamp)
                                    .Replace("{USERNAME}", userName);

            string newLine = orgLine + "\t" + comment;
            codeMod.ReplaceLine(startLine, newLine);

            // Positioneer cursor 2 posities verder (na het template, klaar om te typen)
            // +1 voor VBE 1-based indexing, +2 voor extra spaties naar voren
            int cursorPos = newLine.Length + 1;
            if (cursorPos > 1) cursorPos += 3; // 2 spaties naar voren
            vbe.ActiveCodePane.SetSelection(startLine, cursorPos, startLine, cursorPos);
        }
    }
}
