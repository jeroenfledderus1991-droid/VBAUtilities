using System;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;

namespace VBEAddIn
{
    [ComVisible(true)]
    [Guid("E1F2A3B4-C5D6-4E78-F901-A234B5678C90")]
    [ProgId("VBEAddIn.CompleteCodeFormatter")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class CompleteCodeFormatter
    {
        private static string IndentUnit =>
            FormatterSettings.IndentType == "tabs"
                ? "\t"
                : new string(' ', Math.Max(1, FormatterSettings.IndentSize));

        /// <summary>
        /// Formatteert alleen de procedure waar de cursor in staat.
        /// Toont eerst een bevestigingsdialog met naam en scope.
        /// </summary>
        public string FormatProcedure(VBE vbe)
        {
            if (vbe == null || vbe.ActiveCodePane == null)
                return "Geen actieve code module.";

            var codePane = vbe.ActiveCodePane;
            var codeModule = codePane.CodeModule;

            int cursorLine, dummy1, dummy2, dummy3;
            codePane.GetSelection(out cursorLine, out dummy1, out dummy2, out dummy3);

            int totalLines = codeModule.CountOfLines;
            int procStart = -1, procEnd = -1;
            string procName = "Onbekend";

            for (int i = cursorLine; i >= 1; i--)
            {
                string upper = codeModule.Lines[i, 1].Trim().ToUpper();
                if (Regex.IsMatch(upper, @"^(PUBLIC|PRIVATE|FRIEND)?\s*(SUB|FUNCTION|PROPERTY)\s+"))
                {
                    var m = Regex.Match(codeModule.Lines[i, 1].Trim(),
                        @"(?:Public|Private|Friend)?\s*(?:Sub|Function|Property(?:\s+\w+)?)\s+(\w+)",
                        RegexOptions.IgnoreCase);
                    if (m.Success) procName = m.Groups[1].Value;
                    procStart = i;
                    break;
                }
            }

            if (procStart < 0)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Cursor staat niet in een procedure.",
                    "Geen procedure",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
                return "Cursor staat niet in een procedure.";
            }

            for (int i = cursorLine; i <= totalLines; i++)
            {
                string upper = codeModule.Lines[i, 1].Trim().ToUpper();
                if (Regex.IsMatch(upper, @"^END\s+(SUB|FUNCTION|PROPERTY)"))
                {
                    procEnd = i;
                    break;
                }
            }

            if (procEnd < 0)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Einde van procedure niet gevonden.",
                    "Fout",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return "Einde van procedure niet gevonden.";
            }

            string moduleName = codeModule.Name;
            int procCount = procEnd - procStart + 1;

            var confirm = System.Windows.Forms.MessageBox.Show(
                string.Format("Procedure '{0}' in module '{1}' formatteren?\n\nScope: procedure ({2} regels)",
                    procName, moduleName, procCount),
                "Bevestig: Procedure",
                System.Windows.Forms.MessageBoxButtons.OKCancel,
                System.Windows.Forms.MessageBoxIcon.Question);

            if (confirm != System.Windows.Forms.DialogResult.OK)
                return "Geannuleerd.";

            var lines = new List<CodeLine>();
            for (int i = procStart; i <= procEnd; i++)
                lines.Add(new CodeLine { LineNumber = i, OriginalText = codeModule.Lines[i, 1] });

            SortDeclarationsInProcedures(lines);
            RemoveExcessiveBlankLines(lines);
            AddMissingBlankLines(lines);
            FormatIndentation(lines);

            foreach (var line in lines)
                if (line.NewText == null && !line.MarkedForDeletion)
                    line.NewText = line.OriginalText;
            lines.RemoveAll(l => l.MarkedForDeletion);

            int changes = 0;
            codeModule.DeleteLines(procStart, procCount);
            for (int i = 0; i < lines.Count; i++)
            {
                codeModule.InsertLines(procStart + i, lines[i].NewText);
                if (lines[i].NewText != lines[i].OriginalText) changes++;
            }

            return string.Format("Procedure '{0}' geformatteerd!\n\n{1} regels aangepast\n{2} totale regels",
                procName, changes, lines.Count);
        }

        /// <summary>
        /// Formatteert de gehele huidige module.
        /// Toont eerst een bevestigingsdialog met naam en scope.
        /// </summary>
        public string FormatModule(CodeModule codeModule)
        {
            if (codeModule == null)
                return "Geen code module geselecteerd.";

            var confirm = System.Windows.Forms.MessageBox.Show(
                string.Format("Module '{0}' formatteren?\n\nScope: gehele module ({1} regels)",
                    codeModule.Name, codeModule.CountOfLines),
                "Bevestig: Module",
                System.Windows.Forms.MessageBoxButtons.OKCancel,
                System.Windows.Forms.MessageBoxIcon.Question);

            if (confirm != System.Windows.Forms.DialogResult.OK)
                return "Geannuleerd.";

            return ExecuteFormat(codeModule);
        }

        /// <summary>
        /// Formatteert alle modules in het actieve VBA-project.
        /// Toont eerst een bevestigingsdialog met projectnaam en scope.
        /// </summary>
        public string FormatFile(VBE vbe)
        {
            if (vbe == null || vbe.ActiveVBProject == null)
                return "Geen actief VBA project.";

            var project = vbe.ActiveVBProject;
            var formattable = new List<VBComponent>();
            foreach (VBComponent comp in project.VBComponents)
                if (comp.CodeModule.CountOfLines > 0)
                    formattable.Add(comp);

            if (formattable.Count == 0)
                return "Geen modules gevonden in project.";

            var confirm = System.Windows.Forms.MessageBox.Show(
                string.Format("Alle {0} modules in project '{1}' formatteren?\n\nScope: volledig VBA-bestand",
                    formattable.Count, project.Name),
                "Bevestig: Volledig VBA-bestand",
                System.Windows.Forms.MessageBoxButtons.OKCancel,
                System.Windows.Forms.MessageBoxIcon.Question);

            if (confirm != System.Windows.Forms.DialogResult.OK)
                return "Geannuleerd.";

            var results = new System.Text.StringBuilder();
            foreach (var comp in formattable)
            {
                try
                {
                    string r = ExecuteFormat(comp.CodeModule);
                    results.AppendLine(comp.Name + ": " + r);
                }
                catch (Exception ex)
                {
                    results.AppendLine(comp.Name + ": FOUT - " + ex.Message);
                }
            }

            return string.Format("{0} modules geformatteerd in project '{1}'.\n\n{2}",
                formattable.Count, project.Name, results.ToString().TrimEnd());
        }

        private string ExecuteFormat(CodeModule codeModule)
        {
            if (codeModule == null)
                return "Geen code module geselecteerd.";

            int totalLines = codeModule.CountOfLines;
            
            // Stap 1: Lees alle regels
            List<CodeLine> lines = new List<CodeLine>();
            for (int i = 1; i <= totalLines; i++)
            {
                lines.Add(new CodeLine
                {
                    LineNumber = i,
                    OriginalText = codeModule.Lines[i, 1]
                });
            }

            // Stap 2: Sorteer Dim/Const binnen procedures
            SortDeclarationsInProcedures(lines);

            // Stap 3: Verwijder onnodige lege regels
            RemoveExcessiveBlankLines(lines);

            // Stap 4: Voeg ontbrekende lege regels toe
            AddMissingBlankLines(lines);

            // Stap 5: Formatteer indentatie
            FormatIndentation(lines);

            // Stap 6: Zorg dat alle regels een NewText hebben
            foreach (var line in lines)
            {
                if (line.NewText == null && !line.MarkedForDeletion && line.OriginalText != null)
                    line.NewText = line.OriginalText;
            }

            // Stap 7: Verwijder regels die gemarkeerd zijn voor verwijdering
            lines.RemoveAll(line => line.MarkedForDeletion);

            // Stap 8 (optioneel): Keyword casing
            if (FormatterSettings.KeywordsCase != "preserve")
                ApplyKeywordCasing(lines);

            // Stap 9 (optioneel): Spaties
            if (FormatterSettings.SpacingAroundOperators || FormatterSettings.SpacingAfterComma || FormatterSettings.SpacingInsideParentheses)
                ApplySpacing(lines);

            // Stap 10 (optioneel): Commentaar stijl
            if (FormatterSettings.CommentStyle != "preserve")
                ApplyCommentStyle(lines);

            // Stap 11 (optioneel): Option Explicit
            if (FormatterSettings.DeclarationsOptionExplicit != "preserve")
                ApplyOptionExplicit(lines);

            // Stap 12 (optioneel): Trailing whitespace
            if (FormatterSettings.MiscRemoveTrailingWhitespace)
                foreach (var line in lines)
                    if (line.NewText != null) line.NewText = line.NewText.TrimEnd();

            // Stap 13 (optioneel): Final newline
            if (FormatterSettings.MiscEnsureFinalNewline && lines.Count > 0 && !string.IsNullOrEmpty(lines[lines.Count - 1].NewText))
                lines.Add(new CodeLine { OriginalText = "", NewText = "" });

            // Stap 8: Schrijf terug naar code module
            int changes = 0;
            int lineIndex = 0;
            
            // Eerst alle oude regels verwijderen van achter naar voor
            for (int i = totalLines; i >= 1; i--)
            {
                codeModule.DeleteLines(i, 1);
            }

            // Dan nieuwe regels invoegen
            foreach (var line in lines)
            {
                codeModule.InsertLines(lineIndex + 1, line.NewText);
                lineIndex++;
                if (line.NewText != line.OriginalText)
                    changes++;
            }

            return string.Format("Code formatting voltooid!\n\n{0} regels aangepast\n{1} totale regels",
                changes, lines.Count);
        }

        private void SortDeclarationsInProcedures(List<CodeLine> lines)
        {
            bool inProcedure = false;
            int procedureStart = -1;
            List<DimStatement> dims = new List<DimStatement>();
            List<CodeLine> consts = new List<CodeLine>();
            List<int> declarationIndices = new List<int>();

            for (int i = 0; i < lines.Count; i++)
            {
                string trimmed = lines[i].OriginalText.Trim();
                string upper = trimmed.ToUpper();

                // Procedure start
                if (Regex.IsMatch(upper, @"^(PUBLIC|PRIVATE|FRIEND)?\s*(SUB|FUNCTION|PROPERTY)\s+"))
                {
                    inProcedure = true;
                    procedureStart = i;
                    dims.Clear();
                    consts.Clear();
                    declarationIndices.Clear();
                    continue;
                }

                // Procedure end - sorteer en lijn uit gevonden declaraties
                if (inProcedure && Regex.IsMatch(upper, @"^END\s+(SUB|FUNCTION|PROPERTY)"))
                {
                    if (dims.Count > 0 || consts.Count > 0)
                    {
                        // Sorteer Dims op type volgens configuratie
                        if (FormatterSettings.SortDimsByType)
                        {
                            dims.Sort((a, b) => CompareDimTypes(a.Type, b.Type));
                        }

                        // Bepaal uitlijning positie voor "As Type"
                        int maxVarNameLength = 0;
                        if (FormatterSettings.AlignAsTypes)
                        {
                            foreach (var dim in dims)
                            {
                                if (dim.VariableName.Length > maxVarNameLength)
                                    maxVarNameLength = dim.VariableName.Length;
                            }
                        }

                        // EERST: Markeer oude declaraties voor verwijdering (op originele indices)
                        foreach (int idx in declarationIndices)
                        {
                            lines[idx].MarkedForDeletion = true;
                        }

                        // DAARNA: Voeg nieuwe declaraties toe direct na procedure header
                        int insertPos = procedureStart + 1;
                        
                        // Voeg eerst alle Dims toe
                        foreach (var dim in dims)
                        {
                            string formattedDim = FormatDimStatement(dim, maxVarNameLength);
                            lines.Insert(insertPos, new CodeLine
                            {
                                OriginalText = formattedDim,
                                NewText = null  // Laat FormatIndentation() de indentatie doen
                            });
                            insertPos++;
                        }
                        
                        // Dan alle Consts
                        foreach (var constLine in consts)
                        {
                            lines.Insert(insertPos, new CodeLine
                            {
                                OriginalText = constLine.OriginalText.Trim(),
                                NewText = null  // Laat FormatIndentation() de indentatie doen
                            });
                            insertPos++;
                        }
                    }

                    inProcedure = false;
                    dims.Clear();
                    consts.Clear();
                    declarationIndices.Clear();
                    continue;
                }

                // Zoek Dim/Const binnen procedure
                if (inProcedure && i > procedureStart)
                {
                    if (IsDimStatement(trimmed))
                    {
                        DimStatement dimStmt = ParseDimStatement(trimmed);
                        if (dimStmt != null)
                        {
                            dims.Add(dimStmt);
                            declarationIndices.Add(i);
                        }
                    }
                    else if (IsConstStatement(trimmed))
                    {
                        consts.Add(lines[i]);
                        declarationIndices.Add(i);
                    }
                }
            }
        }

        private DimStatement ParseDimStatement(string line)
        {
            // Parse "Dim variableName As Type"
            Match match = Regex.Match(line, @"^\s*Dim\s+(\w+)\s+As\s+(.+)$", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return new DimStatement
                {
                    VariableName = match.Groups[1].Value,
                    Type = match.Groups[2].Value.Trim().ToUpper(),
                    OriginalLine = line
                };
            }
            return null;
        }

        private string FormatDimStatement(DimStatement dim, int alignPosition)
        {
            if (!FormatterSettings.AlignAsTypes || alignPosition == 0)
            {
                return string.Format("Dim {0} As {1}", dim.VariableName, dim.Type);
            }

            // Bereken aantal spaties nodig voor uitlijning
            int spacesNeeded = alignPosition - dim.VariableName.Length + FormatterSettings.MinimumSpaceBeforeAsType;
            if (spacesNeeded < FormatterSettings.MinimumSpaceBeforeAsType)
                spacesNeeded = FormatterSettings.MinimumSpaceBeforeAsType;

            return string.Format("Dim {0}{1}As {2}", 
                dim.VariableName, 
                new string(' ', spacesNeeded), 
                dim.Type);
        }

        private int CompareDimTypes(string typeA, string typeB)
        {
            int indexA = FormatterSettings.DimTypeSortOrder.IndexOf(typeA);
            int indexB = FormatterSettings.DimTypeSortOrder.IndexOf(typeB);

            // Als beide in de lijst staan, sorteer op positie
            if (indexA >= 0 && indexB >= 0)
                return indexA.CompareTo(indexB);

            // Als alleen A in lijst staat, A komt eerst
            if (indexA >= 0)
                return -1;

            // Als alleen B in lijst staat, B komt eerst
            if (indexB >= 0)
                return 1;

            // Beide niet in lijst, sorteer alfabetisch
            return typeA.CompareTo(typeB);
        }

        private void RemoveExcessiveBlankLines(List<CodeLine> lines)
        {
            // Verwijder meer dan 2 lege regels na elkaar
            for (int i = lines.Count - 1; i >= 0; i--)
            {
                if (string.IsNullOrWhiteSpace(lines[i].OriginalText))
                {
                    int blankCount = 1;
                    int j = i - 1;
                    while (j >= 0 && string.IsNullOrWhiteSpace(lines[j].OriginalText))
                    {
                        blankCount++;
                        j--;
                    }

                    // Als meer dan max lege regels, verwijder overtollige
                    int maxBlank = Math.Max(0, FormatterSettings.MiscKeepBlankLinesMax);
                    if (blankCount > maxBlank)
                    {
                        int toRemove = blankCount - maxBlank;
                        for (int k = 0; k < toRemove && i < lines.Count; k++)
                        {
                            lines[i].MarkedForDeletion = true;
                            i--;
                        }
                    }
                }
            }

            // Verwijder lege regels aan begin en einde van procedures
            bool inProcedure = false;
            int procedureStart = -1;

            for (int i = 0; i < lines.Count; i++)
            {
                string upper = lines[i].OriginalText.Trim().ToUpper();

                if (Regex.IsMatch(upper, @"^(PUBLIC|PRIVATE|FRIEND)?\s*(SUB|FUNCTION|PROPERTY)\s+"))
                {
                    inProcedure = true;
                    procedureStart = i;

                    // Verwijder lege regels direct na procedure header
                    int j = i + 1;
                    while (j < lines.Count && string.IsNullOrWhiteSpace(lines[j].OriginalText))
                    {
                        lines[j].MarkedForDeletion = true;
                        j++;
                    }
                }

                if (inProcedure && Regex.IsMatch(upper, @"^END\s+(SUB|FUNCTION|PROPERTY)"))
                {
                    // Verwijder lege regels direct voor End Sub/Function
                    int j = i - 1;
                    while (j > procedureStart && string.IsNullOrWhiteSpace(lines[j].OriginalText))
                    {
                        lines[j].MarkedForDeletion = true;
                        j--;
                    }
                    inProcedure = false;
                }
            }
        }

        private void AddMissingBlankLines(List<CodeLine> lines)
        {
            int wantedBlanks = FormatterSettings.BlockBlankLinesBetweenProcedures;

            for (int i = lines.Count - 1; i >= 1; i--)
            {
                string current = lines[i].OriginalText.Trim().ToUpper();
                string previous = lines[i - 1].OriginalText.Trim().ToUpper();

                // Na End Sub/Function/Property: voeg gewenst aantal lege regels in
                if (Regex.IsMatch(previous, @"^END\s+(SUB|FUNCTION|PROPERTY)") &&
                    !string.IsNullOrWhiteSpace(current) &&
                    i < lines.Count - 1)
                {
                    for (int b = 0; b < wantedBlanks; b++)
                        lines.Insert(i, new CodeLine { OriginalText = "", NewText = "" });
                }

                // Voor nieuwe procedure: voeg gewenst aantal lege regels in
                if (Regex.IsMatch(current, @"^(PUBLIC|PRIVATE|FRIEND)?\s*(SUB|FUNCTION|PROPERTY)\s+") &&
                    !string.IsNullOrWhiteSpace(previous) &&
                    !Regex.IsMatch(previous, @"^END\s+(SUB|FUNCTION|PROPERTY)") &&
                    i > 0)
                {
                    for (int b = 0; b < wantedBlanks; b++)
                        lines.Insert(i, new CodeLine { OriginalText = "", NewText = "" });
                }
            }

            // Voeg lege regels toe na declaratie-blok binnen procedures
            int wantedDeclBlanks = FormatterSettings.BlockBlankLinesAfterDeclarations;
            bool inProc = false;
            bool inDeclBlock = false;
            int lastDeclIdx = -1;

            for (int i = 0; i < lines.Count; i++)
            {
                if (lines[i].MarkedForDeletion) continue;

                string upper = lines[i].OriginalText.Trim().ToUpper();

                if (Regex.IsMatch(upper, @"^(PUBLIC|PRIVATE|FRIEND)?\s*(SUB|FUNCTION|PROPERTY)\s+"))
                {
                    inProc = true;
                    inDeclBlock = false;
                    lastDeclIdx = -1;
                    continue;
                }

                if (!inProc) continue;

                if (Regex.IsMatch(upper, @"^END\s+(SUB|FUNCTION|PROPERTY)"))
                {
                    inProc = false;
                    inDeclBlock = false;
                    lastDeclIdx = -1;
                    continue;
                }

                string trimmed = lines[i].OriginalText.Trim();
                bool isDecl = IsDimStatement(trimmed) || IsConstStatement(trimmed);
                bool isBlank = string.IsNullOrWhiteSpace(trimmed);

                if (isDecl)
                {
                    inDeclBlock = true;
                    lastDeclIdx = i;
                }
                else if (inDeclBlock && !isBlank)
                {
                    // Eerste niet-lege, niet-declaratie regel na declaratie-blok gevonden.
                    // Verwijder bestaande lege regels tussen einde declaraties en deze regel.
                    for (int j = lastDeclIdx + 1; j < i; j++)
                    {
                        if (!lines[j].MarkedForDeletion && string.IsNullOrWhiteSpace(lines[j].OriginalText))
                            lines[j].MarkedForDeletion = true;
                    }
                    // Voeg gewenst aantal lege regels in vóór deze regel.
                    for (int b = 0; b < wantedDeclBlanks; b++)
                        lines.Insert(i, new CodeLine { OriginalText = "", NewText = "" });
                    i += wantedDeclBlanks;
                    inDeclBlock = false;
                    lastDeclIdx = -1;
                }
            }
        }

        private void FormatIndentation(List<CodeLine> lines)
        {
            int indentLevel = 0;
            bool inProcedure = false;

            for (int i = 0; i < lines.Count; i++)
            {
                if (lines[i].MarkedForDeletion)
                    continue; // Skip regels die verwijderd worden

                string originalLine = lines[i].OriginalText;
                string trimmedLine = originalLine.Trim();

                // Lege regels
                if (string.IsNullOrWhiteSpace(trimmedLine))
                {
                    lines[i].NewText = "";
                    continue;
                }

                // Commentaar regels - gebruik huidige indent level
                if (trimmedLine.StartsWith("'"))
                {
                    lines[i].NewText = string.Concat(Enumerable.Repeat(IndentUnit, indentLevel)) + trimmedLine;
                    continue;
                }

                // Verwijder inline comments voor analyse
                string lineForAnalysis = RemoveInlineComment(trimmedLine);

                // Bepaal nieuwe indent level voor deze regel
                int newIndentLevel = CalculateIndentLevel(lineForAnalysis, ref indentLevel, ref inProcedure);

                // Maak nieuwe regel met correcte indentatie
                lines[i].NewText = string.Concat(Enumerable.Repeat(IndentUnit, newIndentLevel)) + trimmedLine;
            }
        }

        private int CalculateIndentLevel(string line, ref int currentLevel, ref bool inProcedure)
        {
            string lineUpper = line.ToUpper().Trim();
            int levelForThisLine = currentLevel;

            // #If preprocessor directives - geen indentatie verandering
            if (Regex.IsMatch(lineUpper, @"^#IF\s+") || 
                Regex.IsMatch(lineUpper, @"^#ELSEIF\s+") ||
                lineUpper == "#ELSE" ||
                lineUpper == "#END IF")
            {
                return 0; // Preprocessor altijd op niveau 0
            }

            // Procedure start (Sub, Function, Property)
            if (Regex.IsMatch(lineUpper, @"^(PUBLIC|PRIVATE|FRIEND)?\s*(SUB|FUNCTION|PROPERTY)\s+"))
            {
                levelForThisLine = 0;
                currentLevel = 1;
                inProcedure = true;
                return levelForThisLine;
            }

            // Procedure end
            if (Regex.IsMatch(lineUpper, @"^END\s+(SUB|FUNCTION|PROPERTY)"))
            {
                currentLevel = 0;
                levelForThisLine = 0;
                inProcedure = false;
                return levelForThisLine;
            }

            // Class/Type definitions
            if (Regex.IsMatch(lineUpper, @"^(PUBLIC|PRIVATE)?\s*TYPE\s+") ||
                Regex.IsMatch(lineUpper, @"^(PUBLIC|PRIVATE)?\s*ENUM\s+"))
            {
                levelForThisLine = 0;
                currentLevel = 1;
                return levelForThisLine;
            }

            if (Regex.IsMatch(lineUpper, @"^END\s+(TYPE|ENUM)"))
            {
                currentLevel = 0;
                levelForThisLine = 0;
                return levelForThisLine;
            }

            // If statements (eenregelig vs. meerregelig)
            if (Regex.IsMatch(lineUpper, @"^IF\s+.+\s+THEN\s*$"))
            {
                // Meerregelig If zonder code op dezelfde regel
                levelForThisLine = currentLevel;
                currentLevel++;
                return levelForThisLine;
            }
            else if (Regex.IsMatch(lineUpper, @"^IF\s+.+\s+THEN\s+.+"))
            {
                // Eenregelig If met code achter THEN
                return currentLevel;
            }

            // ElseIf, Else
            if (Regex.IsMatch(lineUpper, @"^(ELSEIF|ELSE IF)\s+"))
            {
                levelForThisLine = currentLevel - 1;
                return levelForThisLine;
            }

            if (lineUpper == "ELSE")
            {
                levelForThisLine = currentLevel - 1;
                return levelForThisLine;
            }

            // End If
            if (lineUpper == "END IF")
            {
                currentLevel--;
                levelForThisLine = currentLevel;
                return levelForThisLine;
            }

            // Select Case
            if (Regex.IsMatch(lineUpper, @"^SELECT\s+CASE\s+"))
            {
                levelForThisLine = currentLevel;
                currentLevel++;
                return levelForThisLine;
            }

            // Case statements
            if (Regex.IsMatch(lineUpper, @"^CASE\s+"))
            {
                levelForThisLine = currentLevel - 1;
                return levelForThisLine;
            }

            // End Select
            if (lineUpper == "END SELECT")
            {
                currentLevel--;
                levelForThisLine = currentLevel;
                return levelForThisLine;
            }

            // For loops
            if (Regex.IsMatch(lineUpper, @"^FOR\s+(EACH\s+)?"))
            {
                levelForThisLine = currentLevel;
                currentLevel++;
                return levelForThisLine;
            }

            // Next
            if (Regex.IsMatch(lineUpper, @"^NEXT(\s+|$)"))
            {
                currentLevel--;
                levelForThisLine = currentLevel;
                return levelForThisLine;
            }

            // Do loops
            if (Regex.IsMatch(lineUpper, @"^DO(\s+(WHILE|UNTIL)|$)"))
            {
                levelForThisLine = currentLevel;
                currentLevel++;
                return levelForThisLine;
            }

            // Loop
            if (Regex.IsMatch(lineUpper, @"^LOOP(\s+(WHILE|UNTIL))?"))
            {
                currentLevel--;
                levelForThisLine = currentLevel;
                return levelForThisLine;
            }

            // While loops
            if (Regex.IsMatch(lineUpper, @"^WHILE\s+"))
            {
                levelForThisLine = currentLevel;
                currentLevel++;
                return levelForThisLine;
            }

            // Wend
            if (lineUpper == "WEND")
            {
                currentLevel--;
                levelForThisLine = currentLevel;
                return levelForThisLine;
            }

            // With statements
            if (Regex.IsMatch(lineUpper, @"^WITH\s+"))
            {
                levelForThisLine = currentLevel;
                currentLevel++;
                return levelForThisLine;
            }

            // End With
            if (lineUpper == "END WITH")
            {
                currentLevel--;
                levelForThisLine = currentLevel;
                return levelForThisLine;
            }

            // Line continuation (_) - behoud huidige indent
            if (line.TrimEnd().EndsWith("_"))
            {
                return currentLevel;
            }

            // Default: gebruik huidige indent level
            return currentLevel;
        }

        private string RemoveInlineComment(string line)
        {
            // Verwijder inline comments, maar pas op voor strings met '
            bool inString = false;
            for (int i = 0; i < line.Length; i++)
            {
                if (line[i] == '"')
                    inString = !inString;

                if (!inString && line[i] == '\'')
                    return line.Substring(0, i).TrimEnd();
            }
            return line;
        }

        private bool IsDimStatement(string line)
        {
            string upper = line.ToUpper().Trim();
            return Regex.IsMatch(upper, @"^DIM\s+", RegexOptions.IgnoreCase) &&
                   !upper.StartsWith("'");
        }

        private bool IsConstStatement(string line)
        {
            string upper = line.ToUpper().Trim();
            return Regex.IsMatch(upper, @"^(PUBLIC|PRIVATE)?\s*CONST\s+", RegexOptions.IgnoreCase) &&
                   !upper.StartsWith("'");
        }

        // ── New formatter steps ──────────────────────────────────────────────

        private static readonly HashSet<string> VbaKeywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "And", "As", "Boolean", "ByRef", "ByVal", "Call", "Case", "Const",
            "Date", "Debug", "Dim", "Do", "Double", "Each", "Else", "ElseIf",
            "Empty", "End", "Enum", "Error", "Exit", "False", "For", "Friend",
            "Function", "Get", "GoTo", "If", "In", "Integer", "Is", "Let",
            "Like", "Long", "Loop", "Mod", "New", "Next", "Not", "Nothing",
            "Null", "Object", "On", "Option", "Or", "Private", "Property",
            "Public", "Resume", "Select", "Set", "Static", "Step", "Stop",
            "String", "Sub", "Then", "To", "True", "Type", "Until", "Variant",
            "Wend", "While", "With", "Xor", "WithEvents", "Implements",
            "Explicit", "Compare", "Base", "Optional", "ParamArray", "Declare",
            "Lib", "Alias", "AddressOf", "Me", "ByRef", "ByVal"
        };

        private void ApplyKeywordCasing(List<CodeLine> lines)
        {
            string mode = FormatterSettings.KeywordsCase; // "uppercase"|"lowercase"|"pascal"
            foreach (var cl in lines)
            {
                if (cl.MarkedForDeletion) continue;
                cl.NewText = ApplyKeywordCasingToLine(cl.NewText, mode);
            }
        }

        private string ApplyKeywordCasingToLine(string line, string mode)
        {
            // Split into tokens preserving delimiters; skip content inside string literals
            var result = new System.Text.StringBuilder();
            bool inString = false;
            int i = 0;
            while (i < line.Length)
            {
                char c = line[i];
                if (c == '"')
                {
                    inString = !inString;
                    result.Append(c);
                    i++;
                    continue;
                }
                if (inString)
                {
                    result.Append(c);
                    i++;
                    continue;
                }
                // Comment start — append rest verbatim
                if (c == '\'')
                {
                    result.Append(line.Substring(i));
                    break;
                }
                // Word character?
                if (char.IsLetter(c) || c == '_')
                {
                    int start = i;
                    while (i < line.Length && (char.IsLetterOrDigit(line[i]) || line[i] == '_'))
                        i++;
                    string word = line.Substring(start, i - start);
                    if (VbaKeywords.Contains(word))
                    {
                        switch (mode)
                        {
                            case "uppercase": word = word.ToUpper(); break;
                            case "lowercase": word = word.ToLower(); break;
                            case "pascal":
                                // Find the canonical pascal form from the set
                                foreach (var kw in VbaKeywords)
                                    if (string.Equals(kw, word, StringComparison.OrdinalIgnoreCase))
                                    { word = kw; break; }
                                break;
                        }
                    }
                    result.Append(word);
                }
                else
                {
                    result.Append(c);
                    i++;
                }
            }
            return result.ToString();
        }

        private void ApplySpacing(List<CodeLine> lines)
        {
            bool ops = FormatterSettings.SpacingAroundOperators;
            bool comma = FormatterSettings.SpacingAfterComma;
            bool parens = FormatterSettings.SpacingInsideParentheses;

            foreach (var cl in lines)
            {
                if (cl.MarkedForDeletion) continue;
                string t = cl.NewText;
                string code = RemoveInlineComment(t);
                string comment = t.Length > code.Length ? t.Substring(code.Length) : "";

                if (ops)
                {
                    // Ensure spaces around = + - & * / < > but not inside strings
                    code = ApplyOperatorSpacing(code);
                }
                if (comma)
                {
                    // Ensure space after comma (outside strings)
                    code = Regex.Replace(code, @",(?!\s)", ", ");
                }
                if (parens)
                {
                    // Space after ( and before ) outside strings
                    code = Regex.Replace(code, @"\((?!\s)", "( ");
                    code = Regex.Replace(code, @"(?<!\s)\)", " )");
                }
                cl.NewText = code + comment;
            }
        }

        private string ApplyOperatorSpacing(string code)
        {
            // Use regex to add spaces around binary operators: = + - & * / < > <>  <= >=
            // Skip inside string literals using a character-by-character pass for safety
            var sb = new System.Text.StringBuilder();
            bool inStr = false;
            int i = 0;
            while (i < code.Length)
            {
                char c = code[i];
                if (c == '"') { inStr = !inStr; sb.Append(c); i++; continue; }
                if (inStr) { sb.Append(c); i++; continue; }

                // Check for two-char operators first
                if (i + 1 < code.Length)
                {
                    string two = code.Substring(i, 2);
                    if (two == "<>" || two == "<=" || two == ">=" || two == ":=")
                    {
                        // Ensure space before
                        if (sb.Length > 0 && sb[sb.Length - 1] != ' ')
                            sb.Append(' ');
                        sb.Append(two);
                        i += 2;
                        // Ensure space after
                        if (i < code.Length && code[i] != ' ')
                            sb.Append(' ');
                        continue;
                    }
                }

                if (c == '=' || c == '+' || c == '-' || c == '&' || c == '*' || c == '/' || c == '<' || c == '>')
                {
                    // Don't double-space if neighbour is already space
                    if (sb.Length > 0 && sb[sb.Length - 1] != ' ')
                        sb.Append(' ');
                    sb.Append(c);
                    i++;
                    if (i < code.Length && code[i] != ' ')
                        sb.Append(' ');
                }
                else
                {
                    sb.Append(c);
                    i++;
                }
            }
            return sb.ToString();
        }

        private void ApplyCommentStyle(List<CodeLine> lines)
        {
            string style = FormatterSettings.CommentStyle;
            foreach (var cl in lines)
            {
                if (cl.MarkedForDeletion) continue;
                string trimmed = cl.NewText.TrimStart();
                string indent = cl.NewText.Substring(0, cl.NewText.Length - trimmed.Length);

                if (style == "apostrophe" && Regex.IsMatch(trimmed, @"^Rem\s", RegexOptions.IgnoreCase))
                {
                    string rest = trimmed.Substring(trimmed.IndexOf(' ') + 1);
                    cl.NewText = indent + "' " + rest;
                }
                else if (style == "rem" && trimmed.StartsWith("'"))
                {
                    string rest = trimmed.TrimStart('\'').TrimStart();
                    cl.NewText = indent + "Rem " + rest;
                }
            }
        }

        private void ApplyOptionExplicit(List<CodeLine> lines)
        {
            string setting = FormatterSettings.DeclarationsOptionExplicit;
            if (setting == "preserve") return;

            // Find existing Option Explicit line
            int existingIndex = -1;
            for (int i = 0; i < lines.Count; i++)
            {
                if (!lines[i].MarkedForDeletion &&
                    Regex.IsMatch(lines[i].NewText.Trim(), @"^Option\s+Explicit\s*$", RegexOptions.IgnoreCase))
                {
                    existingIndex = i;
                    break;
                }
            }

            if (setting == "remove" && existingIndex >= 0)
            {
                lines[existingIndex].MarkedForDeletion = true;
            }
            else if (setting == "require" && existingIndex < 0)
            {
                // Insert at the top (after any initial blank lines or Option Compare etc.)
                int insertAt = 0;
                for (int i = 0; i < lines.Count; i++)
                {
                    string t = lines[i].NewText.Trim();
                    if (t == "" || Regex.IsMatch(t, @"^Option\s+", RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(t, @"^'", RegexOptions.IgnoreCase))
                        insertAt = i + 1;
                    else
                        break;
                }
                // Build a synthetic CodeLine
                var newLine = new CodeLine
                {
                    LineNumber = 0,
                    OriginalText = "Option Explicit",
                    NewText = "Option Explicit",
                    MarkedForDeletion = false
                };
                lines.Insert(insertAt, newLine);
            }
        }

        // ─────────────────────────────────────────────────────────────────────

        private class CodeLine
        {
            public int LineNumber { get; set; }
            public string OriginalText { get; set; }
            public string NewText { get; set; }
            public bool MarkedForDeletion { get; set; }
        }

        private class DimStatement
        {
            public string VariableName { get; set; }
            public string Type { get; set; }
            public string OriginalLine { get; set; }
        }
    }
}
