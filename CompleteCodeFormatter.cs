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
        private const string TAB = "    "; // 4 spaties

        public string FormatCode(CodeModule codeModule)
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

            // Stap 6: Zorg dat alle regels een NewText hebben (tenzij gemarkeerd voor verwijdering)
            foreach (var line in lines)
            {
                if (line.NewText == null && !line.MarkedForDeletion && line.OriginalText != null)
                {
                    // Regel is niet gemarkeerd voor verwijdering en heeft geen nieuwe tekst
                    // Gebruik originele tekst
                    line.NewText = line.OriginalText;
                }
            }

            // Stap 7: Verwijder regels die expliciet gemarkeerd zijn voor verwijdering
            lines.RemoveAll(line => line.MarkedForDeletion);

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

                    // Als meer dan 2 lege regels, verwijder overtollige
                    if (blankCount > 2)
                    {
                        int toRemove = blankCount - 2;
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
            // Voeg lege regel toe tussen verschillende procedures
            for (int i = lines.Count - 1; i >= 1; i--)
            {
                string current = lines[i].OriginalText.Trim().ToUpper();
                string previous = lines[i - 1].OriginalText.Trim().ToUpper();

                // Na End Sub/Function/Property moet een lege regel komen (tenzij einde module)
                if (Regex.IsMatch(previous, @"^END\s+(SUB|FUNCTION|PROPERTY)") &&
                    !string.IsNullOrWhiteSpace(current) &&
                    i < lines.Count - 1)
                {
                    lines.Insert(i, new CodeLine
                    {
                        OriginalText = "",
                        NewText = ""
                    });
                }

                // Voor nieuwe procedure moet een lege regel komen (tenzij begin module)
                if (Regex.IsMatch(current, @"^(PUBLIC|PRIVATE|FRIEND)?\s*(SUB|FUNCTION|PROPERTY)\s+") &&
                    !string.IsNullOrWhiteSpace(previous) &&
                    !Regex.IsMatch(previous, @"^END\s+(SUB|FUNCTION|PROPERTY)") &&
                    i > 0)
                {
                    lines.Insert(i, new CodeLine
                    {
                        OriginalText = "",
                        NewText = ""
                    });
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
                    lines[i].NewText = new string(' ', indentLevel * TAB.Length) + trimmedLine;
                    continue;
                }

                // Verwijder inline comments voor analyse
                string lineForAnalysis = RemoveInlineComment(trimmedLine);

                // Bepaal nieuwe indent level voor deze regel
                int newIndentLevel = CalculateIndentLevel(lineForAnalysis, ref indentLevel, ref inProcedure);

                // Maak nieuwe regel met correcte indentatie
                lines[i].NewText = new string(' ', newIndentLevel * TAB.Length) + trimmedLine;
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
