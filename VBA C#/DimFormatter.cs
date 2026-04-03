using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;

namespace VBEAddIn
{
    /// <summary>
    /// Handles formatting of Dim statements in VBA code
    /// </summary>
    [ComVisible(true)]
    [Guid("C1D2E3F4-A5B6-4C78-D901-E234F5678A90")]
    [ProgId("VBEAddIn.DimFormatter")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class DimFormatter
    {
        private const string TAB = "    "; // 4 spaces as one tab

        /// <summary>
        /// Public parameterless constructor for COM
        /// </summary>
        public DimFormatter()
        {
        }

        /// <summary>
        /// Formats Dim statements in the current procedure to be sorted and aligned
        /// </summary>
        /// <param name="codeModule">The VBA code module to format</param>
        /// <returns>A message describing the formatting results</returns>
        public string FormatDimStatements(CodeModule codeModule)
        {
            if (codeModule == null)
            {
                throw new ArgumentNullException("codeModule");
            }

            // Get cursor position to determine current procedure
            int cursorLine;
            int totalLines = codeModule.CountOfLines;
            
            try
            {
                // Get cursor position from active pane selection
                CodePane activePane = codeModule.CodePane;
                if (activePane != null)
                {
                    int startLine, startColumn, endLine, endColumn;
                    activePane.GetSelection(out startLine, out startColumn, out endLine, out endColumn);
                    cursorLine = startLine;
                }
                else
                {
                    return "Geen actieve code pane gevonden.";
                }
            }
            catch
            {
                // Als we selectie niet kunnen krijgen, gebruik eerste regel
                cursorLine = 1;
            }

            // Find current procedure boundaries
            int procedureStart = -1;
            int procedureEnd = -1;

            // Search backwards for procedure start
            for (int i = cursorLine; i >= 1; i--)
            {
                string line = codeModule.Lines[i, 1].Trim().ToUpper();
                if (Regex.IsMatch(line, @"^(PUBLIC|PRIVATE|FRIEND)?\s*(SUB|FUNCTION|PROPERTY)\s+"))
                {
                    procedureStart = i;
                    break;
                }
            }

            if (procedureStart == -1)
            {
                return "Geen procedure gevonden. Zet cursor in een Sub of Function.";
            }

            // Search forwards for procedure end
            for (int i = procedureStart + 1; i <= totalLines; i++)
            {
                string line = codeModule.Lines[i, 1].Trim().ToUpper();
                if (Regex.IsMatch(line, @"^END\s+(SUB|FUNCTION|PROPERTY)"))
                {
                    procedureEnd = i;
                    break;
                }
            }

            if (procedureEnd == -1)
            {
                return "Geen procedure einde gevonden.";
            }

            // Read all lines in procedure
            List<CodeLine> lines = new List<CodeLine>();
            for (int i = procedureStart; i <= procedureEnd; i++)
            {
                lines.Add(new CodeLine
                {
                    LineNumber = i,
                    OriginalText = codeModule.Lines[i, 1]
                });
            }

            // Find and sort Dims
            List<DimStatement> dims = new List<DimStatement>();
            List<CodeLine> consts = new List<CodeLine>();
            List<int> declarationIndices = new List<int>();

            for (int i = 1; i < lines.Count - 1; i++) // Skip procedure header and end
            {
                string trimmed = lines[i].OriginalText.Trim();
                
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

            if (dims.Count == 0 && consts.Count == 0)
            {
                return "Geen Dim of Const statements gevonden in deze procedure.";
            }

            // Sort Dims
            if (FormatterSettings.SortDimsByType)
            {
                dims.Sort((a, b) => CompareDimTypes(a.Type, b.Type));
            }

            // Calculate alignment
            int maxVarNameLength = 0;
            if (FormatterSettings.AlignAsTypes)
            {
                foreach (var dim in dims)
                {
                    if (dim.VariableName.Length > maxVarNameLength)
                        maxVarNameLength = dim.VariableName.Length;
                }
            }

            // Mark old declarations for deletion
            foreach (int idx in declarationIndices)
            {
                lines[idx].MarkedForDeletion = true;
            }

            // Insert new declarations after procedure header
            int insertPos = 1; // After procedure header (index 0)
            foreach (var dim in dims)
            {
                string formattedDim = FormatDimStatement(dim, maxVarNameLength);
                lines.Insert(insertPos, new CodeLine
                {
                    OriginalText = TAB + formattedDim,
                    NewText = TAB + formattedDim
                });
                insertPos++;
            }

            foreach (var constLine in consts)
            {
                lines.Insert(insertPos, new CodeLine
                {
                    OriginalText = TAB + constLine.OriginalText.Trim(),
                    NewText = TAB + constLine.OriginalText.Trim()
                });
                insertPos++;
            }

            // Remove marked lines
            lines.RemoveAll(line => line.MarkedForDeletion);

            // Write back to code module
            for (int i = procedureEnd; i >= procedureStart; i--)
            {
                codeModule.DeleteLines(i, 1);
            }

            int lineIndex = procedureStart;
            foreach (var line in lines)
            {
                codeModule.InsertLines(lineIndex, line.NewText ?? line.OriginalText);
                lineIndex++;
            }

            return string.Format("Procedure geformatteerd!\n\n{0} Dim(s) gesorteerd en uitgelijnd\n{1} Const(s) toegevoegd",
                dims.Count, consts.Count);
        }

        private DimStatement ParseDimStatement(string line)
        {
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

            int spacesNeeded = alignPosition - dim.VariableName.Length + FormatterSettings.MinimumSpaceBeforeAsType;
            if (spacesNeeded < FormatterSettings.MinimumSpaceBeforeAsType)
                spacesNeeded = FormatterSettings.MinimumSpaceBeforeAsType;

            return string.Format("Dim {0}{1}As {2}", dim.VariableName, new string(' ', spacesNeeded), dim.Type);
        }

        private int CompareDimTypes(string typeA, string typeB)
        {
            int indexA = FormatterSettings.DimTypeSortOrder.IndexOf(typeA);
            int indexB = FormatterSettings.DimTypeSortOrder.IndexOf(typeB);

            if (indexA >= 0 && indexB >= 0)
                return indexA.CompareTo(indexB);

            if (indexA >= 0)
                return -1;

            if (indexB >= 0)
                return 1;

            return typeA.CompareTo(typeB);
        }

        private bool IsDimStatement(string line)
        {
            if (string.IsNullOrWhiteSpace(line))
                return false;

            if (line.TrimStart().StartsWith("'") || 
                line.TrimStart().StartsWith("Rem ", StringComparison.OrdinalIgnoreCase))
                return false;

            return Regex.IsMatch(line, @"^\s*Dim\s+", RegexOptions.IgnoreCase);
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
