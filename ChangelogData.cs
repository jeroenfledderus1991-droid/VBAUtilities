namespace VBEAddIn
{
    internal static class ChangelogData
    {
        internal const string CurrentVersion = "1.3.2";

        // -----------------------------------------------------------------------
        // AGENT INSTRUCTIE: Bij elke nieuwe release voeg je een nieuw item toe
        // BOVENAAN deze array (nieuwste versie eerst).
        // Update ook CHANGELOG.md met dezelfde informatie.
        // Prefix-conventie voor regels:
        //   "+"  = nieuw toegevoegd
        //   "*"  = fix / verbetering
        //   "-"  = verwijderd
        // -----------------------------------------------------------------------
        internal static readonly ChangelogEntry[] Entries = new[]
        {
            new ChangelogEntry("1.3.2", "2026-04-05", new[]
            {
                "+ Testrelease 1.3.2 — interne testversie voor updatecontrole",
            }),
            new ChangelogEntry("1.3.0", "2026-04-04", new[]
            {
                "+ VBA Wachtwoord Verwijderen — bypass VBA projectwachtwoord via DialogBoxParamA API hooking",
                "+ Versiegeschiedenis — bekijk deze changelog direct vanuit de add-in",
                "* Fix: quote escaping hersteld in embedded VBA string",
                "* Fix: code-validatie nu hoofdletterongevoelig (VBE past varPtr → VarPtr aan)",
                "* Fix: geïnjecteerde module bleef niet staan — cleanup verwijderd uit finally-blok",
            }),
            new ChangelogEntry("1.2.0", "2026-02-20", new[]
            {
                "+ Code Library — importeer en beheer VBA modules via een centrale bibliotheek",
                "+ Export naar Library — exporteer modules naar de code library",
                "+ UnifiedCodeLibraryForm — gecombineerde interface voor library beheer",
                "+ LibraryPathsForm — configureer paden voor de code library",
                "+ Insert Comment — commentaarregel met timestamp (normaal / SHIFT / CTRL)",
            }),
            new ChangelogEntry("1.1.0", "2026-01-15", new[]
            {
                "+ Reference Manager — beheer VBA references via een UI",
                "+ WhoAmI — toon workbook FullName en ReadOnly status",
                "+ Optimalisatie UIT / AAN — schakel Excel optimalisaties uit/aan",
                "+ Export VBA Componenten — exporteer alle VBA modules naar schijf",
            }),
            new ChangelogEntry("1.0.0", "2026-01-01", new[]
            {
                "+ VBE COM Add-in structuur (.NET 4.8, COM Interop)",
                "+ Formatteer Dim Statements — sorteer en lijn Dims uit",
                "+ Formatteer Complete Code — volledige code formatting",
                "+ Instellingen form (bewaard in registry)",
                "+ CommandBar (toolbar) met instelbare knoppen",
            }),
        };
    }

    internal sealed class ChangelogEntry
    {
        internal string Version { get; private set; }
        internal string Date { get; private set; }
        internal string[] Lines { get; private set; }

        internal ChangelogEntry(string version, string date, string[] lines)
        {
            Version = version;
            Date = date;
            Lines = lines;
        }
    }
}
