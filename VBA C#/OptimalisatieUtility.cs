using System;
using System.Runtime.InteropServices;

namespace VBEAddIn
{
    /// <summary>
    /// Utility voor Excel optimalisaties (uit/aan zetten)
    /// </summary>
    public static class OptimalisatieUtility
    {
        /// <summary>
        /// Zet Excel optimalisaties UIT (events, screenupdating, alerts, etc.)
        /// </summary>
        public static void ZetUit()
        {
            try
            {
                var excelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
                
                if (excelApp == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Excel applicatie niet gevonden.",
                        "Optimalisatie UIT",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                excelApp.EnableEvents = false;
                excelApp.ScreenUpdating = false;
                excelApp.DisplayAlerts = false;
                excelApp.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait;
                
                if (excelApp.ActiveWorkbook != null)
                {
                    excelApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual;
                }

                // Toon status in Excel StatusBar (onderaan Excel venster)
                excelApp.StatusBar = "Optimalisatie UIT: Events/Updating/Alerts uit, Calculation=Manual";
                
                // Wacht 2 seconden en reset statusbar
                System.Threading.Tasks.Task.Delay(2000).ContinueWith(_ => 
                {
                    try { excelApp.StatusBar = false; } catch { }
                });
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Fout bij Optimalisatie UIT: " + ex.Message,
                    "Fout",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Zet Excel optimalisaties weer AAN
        /// </summary>
        public static void ZetAan()
        {
            try
            {
                var excelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
                
                if (excelApp == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Excel applicatie niet gevonden.",
                        "Optimalisatie AAN",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                excelApp.ScreenUpdating = true;
                excelApp.StatusBar = false;
                excelApp.DisplayAlerts = true;
                excelApp.EnableEvents = true;
                excelApp.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault;
                
                if (excelApp.ActiveWorkbook != null)
                {
                    excelApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic;
                    excelApp.Calculate();
                }

                // Toon status in Excel StatusBar (onderaan Excel venster)
                excelApp.StatusBar = "Optimalisatie AAN: Alles hersteld, Calculation=Automatic";
                
                // Wacht 2 seconden en reset statusbar
                System.Threading.Tasks.Task.Delay(2000).ContinueWith(_ => 
                {
                    try { excelApp.StatusBar = false; } catch { }
                });
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Fout bij Optimalisatie AAN: " + ex.Message,
                    "Fout",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}
