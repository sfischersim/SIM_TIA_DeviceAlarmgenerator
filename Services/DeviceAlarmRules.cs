using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace SIM_TIA_DeviceAlarmgenerator.Services.DeviceRules
{
    // Hinweis: IDeviceAlarmRule wird an anderer Stelle definiert.
    // Die nachfolgenden Implementierungen verwenden <inheritdoc />,
    // damit die Doku zentral aus dem Interface übernommen wird.

    // -----------------------------
    // OMODE
    // -----------------------------
    /// <summary>
    /// Regelwerk für den Gerätetyp „Omode“ mit vordefiniertem Fehler-Template.
    /// </summary>
    public sealed class OmodeRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "Omode";

        /// <inheritdoc />
        public int AlarmsPerDevice => 11;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Sammelfehler HW & CFG","Errors"),
            ("Sammelstörung","Errors"),
            ("Fehler .2","Errors"),
            ("Fehler .3","Errors"),
            ("Fehler .4","Errors"),
            ("Fehler .5","Errors"),
            ("Fehler .6","Errors"),
            ("Konfigurationsfehler","Errors"),
            ("Kommunikationsfehler","Errors"),
            ("IO Fehler","Errors"),
            ("Geräte Fehler","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // ALM16
    // -----------------------------
    /// <summary>
    /// Regelwerk für „Alm16“. Die Meldungstexte (und optional Klassen) werden
    /// aus dem Sheet „AppAlarm (_AA)“ bezogen.
    /// 
    /// Zuordnung (wie im Alt-Tool):
    ///   AppRow = startRow (Default 2) + (DeviceIndex1Based - 1) * 16 + AlarmIndex0Based
    ///   Text   = (RowId + " " + appSheet[AppRow, commentCol])
    ///   Klasse = aus appSheet[AppRow, classCol] (Code "3" -> "Warnings", sonst "Errors"),
    ///            für ErrorsTemplate initial nur für Bit 0..15 (Default "Errors").
    /// </summary>
    public sealed class Alm16Rule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "Alm16";

        /// <inheritdoc />
        public int AlarmsPerDevice => 16;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate { get; }

        // --- AppAlarm-Sheet Bezug ---
        private readonly string[,] _appSheet;
        private readonly int _startRow;     // Default 2 (0-basiert im Code, fachlich „ab Zeile 3“)
        private readonly int _nameCol;      // Default 0
        private readonly int _classCol;     // Default 3  (Alt-Tool: "3" => Warnings)
        private readonly int _commentCol;   // Default 7

        /// <summary>
        /// Erstellt die Alm16-Rule, die Texte/Klasse aus dem AppAlarm-Sheet bezieht.
        /// </summary>
        /// <param name="appSheet">Matrix des Sheets „AppAlarm (_AA)“ als [row, col].</param>
        /// <param name="startRow">Erste Datenzeile (0-basiert). Fachlich: i. d. R. 2 → ab Excel-Zeile 3.</param>
        /// <param name="nameCol">Spalte für den Namen/Key (Default 0).</param>
        /// <param name="classCol">Spalte für Klassen-/Severity-Code (Default 3; "3" → "Warnings").</param>
        /// <param name="commentCol">Spalte für Kommentar/Meldungstext (Default 7).</param>
        public Alm16Rule(
            string[,] appSheet,
            int startRow = 2,
            int nameCol = 0,
            int classCol = 3,
            int commentCol = 7)
        {
            _appSheet = appSheet ?? throw new ArgumentNullException(nameof(appSheet));
            _startRow = startRow;
            _nameCol = nameCol;
            _classCol = classCol;
            _commentCol = commentCol;

            // ErrorsTemplate initial für Bits 0..15 aus den ersten 16 App-Zeilen (falls vorhanden) ableiten.
            // Ansonsten "Errors". Der Text im Template ist hier nur informativ; der finale Zeilentext
            // wird in ComposeMessage(...) pro Instanz/Bit aus dem passenden AppRow erzeugt.
            ErrorsTemplate = new (string Text, string Class)[AlarmsPerDevice];
            for (int bit = 0; bit < AlarmsPerDevice; bit++)
            {
                var (txt, cls) = TryReadFromApp(bit, deviceIndex1Based: 1, rowId: 0, forTemplate: true);
                if (string.IsNullOrWhiteSpace(cls)) cls = "Errors";
                if (string.IsNullOrWhiteSpace(txt)) txt = $"Fehler {bit}";
                ErrorsTemplate[bit] = (txt, cls);
            }
        }

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            // 1) Text-Override aus AppAlarm (_AA) holen – OHNE RowId-Prefix
            var (overrideStdText, _) =
                TryReadFromApp(ctx.AlarmIndex0Based, ctx.DeviceIndex1Based, rowId: 0, forTemplate: true);

            // 2) Fallback auf das Template (ErrorsTemplate) wenn im Sheet nichts steht
            var (tplText, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            string stdText = string.IsNullOrWhiteSpace(overrideStdText) ? tplText : overrideStdText;

            // 3) Finale Meldung exakt wie bei Omode/SQ etc. zusammensetzen
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, stdText);
        }

        /// <summary>
        /// Liest Name/Text/Klasse aus dem AppAlarm-Sheet für die gegebene (Instanz, Bit)-Position.
        /// </summary>
        /// <param name="bit">0..15</param>
        /// <param name="deviceIndex1Based">Instanz (1-basiert)</param>
        /// <param name="rowId">Globale RowId (für Prefix); 0 = ohne Prefix (Template)</param>
        /// <param name="forTemplate">Wenn true, wird nur ein generischer Zeilentext ohne RowId gebildet (für ErrorsTemplate).</param>
        private (string Text, string Class) TryReadFromApp(int bit, int deviceIndex1Based, int rowId, bool forTemplate)
        {
            // AppRow wie im Alt-Tool:
            //   startRow + (Instanz-1)*16 + Bit
            int appRow = _startRow + (deviceIndex1Based - 1) * AlarmsPerDevice + bit;

            int rows = _appSheet.GetLength(0);
            int cols = _appSheet.GetLength(1);
            if (appRow < 0 || appRow >= rows) return (string.Empty, string.Empty);

            string Safe(int r, int c)
            {
                if (c < 0 || c >= cols) return "";
                var v = _appSheet[r, c];
                return string.IsNullOrWhiteSpace(v) ? "" : v.Trim();
            }

            string name = Safe(appRow, _nameCol);
            string comment = Safe(appRow, _commentCol);
            string clsCode = Safe(appRow, _classCol);

            // Klassenmapping wie im Alt-Tool: Code "3" ⇒ "Warnings", sonst "Errors".
            string @class = (clsCode == "3") ? "Warnings" : "Errors";

            // Text: im Template-Fall bauen wir einen neutralen Text ohne RowId auf (nur für Preview im Template).
            // Im realen ComposeMessage-Fall prefixen wir die RowId (wie im Alt-Tool).
            string textForTemplate = !string.IsNullOrWhiteSpace(comment) ? comment
                                : (!string.IsNullOrWhiteSpace(name) ? name : $"Fehler {bit}");

            if (forTemplate)
                return (textForTemplate, @class);

            string textFinal = !string.IsNullOrWhiteSpace(comment) ? comment
                             : (!string.IsNullOrWhiteSpace(name) ? name : $"Fehler {bit}");

            // RowId-Prefix wie im Alt-Tool (z. B. "123 Mein Text")
            if (rowId > 0)
                textFinal = $"{rowId} {textFinal}";

            return (textFinal, @class);
        }
    }

    // -----------------------------
    // SQ
    // -----------------------------
    /// <summary>
    /// Regelwerk für den Schrittkettentyp „SQ“ (Sequencer) inklusive Zeitüberschreitungen und Bedienhinweisen.
    /// </summary>
    public sealed class SqRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "SQ";

        /// <inheritdoc />
        public int AlarmsPerDevice => 8;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Interlock Fehler","Errors"),
            ("Gesamtlaufzeit überschritten","Errors"),
            ("Schritt-Laufzeit überschritten","Errors"),
            ("Fehlerdialog","Errors"),
            ("Bediener-Dialog benötigt","Warnings"),
            ("Grundstellung nicht erreicht","Warnings"),
            ("AutoStop nicht erreicht","Warnings"),
            ("Systemfehler","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // 2DS
    // -----------------------------
    /// <summary>
    /// Regelwerk für zweikanalige Sensorik „2DS“ mit Sammelfehlern und Kommunikations-/IO-Aspekten.
    /// </summary>
    public sealed class TwoDsRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "2DS";

        /// <inheritdoc />
        public int AlarmsPerDevice => 8;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Sammelfehler HW & CFG","Errors"),
            ("Sammelstörung","Errors"),
            ("Sensorfehler","Errors"),
            ("Sensorfehler","Errors"),
            ("Kommunikationsfehler","Errors"),
            ("IO Fehler","Errors"),
            ("Geräte Fehler","Errors"),
            ("Konfigurationsfehler","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // 2DSPM
    // -----------------------------
    /// <summary>
    /// Regelwerk für „2DSPM“ mit erweitertem Sensorfehler-Spektrum.
    /// </summary>
    public sealed class TwoDsPmRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "2DSPM";

        /// <inheritdoc />
        public int AlarmsPerDevice => 11;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Sammelfehler HW & CFG","Errors"),
            ("Sammelstörung","Errors"),
            ("Sensorfehler","Errors"),
            ("Sensorfehler","Errors"),
            ("Sensorfehler","Errors"),
            ("Sensorfehler","Errors"),
            ("Sensorfehler","Errors"),
            ("Kommunikationsfehler","Errors"),
            ("IO Fehler","Errors"),
            ("Geräte Fehler","Errors"),
            ("Konfigurationsfehler","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // AI
    // -----------------------------
    /// <summary>
    /// Regelwerk für Analogeingänge „AI“ inkl. Grenzwertverletzungen und Qualitätshinweisen.
    /// </summary>
    public sealed class AiRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "AI";

        /// <inheritdoc />
        public int AlarmsPerDevice => 13;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Sammelfehler HW & CFG","Errors"),
            ("Sammelstörung","Errors"),
            ("Datentyp Rohwert unzulässig","Errors"),
            ("Rohwert außerhalb Messbereich","Errors"),
            ("Wertesprung","Errors"),
            ("MinMin unterschritten","Errors"),
            ("MaxMax überschritten","Errors"),
            ("Min unterschritten","Warnings"),
            ("Max überschritten","Warnings"),
            ("Kommunikationsfehler","Errors"),
            ("IO Fehler","Errors"),
            ("Geräte Fehler","Errors"),
            ("Konfigurationsfehler","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // DO
    // -----------------------------
    /// <summary>
    /// Regelwerk für Digitale Ausgänge „DO“ inkl. Laufzeit- und Vakuumüberwachung.
    /// </summary>
    public sealed class DoRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "DO";

        /// <inheritdoc />
        public int AlarmsPerDevice => 9;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Konfigurationsfehler","Errors"),
            ("Sensorfehler","Errors"),
            ("Sensorfehler","Errors"),
            ("Laufzeitüberwachung","Errors"),
            ("Fehler Vakuum","Errors"),
            ("IO Fehler","Errors"),
            ("Geräte Fehler","Errors"),
            ("Konfigurationsfehler","Errors"),
            ("Allgemeiner Fehler","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // DI
    // -----------------------------
    /// <summary>
    /// Regelwerk für Digitale Eingänge „DI“ mit typischen Signal- und Anfangszustandsprüfungen.
    /// </summary>
    public sealed class DiRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "DI";

        /// <inheritdoc />
        public int AlarmsPerDevice => 8;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Konfigurationsfehler","Errors"),
            ("Fehler erwartet Sensor nicht geschaltet","Errors"),
            ("Fehler erwartet Sensor geschaltet","Errors"),
            ("Falsches Anfangssignal","Errors"),
            ("Signal nicht wie erwartet","Errors"),
            ("Kommunikationsfehler","Errors"),
            ("Hardwarefehler","Errors"),
            ("Gerätefehler","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // APOSI
    // -----------------------------
    /// <summary>
    /// Regelwerk für Positionsmodule „APosi“ mit Applikations- und HW-/Modulzuständen.
    /// </summary>
    public sealed class AposiRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "APosi";

        /// <inheritdoc />
        public int AlarmsPerDevice => 4;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Applikationswarnung","Warnings"),
            ("Applikationsfehler","Errors"),
            ("Modul / HW Warnung","Warnings"),
            ("Modul / HW Fehler","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // FEEDS
    // -----------------------------
    /// <summary>
    /// Regelwerk für Zuführsysteme „Feeds“ mit Teilmangel- und Laufzeitmeldungen.
    /// </summary>
    public sealed class FeedsRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "Feeds";

        /// <inheritdoc />
        public int AlarmsPerDevice => 7;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Teilmangel Wendelförderer","Warnings"),
            ("Teilmangel Bunker","Warnings"),
            ("Teilmangel Kopfstück","Warnings"),
            ("Teilmangel Min","Warnings"),
            ("Teilmangel Max","Warnings"),
            ("Laufzeit IN1","Warnings"),
            ("Laufzeit IN2","Warnings")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // BC
    // -----------------------------
    /// <summary>
    /// Regelwerk für Bildverarbeitung „BC“ mit Sammelfehlern und Kommunikationsaspekten.
    /// </summary>
    public sealed class BcRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "BC";

        /// <inheritdoc />
        public int AlarmsPerDevice => 7;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Sammelfehler HW & CFG","Errors"),
            ("Sammelstörung","Errors"),
            ("Fehler Auswertung","Errors"),
            ("Kommunikationsfehler","Errors"),
            ("IO Fehler","Errors"),
            ("Allgemeiner Fehler","Errors"),
            ("Konfigurationsfehler","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // CAM_XG
    // -----------------------------
    /// <summary>
    /// Regelwerk für Kamerasysteme „Cam_XG“ mit Auswertungs- und IO-/Kommunikationsfehlern.
    /// </summary>
    public sealed class CamXgRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "Cam_XG";

        /// <inheritdoc />
        public int AlarmsPerDevice => 7;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Sammelfehler HW & CFG","Errors"),
            ("Sammelstörung","Errors"),
            ("Fehler Auswertung","Errors"),
            ("Kommunikationsfehler","Errors"),
            ("IO Fehler","Errors"),
            ("Allgemeiner Fehler","Errors"),
            ("Konfigurationsfehler","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // CNVM
    // -----------------------------
    /// <summary>
    /// Regelwerk für Förder- oder Handhabungsmodul „CnvM“ mit Laufzeitüberwachungen und Befehlsprüfung.
    /// </summary>
    public sealed class CnvMRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "CnvM";

        /// <inheritdoc />
        public int AlarmsPerDevice => 11;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Laufzeit Verlassen der Belegung überschritten","Errors"),
            ("Laufzeit Neue Belegung überschritten","Errors"),
            ("Laufzeit Hub überschritten","Errors"),
            ("Laufzeit Stopper überschritten","Errors"),
            ("Unzulässiges Kommando","Errors"),
            ("Fehlende Freigabe fürs Kommando","Errors"),
            ("Unzulässige Änderung Belegtstatus","Errors"),
            ("Kommunikationsfehler","Errors"),
            ("IO Fehler","Errors"),
            ("Allgemeiner Fehler","Errors"),
            ("Konfigurationsfehler","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // COM_RS
    // -----------------------------
    /// <summary>
    /// Regelwerk für serielle Kommunikation „Com_RS“ inkl. Zeitüberschreitung und Sammelwarnung.
    /// </summary>
    public sealed class ComRsRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "Com_RS";

        /// <inheritdoc />
        public int AlarmsPerDevice => 11;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Sammelfehler HW & CFG","Errors"),
            ("Kommunikationsfehler","Errors"),
            ("IO Fehler","Errors"),
            ("Allgemeiner Fehler","Errors"),
            ("Konfigurationsfehler","Errors"),
            ("Sammelstörung","Errors"),
            ("Timeout Komm.-Auftrag","Errors"),
            ("Keine gültige Antwort","Errors"),
            ("Störung Komm-Prozessor","Errors"),
            ("Sammelstörung","Errors"),
            ("Sammelwarnung","Warnings")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // CS
    // -----------------------------
    /// <summary>
    /// Regelwerk für den Gerätetyp „CS“ (mehrteilige Komponenten: Hardware, CS, Connector, PD).
    /// </summary>
    public sealed class CsRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "CS";

        /// <inheritdoc />
        public int AlarmsPerDevice => 16;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Konfigurationsfehler Hardware (Details siehe Device)","Errors"),
            ("Systemfehler Hardware (Details siehe Device)","Errors"),
            ("Ablauffehler Hardware (Details siehe Device)","Errors"),
            ("General Fault Hardware (Details siehe Device)","Errors"),
            ("Konfigurationsfehler CS (Details siehe Device)","Errors"),
            ("Systemfehler CS (Details siehe Device)","Errors"),
            ("Ablauffehler CS (Details siehe Device)","Errors"),
            ("Allgemeiner Fehler CS (Details siehe Device)","Errors"),
            ("Konfigurationsfehler Connector (Details siehe Device)","Errors"),
            ("Systemfehler Connector (Details siehe Device)","Errors"),
            ("Ablauffehler Connector (Details siehe Device)","Errors"),
            ("Allgemeiner Fehler Connector (Details siehe Device)","Errors"),
            ("Konfigurationsfehler PD (Details siehe Device)","Errors"),
            ("Systemfehler PD (Details siehe Device)","Errors"),
            ("Ablauffehler PD (Details siehe Device)","Errors"),
            ("Allgemeiner Fehler PD (Details siehe Device)","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // HT_KSV
    // -----------------------------
    /// <summary>
    /// Regelwerk für Heizung/Temperatur „HT_KSV“ mit Fühler- und Loop-Alarmen.
    /// </summary>
    public sealed class HtKsvRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "HT_KSV";

        /// <inheritdoc />
        public int AlarmsPerDevice => 8;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Sammelfehler HW & CFG","Errors"),
            ("Sammelstörung","Errors"),
            ("Loopalarm","Errors"),
            ("Fühlerfehler","Errors"),
            ("Kommunikationsfehler","Errors"),
            ("IO Fehler","Errors"),
            ("Allgemeiner Fehler","Errors"),
            ("Konfigurationsfehler","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // KEYENCE_MC
    // -----------------------------
    /// <summary>
    /// Regelwerk für Keyence-Markiercontroller „Keyence_MC“ mit Befehls-/Programmprüfungen.
    /// </summary>
    public sealed class KeyenceMcRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "Keyence_MC";

        /// <inheritdoc />
        public int AlarmsPerDevice => 7;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Konfigurationsfehler","Errors"),
            ("Laufzeit überschritten","Errors"),
            ("Falsches Kommando","Errors"),
            ("Falsches Programm","Errors"),
            ("Kommunikationsfehler","Errors"),
            ("IO Fehler","Errors"),
            ("Geräte Fehler","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // LOAD_DISPL
    // -----------------------------
    /// <summary>
    /// Regelwerk für Last-/Anzeigeeinheit „Load_Displ“ mit Überwachung und Programm-/Befehlsprüfungen.
    /// </summary>
    public sealed class LoadDisplRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "Load_Displ";

        /// <inheritdoc />
        public int AlarmsPerDevice => 8;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Konfigurationsfehler","Errors"),
            ("Laufzeit überschritten","Errors"),
            ("Falsches Kommando","Errors"),
            ("Falsches Programm","Errors"),
            ("Überwachung hat angesprochen","Errors"),
            ("Kommunikationsfehler","Errors"),
            ("IO Fehler","Errors"),
            ("Geräte Fehler","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // MM_STORE
    // -----------------------------
    /// <summary>
    /// Regelwerk für „MM_Store“ (Material-/Datenablage) mit generischem Fehler.
    /// </summary>
    public sealed class MmStoreRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "MM_Store";

        /// <inheritdoc />
        public int AlarmsPerDevice => 1;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Kein spezifischer Fehler definiert","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // PN
    // -----------------------------
    /// <summary>
    /// Regelwerk für Prüf-/Prozessmodul „PN“ mit Laufzeit- und Hardwarefehlern sowie Warnungen.
    /// </summary>
    public sealed class PnRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "PN";

        /// <inheritdoc />
        public int AlarmsPerDevice => 7;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Laufzeitfehler","Errors"),
            ("Falsches Kommando","Errors"),
            ("Falsche Programmnummer","Errors"),
            ("Warnung","Warnings"),
            ("Kommunikationsfehler","Errors"),
            ("Hardwarefehler","Errors"),
            ("Device nicht bereit","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // RDIF
    // -----------------------------
    /// <summary>
    /// Regelwerk für RFID-System „RDIF“ mit Auswertung, IO und Kommunikation.
    /// </summary>
    public sealed class RdifRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "RDIF";

        /// <inheritdoc />
        public int AlarmsPerDevice => 7;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Sammelfehler HW & CFG","Errors"),
            ("Sammelstörung","Errors"),
            ("Fehler Auswertung","Errors"),
            ("Kommunikationsfehler","Errors"),
            ("IO Fehler","Errors"),
            ("Allgemeiner Fehler","Errors"),
            ("Konfigurationsfehler","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // RIT_WEIS
    // -----------------------------
    /// <summary>
    /// Regelwerk für Sicherheits-/Positionssystem „RIT_WEIS“ mit Safety- und Timeout-Alarmen.
    /// </summary>
    public sealed class RitWeisRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "RIT_WEIS";

        /// <inheritdoc />
        public int AlarmsPerDevice => 13;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Sammelfehler HW & CFG","Errors"),
            ("Allgemeiner Fehler","Errors"),
            ("Störung Schutz","Errors"),
            ("Alarm Timeout","Errors"),
            ("Alarm Position überfahren","Errors"),
            ("Alarm Safety","Errors"),
            ("Alarm Summe Störung","Errors"),
            ("Alarm Summe Warnung","Warnings"),
            ("Alarm Zwangsdynamisierung","Errors"),
            ("Kommunikationsfehler","Errors"),
            ("IO Fehler","Errors"),
            ("Device Fehler","Errors"),
            ("Konfigurationsfehler","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // VISION_S
    // -----------------------------
    /// <summary>
    /// Regelwerk für Vision-Sensorik „VisionS“ mit Kommunikations- und IO-Fehlern.
    /// </summary>
    public sealed class VisionSRule : IDeviceAlarmRule
    {
        /// <inheritdoc />
        public string DeviceTypeName => "VisionS";

        /// <inheritdoc />
        public int AlarmsPerDevice => 5;

        /// <inheritdoc />
        public (string Text, string Class)[] ErrorsTemplate => new (string, string)[]
        {
            ("Konfigurationsfehler","Errors"),
            ("Laufzeit überschritten","Errors"),
            ("Kommunikationsfehler","Errors"),
            ("IO Fehler","Errors"),
            ("Geräte Fehler","Errors")
        };

        /// <inheritdoc />
        public string ComposeMessage(DeviceMessageContext ctx)
        {
            var (txt, _) = ErrorsTemplate[ctx.AlarmIndex0Based];
            return GenericDeviceAlarmBuilder.ComposeDefaultMessage(ctx, DeviceTypeName, txt);
        }
    }

    // -----------------------------
    // DEVICE ALARM RULE FACTORY
    // -----------------------------

    /// <summary>
    /// Zentrale Fabrikklasse zur Bereitstellung aller bekannten <see cref="IDeviceAlarmRule"/>-Implementierungen.
    /// Erkennt neue Regelklassen automatisch über Reflection (Auto-Discovery).
    /// </summary>
    /// <remarks>
    /// Die Auto-Discovery lädt alle nicht-abstrakten Klassen im aktuellen AppDomain-Kontext,
    /// die <see cref="IDeviceAlarmRule"/> implementieren. Instanzierungsfehler werden bewusst ignoriert.
    /// </remarks>
    public static class DeviceAlarmRuleFactory
    {
        private static readonly Dictionary<string, IDeviceAlarmRule> _rules =
            new(StringComparer.OrdinalIgnoreCase);

        private static bool _initialized;

        /// <summary>
        /// Initialisiert die Fabric und lädt alle <see cref="IDeviceAlarmRule"/>-Implementierungen
        /// automatisch aus dem aktuellen Assembly-Kontext.
        /// </summary>
        public static void AutoDiscoverRules()
        {
            if (_initialized) return;

            var ruleType = typeof(IDeviceAlarmRule);

            // Alle Klassen im gleichen Assembly, die IDeviceAlarmRule implementieren
            var ruleTypes = AppDomain.CurrentDomain
                .GetAssemblies()
                .SelectMany(a =>
                {
                    try { return a.GetTypes(); }
                    catch (ReflectionTypeLoadException ex) { return ex.Types.Where(t => t != null)!; }
                })
                .Where(t => t is { IsClass: true, IsAbstract: false } && ruleType.IsAssignableFrom(t))
                .ToList();

            foreach (var type in ruleTypes)
            {
                try
                {
                    if (Activator.CreateInstance(type) is IDeviceAlarmRule instance &&
                        !_rules.ContainsKey(instance.DeviceTypeName))
                    {
                        _rules.Add(instance.DeviceTypeName.Trim(), instance);
                    }
                }
                catch
                {
                    // Wenn eine Rule nicht instanziiert werden kann, wird sie übersprungen.
                }
            }

            _initialized = true;
        }

        /// <summary>
        /// Gibt alle aktuell registrierten Regelinstanzen zurück.
        /// </summary>
        /// <remarks>
        /// Führt bei erstmaligem Zugriff automatisch die Auto-Discovery aus.
        /// </remarks>
        public static IEnumerable<IDeviceAlarmRule> GetAllRules()
        {
            if (!_initialized) AutoDiscoverRules();
            return _rules.Values;
        }

        /// <summary>
        /// Versucht, eine passende Regel anhand des Gerätetyps zu ermitteln.
        /// </summary>
        /// <param name="deviceType">Gerätetypbezeichnung (z. B. "Omode", "2DSPM", "AI").</param>
        /// <param name="rule">Gibt die gefundene Regelinstanz zurück, falls vorhanden.</param>
        /// <returns><c>true</c>, wenn eine Regel gefunden wurde, andernfalls <c>false</c>.</returns>
        /// <remarks>
        /// Zusätzliche Endzeichen wie Unterstrich/Leerzeichen werden toleriert und entfernt.
        /// </remarks>
        public static bool TryGetRule(string deviceType, out IDeviceAlarmRule rule)
        {
            if (!_initialized) AutoDiscoverRules();

            rule = null!;
            if (string.IsNullOrWhiteSpace(deviceType))
                return false;

            string key = deviceType.TrimEnd('_', ' ');
            return _rules.TryGetValue(key, out rule!);
        }
    }
}
