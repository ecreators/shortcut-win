using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using IWshRuntimeLibrary;
using shortcut.model;

namespace shortcut.util
{
    // ReSharper disable once ClassNeverInstantiated.Global
    // ReSharper disable once ClassWithVirtualMembersNeverInherited.Global
    public class ShortcutUtil
    {
        // ReSharper disable once InconsistentNaming
        private const string MISSING_PARAMETERS = @"Parameter fehlen! Es muss wenigstens der Pfad des Links angegeben werden.
        
-l ""{Absoluter Pfad zum Programm}"" -> Wenn das Programm per Umgebungsvariablen hinterlegt ist kann auch dies verwendet werden.";

        // ReSharper disable once EmptyConstructor
        public ShortcutUtil() {

        }

        // ReSharper disable once UnusedMember.Global
        public static void CreateLink(ShortcutArgs args)
        {
            Main(args.ToArgs());
        }

        [STAThread]
        public static void Main(string[] args)
        {
            var util = new ShortcutUtil();
            util.run(args);
        }

        private void run(string[] arguments)
        {
            // Lädt die eingebettete Ressource nach, ohne die das erstellen unter Windows nicht möglich wäre
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) => {
                var resourceName = new AssemblyName(args.Name).Name + ".dll";
                var resource = Array.Find(GetType().Assembly.GetManifestResourceNames(), element => element.EndsWith(resourceName));

                Debug.WriteLine($"Search for dll: {resourceName}");

                using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resource)) {
                    Debug.Assert(stream != null, "stream != null");
                    var assemblyData = new byte[stream.Length];
                    stream.Read(assemblyData, 0, assemblyData.Length);
                    return Assembly.Load(assemblyData);
                }
            };

            Console.InputEncoding = Encoding.UTF8;
            Console.WriteLine("Parameter: " + string.Join(",", arguments));

            // P -help, /help, --help, -?, /?, --? oder gar keine Parameter
            // zeigt die Anleitung (Hilfe)
            
            var help = isParameterSet(arguments, "?", false);
            if (!help.IsSet) {
                help = isParameterSet(arguments, "help", false);
            }

            if (help.IsSet || arguments.Length == 0) {
                // show help
                Console.WriteLine(getHelpText());
                return;
            }

            // create

            var desktop = isParameterSet(arguments, "d", true, @"\.*");
            if (!desktop.IsSet) {
                desktop = isParameterSet(arguments, "desktop", true, @"\.*");
            }

            var name = isParameterSet(arguments, "n", true, @"\.*");
            if (!name.IsSet) {
                name = isParameterSet(arguments, "name", true, @"\.*");
            }

            var style = isParameterSet(arguments, "s", true, @"\d+");
            if (!style.IsSet) {
                style = isParameterSet(arguments, "style", true, @"\d+");
            }
            var styleInt = int.TryParse(style.Value, out int i) ? i : 0;

            var link = isParameterSet(arguments, "l", true, @"\.*");
            if (!link.IsSet) {
                link = isParameterSet(arguments, "link", true, @"\.*");
            }

            var param = isParameterSet(arguments, "p", true, @"\.*");
            if (!param.IsSet) {
                param = isParameterSet(arguments, "param", true, @"\.*");
            }

            var icon = isParameterSet(arguments, "i", true, @"\.*");
            if (!icon.IsSet) {
                icon = isParameterSet(arguments, "icon", true, @"\.*");
            }

            var workingDir = isParameterSet(arguments, "w", true, @"\.*");
            if (!workingDir.IsSet) {
                workingDir = isParameterSet(arguments, "wdir", true, @"\.*");
            }

            var desc = isParameterSet(arguments, "desc", true, @"\.*");

            var shortcut = new ShortcutArgs
            {
                Arguments = param.Value,
                Description = desc.Value ?? string.Empty,
                IconImageFilePath = icon.Value ?? string.Empty,
                WindowStyleConstant = styleInt,
                WorkingDirectory = workingDir.Value ?? string.Empty,
                TargetFileAssosiation = link.Value ?? string.Empty
            };

            if (desktop.IsSet) {
                var linkName = name.IsSet ? name.Value : desktop.Value;
                shortcut.createOnDesktop(linkName);
            } else {
                shortcut.LinkPath = name.Value;
            }

            if (canShortcut(shortcut)) {
                createShortcut(shortcut);
            } else {
                Console.WriteLine(MISSING_PARAMETERS);
            }
        }

        // absolutes Minimum an Parametern
        protected static bool canShortcut(ShortcutArgs shortcut)
        {
            return !isNullOrWhitespace(shortcut.TargetFileAssosiation);
        }

        // So wird System.CSharp nicht nötig als dll und damit net 3.5 tauglich
        protected static bool isNullOrWhitespace(string text)
        {
            return (text == null || text.Trim().Length == 0);
        }

        private void createShortcut(ShortcutArgs e)
        {
            // ohne wdir, wird das aktuelle Arbeitsverzeichnis benutzt
            if (isNullOrWhitespace(e.WorkingDirectory))
            {
                e.WorkingDirectory = Environment.CurrentDirectory + "\\";
            }
            
            // ohne -n "{abs. Pfad + Name}" oder -d "{Name auf Desktop}", wird die Verknüpfung beim Ziel abgelegt mit dem Namen des Ziels + .lnk
            if (isNullOrWhitespace(e.LinkPath))
            {
                e.LinkPath = Path.Combine(e.WorkingDirectory, Path.GetFileNameWithoutExtension("" + e.TargetFileAssosiation));
            }
            var shortcutLocation = e.LinkPath.EndsWith(".lnk") ? e.LinkPath : (e.LinkPath + ".lnk");

            var shortcut = prepareShortcutClass(shortcutLocation);

            shortcut.Description = e.Description;
            shortcut.TargetPath = e.TargetFileAssosiation;
            shortcut.Arguments = e.Arguments;
            shortcut.WindowStyle = e.WindowStyleConstant;
            shortcut.WorkingDirectory = e.WorkingDirectory;
            try
            {
                // Keine Angabe => z.B. "[exe],0" oder File
                var shortcutIconLocation = isNullOrWhitespace(e.IconImageFilePath) ? (e.TargetFileAssosiation + ",0") : e.IconImageFilePath;
                Console.WriteLine($"Teste Icon '{shortcutIconLocation}'");
                e.IconImageFilePath = shortcutIconLocation;
                shortcut.IconLocation = shortcutIconLocation;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                shortcut.IconLocation = null;
            }

            Console.WriteLine($"{{\r\n\t-n: \"{e.LinkPath}\",\r\n\t-w: \"{e.WorkingDirectory}\",\r\n\t-l: \"{e.TargetFileAssosiation}\",\r\n\t-param: \"{e.Arguments}\",\r\n\t-i: \"{e.IconImageFilePath}\",\r\n\t-desc: \"{e.Description}\"\r\n}}");
            onSaveShortcut(ref shortcut);
            
            // anlegen der lnk
            shortcut.Save();
        }

        protected static IWshShortcut prepareShortcutClass(string shortcutLocation)
        {
            var shell = new WshShell();
            var shortcut = (IWshShortcut) shell.CreateShortcut(shortcutLocation);
            return shortcut;
        }

        protected virtual void onSaveShortcut(ref IWshShortcut shortcut)
        {
        }

        // Bestimmt Parameter und ggf. Wert
        public static Tuple<bool, string> isParameterSet(string[] args, string parameter, bool requiresValue, string regexValue = null)
        {
            for (var i = 0; i < args.Length; i++)
            {
                var arg = args[i];
                if (arg.ToLower().Equals("/" + parameter.ToLower())
                    || arg.ToLower().Equals("-" + parameter.ToLower())
                    || arg.ToLower().Equals("--" + parameter.ToLower()))
                {
                    if (requiresValue && i + 1 >= args.Length)
                    {
                        throw new ArgumentException($"Es fehlt die Wert angabe zu Parameter '{parameter}'");
                    }

                    if (regexValue != null && !Regex.IsMatch(args[i + 1], regexValue))
                    {
                        throw new ArgumentException($"Die Wert enspricht nicht dem erwarteten Pattern '{regexValue}'! Es wurde angewendet auf den Wert '{(args.Length < i + 1 ? args[i + 1] : "leer")}'");
                    }
                    return new Tuple<bool, string>(true, (requiresValue ? args[i + 1].Replace("\"", "") : null));
                }
            }
            return new Tuple<bool, string>(false, null);
        }

        private static string getHelpText()
        {
            return @"shortcut.exe author: Bjoern Frohberg, MyData GmbH Version 1.1 @10.02.2017
 
+----------------------+
| Hilfe                |
+----------------------+
 
Parameter:
 
desktop d D  - Verknuepfung auf dem Desktop mit dem Namen [name]
name n N     - Ohne [desktop] ist dies der absolute Pfad. Mit [desktop] ist dies nur der Name
style s S    - Stil des Fensters der Anwendung (Zahlenwert 0 bis 4)
link l L     - Ziel des Links (absoluter Pfad + Datei)
param p P    - Linkparameter
icon i I     - Bild des Links (absoluter Pfad + Bild oder z.B. App.exe,0)
wdir w W     - Arbeitsverzeichnis der Anwendung hinter dem Link (wenn verwendet, kann die [link] Angabe relativ hierzu sein
desc         - Beschreibungstext des Links
 
Allgemeine Anmerkung 
 
1. Alle Parameter (nicht Werte) dürfen mit - -- oder / beginnen.
2. Die Gross- und Kleinschreibung wird bei den Parametern ignoriert
3. Die mit Parametern angegebenen Werte beduerfen keiner Reihenfolge

Achtung: Das Definieren eines Shortcut-Keys wird nicht unterstuetzt! 
 
+----------------------+
| Parameter            |
+----------------------+
 
Alle Angaben müssen in einfachen oder normalen Anfuehrungszeichen stehen!

1.a [Desktop]   - Link-Name auf Desktop: Name der Verkuepfung, aber ohne lnk
                  Parameter: ""d"" oder ""D"" oder ""desktop""
                  Beispiel -d 'Verknuepfung'
                  
                  ODER
                  
1.b [Name]      - Link-Name: Absoluter Pfad zum Link inklusive Name, aber ohne lnk. Am Zielordner sind Schreibrechte nötig!                  
                  Der Ordner, muss bereits existieren!
                  Parameter: ""n"" oder ""N"" oder ""name""
                  Beispiel -n 'c:\\Verknuepfung'
 
2. [Link]       - Link-Angabe: Absoluter Pfad zur Zieldatei
                  Parameter: ""l"" oder ""L"" oder ""link""
                  Beispiel -l 'c:\\file.exe'
 
3. [Link-Param] - Link-Angabe-Parameter: Parameter, die dem Link mitgegeben werden
                  Parameter: ""p"" oder ""P"" oder ""param""
                  Beispiel -p '-debug ""x""'
 
4. [Icon]       - Icon-Angabe: Absoluter Pfad zu einer Bilddatei (ico) oder einem eingebetteten Icon per z.B. 'notepad.exe, 0'
                               Der Pfad kann auch relative zum Arbeitsverzeichnis sein.
                  Parameter: ""i"" oder ""I"" oder ""icon""
                  Beispiel -i 'c:\\file.exe, 0'
 
5. [WorkingDir] - Arbeitsverzeichnis: Absoluter Pfad ist hier Pflicht!
                  Parameter: ""w"" oder ""W"" oder ""wdir""
                  Beispiel -w 'c:\\'
 
6. [Desc]       - Link Beschreibung: Ein freier Text. Einzeilig!
                  Parameter: ""desc""
                  Beispiel -desc 'Meine Verknuepfung ist das!'
                  
7. [Style]      - Fensterstil: 0, 1, 4 ohne Anführungszeichen.
                  0 = Normales Fenster
                  1 = Minimiert
                  4 = Maximiert
                  Parameter: ""s"" oder ""S"" oder ""style""
                  Beispiel -s 0
                  
                  
Ein Link kann bei der Zieldatei erzeugt werden mit nzr dem Parameter -l.

In diesem Fall ist
-wdir das aktuelle Verzeichnis zum Ziel
-icon wird versucht auf ""Ziel,0""
-name wird das -wdir + Dateiname von -l + .lnk
-style wird 0
";
        }
    }

    public class ShortcutArgs
    {
        public string LinkPath { get; set; }
        public string Description { get; set; }
        public string IconImageFilePath { get; set; }
        public string TargetFileAssosiation { get; set; }
        public string Arguments { get; set; }
        public int WindowStyleConstant { get; set; }
        public string WorkingDirectory { get; set; }

        public void createOnDesktop(string linkName)
        {
            LinkPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), linkName);
        }

        public string[] ToArgs() => new[] {"-n", LinkPath, "-i", IconImageFilePath, "-l", TargetFileAssosiation, "-p", Arguments, "-s", WindowStyleConstant.ToString(), "-w", WorkingDirectory, "-desc", Description};
    }
}
