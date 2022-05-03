using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using DotNetDBF;
using Spire.Doc;

namespace DbfProtocolsParser;

internal static class Program
{
    private static int _consumption;
    private static ProtocolIndexes[]? _indexes;
    private static readonly Dictionary<string, MemoryStream> TemplateCache = new();

    private static readonly string[] Mouth =
    {
        "января",
        "февраля",
        "марта",
        "апреля",
        "мая",
        "июня",
        "июля",
        "августа",
        "сентября",
        "октября",
        "ноября",
        "декабря"
    };

    private static void Main(string[] args)
    {
        if (!File.Exists("templates/indexes.json"))
        {
            var resources = new Dictionary<string, byte[]>
            {
                {"mi.docx", Resources.mi},
                {"mi_r.docx", Resources.mi_r},
                {"gost.docx", Resources.gost},
                {"gost_r.docx", Resources.gost_r},
                {"strk.docx", Resources.strk},
                {"strk_r.docx", Resources.strk_r},
                {"indexes.json", Resources.indexes}
            };

            if (!Directory.Exists("templates")) Directory.CreateDirectory("templates");

            foreach (var (key, value) in resources)
            {
                var fs = new FileStream("templates/" + key, FileMode.CreateNew);
                var bw = new BinaryWriter(fs);
                bw.Write(value);
                bw.Close();
            }
        }

        _indexes = JsonSerializer.Deserialize<ProtocolIndexes[]>(File.ReadAllText("templates/indexes.json"));

        string fileName;
        if (args.Length == 0 || (File.GetAttributes(args[0]) & FileAttributes.Directory) == 0)
            Exit("Переместите папку с программой на exe файл", 1);

        fileName = args[0] + "\\kmsp.DBF";
        if (!File.Exists(fileName)) Exit("Файл: " + fileName + " не найден.", 1);


        using Stream fos = File.Open(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
        using var reader = new DBFReader(fos);
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        reader.CharEncoding = Encoding.GetEncoding(1251);

        Console.WriteLine(@"Ввод расхода");
        Console.WriteLine(@"0 - без ввода расхода");
        Console.WriteLine(@"1 - в консоли");
        _consumption = ReadLineOfRange(0, 1);

        Console.WriteLine(@"С какого протокола начать?");
        Console.WriteLine(@"0 - за сегодня");
        var startsWith = ReadLineOfRange(0, reader.RecordCount);

        var startTime = DateTime.Now;
        while (true)
        {
            var r = reader.NextRecord();
            if (r == null) break;
            if (startsWith == 0)
            {
                if ((string) r[18] != DateTime.Now.ToShortDateString()) continue;
            }
            else
            {
                if (Convert.ToInt32(Regex.Replace((string) r[5], ".*-07-", "")) < startsWith) continue;
            }

            var dt = Convert.ToDateTime((string) r[18]);

            WriteProtocol(new Protocol
            {
                protocol_number = Regex.Replace((string) r[5], ".*-07-", ""),
                type = (string) r[7],
                factory_number = (string) r[8],
                manufacturer = (string) r[10],
                high_year = (string) r[11],
                user = (string) r[12],
                verification_conditions = (string) r[13],
                date_of_check = $"«{dt.Day}» {Mouth[dt.Month - 1]} {dt.Year} года",
                verifier = (string) r[21],
                label_number = (string) r[27]
            }, dt);
        }

        if (Debugger.IsAttached && _consumption == 0)
            Console.WriteLine(@"Time: " + Convert.ToInt32((DateTime.Now - startTime).TotalMilliseconds));
        Console.ReadKey();
    }

    private static void WriteProtocol(Protocol protocol, DateTime dt)
    {
        var protocolIndex =
            _indexes.FirstOrDefault(indexes => protocol.verification_conditions.StartsWith(indexes.name));
        if (protocolIndex == null)
        {
            Console.WriteLine(
                $@"Тип протокола не опознан. NL-07-{protocol.protocol_number} ");
            Console.WriteLine(protocol.verification_conditions);
            for (var i = 0; i < _indexes.Length; i++) Console.WriteLine($@"{i + 1} - {_indexes[i].name}");
            protocolIndex = _indexes[ReadLineOfRange(1, _indexes.Length) - 1];
        }

        if (!Directory.Exists("output")) Directory.CreateDirectory("output");

        if (!TemplateCache.ContainsKey(protocolIndex.name))
            TemplateCache.Add(protocolIndex.name,
                new MemoryStream(File.ReadAllBytes(_consumption == 0 ? protocolIndex.path : protocolIndex.path_r)));

        var doc = new Document();
        doc.LoadFromStream(TemplateCache[protocolIndex.name], FileFormat.Docx);
        foreach (var propertyInfo in protocol.GetType().GetProperties())
            doc.Replace("{" + propertyInfo.Name + "}", (string) propertyInfo.GetValue(protocol, null)!, false, true);

        if (_consumption == 1)
        {
            Console.WriteLine($@"Номер протокола: NL-07-{protocol.protocol_number} Тип: {protocolIndex.name}");
            Console.WriteLine(@"Введите расход в виде последних цифр процента");
            Console.WriteLine(@"Например расход 1,4 1,5 1,6 ввести - 456");
            var ras = Console.ReadLine() ?? "";
            while (ras.Length != protocolIndex.space.Length)
            {
                Console.WriteLine(@"Неправильно введен расход");
                ras = Console.ReadLine() ?? "";
            }

            var prc = new StringBuilder(ras);

            for (var i = 0; i < ras.Length; i++)
            {
                var rr = (Convert.ToDouble("1," + prc[i]) / 100 * protocolIndex.space[i] + protocolIndex.space[i]).ToString(CultureInfo
                    .CurrentCulture);
                if (rr.Length - protocolIndex.space[i].ToString(CultureInfo.CurrentCulture).Length > 3)
                    rr = rr.Substring(0, protocolIndex.space[i].ToString(CultureInfo.CurrentCulture).Length + 3);
                if (rr.EndsWith("0"))
                    rr = rr.Substring(0, protocolIndex.space[i].ToString(CultureInfo.CurrentCulture).Length + 2);
                doc.Replace("{r" + protocolIndex.space[i] + "}",rr, false, true);
                doc.Replace("{p" + protocolIndex.space[i] + "}", "1," + prc[i], false, true);
            }
        }

        var fileName =
            $"NL-07-{protocol.protocol_number} {dt.ToShortDateString().Replace("/", ".")} {protocolIndex.name}.docx";

        if (File.Exists("output/" + fileName))
        {
            Console.WriteLine(@"Файл с названием " + fileName + @" уже существует");
            Console.WriteLine(@"0 - пропустить");
            Console.WriteLine(@"1 - сохранить");
            if(ReadLineOfRange(0, 1) == 0) return;
        }

        while (true) {
            try
            {
                doc.SaveToFile("output/" + fileName, FileFormat.Docx);
                break;
            }
            catch (IOException e)
            {
                Console.WriteLine(@"Файл открыт в другой программе");
                Console.WriteLine(@"Закройте другие программы и нажмите любую клавишу.");
                Console.ReadKey();
            }
        }
    }

    private static int ReadLineOfRange(int f, int s)
    {
        while (true)
        {
            var line = Console.ReadLine();
            if (line != null && Regex.IsMatch(line, "[0-9]"))
            {
                var i = Convert.ToInt32(line);
                if (f <= i && i <= s) return i;
            }

            Console.WriteLine($@"Введите число от {f} до {s}");
        }
    }

    private static void Exit(string message, int code)
    {
        Console.WriteLine(message);
        Exit(code);
    }

    private static void Exit(int code)
    {
        Console.ReadKey();
        Environment.Exit(code);
    }
}