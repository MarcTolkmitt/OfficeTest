using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Dao = Microsoft.Office.Interop.Access.Dao;

namespace OfficeTest
{
    /// <summary>
    /// stores simple name/value parameters in an Access database (DAO / ACE engine)
    /// table: Settings( Name TEXT(64) PRIMARY KEY, Wert TEXT(255) )
    /// </summary>
    public static class SettingsDb
    {
        private const string TableName = "Settings";
        private const string IndexName = "PK_Settings";
        private const string LocaleGeneral = ";LANGID=0x0409;CP=1252;COUNTRY=0";

        /// <summary>
        /// full path of the database file, %LOCALAPPDATA%\OfficeTest\OfficeTest.accdb
        /// </summary>
        public static string DbPath { get; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OfficeTest", "OfficeTest.accdb");

        /// <summary>
        /// reads all parameters, empty dictionary if the database does not exist yet
        /// </summary>
        public static Dictionary<string, string> Load()
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (!File.Exists(DbPath))
                return result;

            WithDatabase(db =>
            {
                Dao.Recordset rs = db.OpenRecordset(
                    $"SELECT Name, Wert FROM {TableName}",
                    Dao.RecordsetTypeEnum.dbOpenSnapshot);
                try
                {
                    while (!rs.EOF)
                    {
                        string name = Convert.ToString(rs.Fields["Name"].Value) ?? "";
                        string value = Convert.ToString(rs.Fields["Wert"].Value) ?? "";
                        result[name] = value;
                        rs.MoveNext();
                    }
                }
                finally
                {
                    rs.Close();
                    Marshal.ReleaseComObject(rs);
                }
            });
            return result;

        }   // end: Load

        /// <summary>
        /// inserts or updates the given parameters (database is created if needed)
        /// </summary>
        public static void Save(IDictionary<string, string> values)
        {
            WithDatabase(db =>
            {
                Dao.Recordset rs = db.OpenRecordset(TableName, Dao.RecordsetTypeEnum.dbOpenTable);
                try
                {
                    rs.Index = IndexName;
                    foreach (var pair in values)
                    {
                        rs.Seek("=", pair.Key);
                        if (rs.NoMatch)
                        {
                            rs.AddNew();
                            rs.Fields["Name"].Value = pair.Key;
                        }
                        else
                        {
                            rs.Edit();
                        }
                        rs.Fields["Wert"].Value = pair.Value;
                        rs.Update();
                    }
                }
                finally
                {
                    rs.Close();
                    Marshal.ReleaseComObject(rs);
                }
            });

        }   // end: Save

        // ---------------------------------------------     helper functions

        public static string ToText(double value) => value.ToString("R", CultureInfo.InvariantCulture);

        public static bool TryGetDouble(Dictionary<string, string> values, string name, out double result)
        {
            result = 0;
            return values.TryGetValue(name, out string? text)
                && double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out result);
        }

        /// <summary>
        /// opens (or creates) the database, makes sure the table exists, runs the action
        /// </summary>
        private static void WithDatabase(Action<Dao.Database> action)
        {
            Type engineType = Type.GetTypeFromProgID("DAO.DBEngine.120")
                ?? throw new InvalidOperationException(
                    "Access-Datenbankmodul (DAO.DBEngine.120) ist nicht installiert.");
            var engine = (Dao.DBEngine)Activator.CreateInstance(engineType)!;
            Dao.Database? db = null;
            try
            {
                if (File.Exists(DbPath))
                {
                    db = engine.OpenDatabase(DbPath);
                }
                else
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(DbPath)!);
                    db = engine.CreateDatabase(DbPath, LocaleGeneral, Dao.DatabaseTypeEnum.dbVersion120);
                }

                if (!TableExists(db))
                    db.Execute($"CREATE TABLE {TableName} " +
                               $"(Name TEXT(64) CONSTRAINT {IndexName} PRIMARY KEY, Wert TEXT(255))");

                action(db);
            }
            finally
            {
                if (db != null)
                {
                    db.Close();
                    Marshal.ReleaseComObject(db);
                }
                Marshal.ReleaseComObject(engine);
            }

        }   // end: WithDatabase

        private static bool TableExists(Dao.Database db)
        {
            foreach (Dao.TableDef td in db.TableDefs)
            {
                if (string.Equals(td.Name, TableName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

    }   // end: class SettingsDb

}   // end: namespace OfficeTest
