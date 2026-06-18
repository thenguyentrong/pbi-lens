using System;
using System.Collections.Generic;
using System.Text.Json;
using Microsoft.AnalysisServices.AdomdClient;

// adomd-runner <port> [database]
// Reads a DAX query from stdin, runs it against a local Analysis Services
// (Power BI Desktop) instance on localhost:<port>, prints {"rows":[...]} to
// stdout. Errors go to stderr with a non-zero exit code.
internal static class Program
{
    private static int Main(string[] args)
    {
        try
        {
            if (args.Length < 1)
            {
                Console.Error.WriteLine("usage: adomd-runner <port> [database]  (DAX query on stdin)");
                return 2;
            }

            string port = args[0];
            string? db = args.Length > 1 ? args[1] : null;
            string dax = Console.In.ReadToEnd();
            if (string.IsNullOrWhiteSpace(dax))
            {
                Console.Error.WriteLine("no DAX query on stdin");
                return 2;
            }

            using var conn = new AdomdConnection($"Data Source=localhost:{port}");
            conn.Open();

            // Desktop instances host exactly one (GUID-named) catalog; pick it
            // unless the caller named one.
            if (string.IsNullOrEmpty(db))
            {
                using var dcmd = new AdomdCommand("SELECT [CATALOG_NAME] FROM $SYSTEM.DBSCHEMA_CATALOGS", conn);
                using var dr = dcmd.ExecuteReader();
                if (dr.Read()) db = dr.GetValue(0)?.ToString();
                dr.Close();
            }
            if (!string.IsNullOrEmpty(db)) conn.ChangeDatabase(db);

            using var cmd = new AdomdCommand(dax, conn);
            using var reader = cmd.ExecuteReader();
            var rows = new List<Dictionary<string, object?>>();
            while (reader.Read())
            {
                var row = new Dictionary<string, object?>();
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    row[reader.GetName(i)] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                }
                rows.Add(row);
            }

            Console.Out.Write(JsonSerializer.Serialize(new { rows }));
            return 0;
        }
        catch (Exception e)
        {
            Console.Error.WriteLine(e.Message);
            return 1;
        }
    }
}
