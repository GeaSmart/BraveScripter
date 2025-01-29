using Microsoft.Data.SqlClient;
using Dapper;
using ExcelDataReader;
using System.Data;


class Program
{
    static void Main()
    {
        string excelPath = Path.Combine(Directory.GetCurrentDirectory(), "Excel", "Datos.xlsx");
        string outputSqlPath = Path.Combine(Directory.GetCurrentDirectory(), "Excel", "InsertScripts.sql");
        string connectionString = "Server=.;Database=QuickTest;User Id=sa;Password=EC1admin;trust server certificate=true";
        string tablaDestino = "SAP_INSP_AREA";

        Console.WriteLine($"📂 Directorio actual: {Directory.GetCurrentDirectory()}");
        Console.WriteLine($"📂 Leyendo archivo: {excelPath}");

        if (ConfirmarContinuacion())
        {
            if (File.Exists(excelPath))
            {
                try
                {
                    var columnasSql = ObtenerColumnasSql(connectionString, tablaDestino, out string? columnaAutoIncremental);
                    var datosExcel = LeerExcel(excelPath, columnasSql);
                    var scripts = GenerarInsertScripts(tablaDestino, datosExcel, columnaAutoIncremental);

                    File.WriteAllLines(outputSqlPath, scripts);
                    Console.WriteLine($"✅ Scripts generados en: {outputSqlPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"⛔ Error: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"⛔ No se encontró el archivo en: {excelPath}");
            }
        }
    }

    static List<string> ObtenerColumnasSql(string connectionString, string tabla, out string? columnaAutoIncremental)
    {
        using var connection = new SqlConnection(connectionString);
        connection.Open();

        var columnas = connection.Query<(string ColumnName, int IsIdentity)>(
            @"SELECT COLUMN_NAME AS ColumnName, COLUMNPROPERTY(OBJECT_ID(@tabla), COLUMN_NAME, 'IsIdentity') AS IsIdentity
              FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @tabla",
            new { tabla }).ToList();

        // Determinar si hay una columna autoincremental
        columnaAutoIncremental = columnas.FirstOrDefault(c => c.IsIdentity == 1).ColumnName;

        // Devolver solo las columnas que no son autoincrementales
        return columnas.Where(c => c.IsIdentity == 0).Select(c => c.ColumnName).ToList();
    }

    static List<Dictionary<string, object>> LeerExcel(string filePath, List<string> columnasValidas)
    {
        var resultado = new List<Dictionary<string, object>>();
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
        {
            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
        });

        var table = dataSet.Tables[0];
        var columnasExcel = new List<string>();

        // Filtrar las columnas que existen en la tabla destino
        foreach (DataColumn col in table.Columns)
        {
            string columnName = col.ColumnName.Trim();
            if (columnasValidas.Contains(columnName))
                columnasExcel.Add(columnName);
        }

        // Leer filas de datos
        foreach (DataRow row in table.Rows)
        {
            var fila = new Dictionary<string, object>();
            foreach (var columna in columnasExcel)
            {
                fila[columna] = row[columna] != DBNull.Value ? row[columna] : null;
            }
            resultado.Add(fila);
        }

        return resultado;
    }

    static List<string> GenerarInsertScripts(string tabla, List<Dictionary<string, object>> datos, string? columnaAutoIncremental)
    {
        var scripts = new List<string>();

        foreach (var fila in datos)
        {
            // Excluir la columna autoincremental si existe
            var columnas = fila.Keys.Where(k => k != columnaAutoIncremental);
            var valores = columnas.Select(col => fila[col] is string ? $"'{fila[col]}'" : fila[col]?.ToString() ?? "NULL");

            scripts.Add($"INSERT INTO {tabla} ({string.Join(", ", columnas)}) VALUES ({string.Join(", ", valores)});");
        }

        return scripts;
    }

    public static bool ConfirmarContinuacion(string mensaje = "¿Desea continuar? (y/n): ")
    {
        Console.Write(mensaje);

        while (true)
        {
            var tecla = Console.ReadKey(true).Key;
            Console.WriteLine();

            if (tecla == ConsoleKey.Y) return true;
            if (tecla == ConsoleKey.N) return false;

            Console.Write("Entrada no válida. Por favor, presione Y o N: ");
        }
    }
}
