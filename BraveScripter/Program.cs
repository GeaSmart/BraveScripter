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
        string connectionString =
            "Server=.;Database=db;User Id=sa;Password=*****;trust server certificate=true";
        //Lista de tablas a procesar, cada tabla debe tener una hoja en Datos.xlsx con el mismo nombre conteniendo los datos
        string[] tablasDestino = ["tabla1", "tabla2"];

        Console.WriteLine($"***** Brave Scripter *****");
        Console.WriteLine($"==========================");

        Console.WriteLine($"Directorio actual: {Directory.GetCurrentDirectory()}");
        Console.WriteLine($"Leyendo archivo: {excelPath}");
        Console.WriteLine($"Se generarán scripts para las siguientes tablas: {string.Join(",", tablasDestino)}");

        if (ConfirmarContinuacion())
        {
            if (File.Exists(excelPath))
            {
                try
                {
                    var columnasPorTabla = ObtenerColumnasSql(connectionString, tablasDestino);
                    var scriptsGenerales = new List<string>();

                    using var stream = File.Open(excelPath, FileMode.Open, FileAccess.Read);
                    using var reader = ExcelReaderFactory.CreateReader(stream);
                    var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    });

                    foreach (var tabla in tablasDestino)
                    {
                        var hoja = dataSet.Tables[tabla]; // Buscar hoja con el mismo nombre de la tabla
                        if (hoja != null)
                        {
                            var (columnasValidas, columnaAutoIncremental) = columnasPorTabla[tabla];
                            var datosExcel = LeerExcel(hoja, columnasValidas);
                            var scripts = GenerarInsertScripts(tabla, datosExcel, columnaAutoIncremental);
                            scriptsGenerales.AddRange(scripts);
                        }
                        else
                        {
                            Console.WriteLine($"No se encontró la hoja '{tabla}' en el Excel.");
                        }
                    }

                    File.WriteAllLines(outputSqlPath, scriptsGenerales);
                    Console.WriteLine($"Scripts generados en: {outputSqlPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"No se encontró el archivo en: {excelPath}");
            }
        }
    }

    static Dictionary<string, (List<string> columnas, string? columnaAutoIncremental)> ObtenerColumnasSql(string connectionString, string[] tablas)
    {
        using var connection = new SqlConnection(connectionString);
        connection.Open();

        var resultado = new Dictionary<string, (List<string>, string?)>();

        foreach (var tabla in tablas)
        {
            var columnas = connection.Query<(string ColumnName, int IsIdentity)>(
                @"SELECT COLUMN_NAME AS ColumnName, COLUMNPROPERTY(OBJECT_ID(@tabla), COLUMN_NAME, 'IsIdentity') AS IsIdentity
                  FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @tabla",
                new { tabla }).ToList();

            string? columnaAutoIncremental = columnas.FirstOrDefault(c => c.IsIdentity == 1).ColumnName;
            var columnasValidas = columnas.Where(c => c.IsIdentity == 0).Select(c => c.ColumnName).ToList();

            resultado[tabla] = (columnasValidas, columnaAutoIncremental);
        }

        return resultado;
    }

    static List<Dictionary<string, object>> LeerExcel(System.Data.DataTable hoja, List<string> columnasValidas)
    {
        var resultado = new List<Dictionary<string, object>>();
        var columnasExcel = new List<string>();

        foreach (DataColumn col in hoja.Columns)
        {
            string columnName = col.ColumnName.Trim();
            if (columnasValidas.Contains(columnName))
                columnasExcel.Add(columnName);
        }

        foreach (DataRow row in hoja.Rows)
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
