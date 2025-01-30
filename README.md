# Brave Scripter 🐿️

![.NET 8](https://devblogs.microsoft.com/dotnet/wp-content/uploads/sites/10/2023/03/asp_blog_image.png)

Aplicación de consola para generar consultas de insersión de datos en SQL Server.

## Características principales

- **Flexibilidad**: ✅ Configura tu cadena de conexión en un json, tambien tus archivos y rutas.
- **Resiliencia**: Validaciones en archivos y operaciones.

## **Tecnologías**
[![.NET 8](https://img.shields.io/badge/.NET-8.0-blue)](https://dotnet.microsoft.com/)
[![Dapper](https://img.shields.io/badge/dapper-microOrm-orm)](https://www.learndapper.com/)
[![SQL Server](https://img.shields.io/badge/SQL%20Server-Database-red)](https://www.microsoft.com/en-us/sql-server)

## **Primeros pasos**

1. **Obtener**
   * Clonar el repo
	   ``` bash
	   git clone https://github.com/GeaSmart/BraveScripter.git
	   ```
   
2. **Configurar**
   * Establecer la configuración en el archivo config.json
   * Crear la carpeta definida en 'ExcelPath' y 'OutputSqlPath' 

3. **Validar**
   * Su archivo excel deberá contener los datos para insertar
   * Ese archivo deberá tener tantas hojas de cálculo como tablas configuradas en 'TablasDestino'
   * El nombre de la hoja deberá coincidir con el nombre de su tabla
   * En la primera fila de cada hoja van los nombres de las columnas

4. **Ejecutar**
	* Ejecute el archivo BraveScripter.exe
   

## **Acerca de**
Proyecto realizado por Gerson Azabache bajo MIT License.
bravedeveloper.com
