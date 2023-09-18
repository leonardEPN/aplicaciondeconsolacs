/*
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NuestraPrimeraAplicacionDeConsola
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Introduzca el primer string: ");    //con esta linea de comando vamos a mostrar por consola el mensaje introduzca el primer string 
            String primerString = Console.ReadLine();             //el contenido de tipo string que el usuario haya introducido lo debemos almacenar en una variable
            
            Console.WriteLine("Introduzca el segundo string: ");
            String segundoString = Console.ReadLine();

            string resultado = primerString + " " + segundoString; //para concatenar los dos string se ingresa esta linea de comandos
            Console.WriteLine(resultado); //vamos a mostrar las variable resultado por pantalla
            Console.ReadLine(); //se  deja bloqueado la consola hasta que el usuario introduzca una tecla
        }
    }
}
*/
/*
using System; // Importa el espacio de nombres System, que contiene clases básicas y tipos para trabajar con C#.

using System.Collections.Generic; // Importa el espacio de nombres System.Collections.Generic, que contiene tipos de colección genéricos.

using System.Linq; // Importa el espacio de nombres System.Linq, que proporciona extensiones de lenguaje para trabajar con colecciones.

using System.Text; // Importa el espacio de nombres System.Text, que proporciona clases para trabajar con cadenas de caracteres.

using System.Threading.Tasks; // Importa el espacio de nombres System.Threading.Tasks, que proporciona clases para trabajar con tareas asincrónicas.

namespace NuestraPrimeraAplicacionDeConsola // Define un espacio de nombres llamado "NuestraPrimeraAplicacionDeConsola".
{
    internal class Program // Define una clase llamada "Program". La clase es "internal" (accesible solo dentro del ensamblado).
    {
        static void Main(string[] args) // Define el método de entrada principal llamado "Main" que recibe un arreglo de cadenas llamado "args".
        {
            Console.WriteLine("Introduzca el primer string: "); // Muestra por consola el mensaje "Introduzca el primer string".

            String primerString = Console.ReadLine(); // Lee una línea de texto ingresada por el usuario y la almacena en la variable "primerString".

            Console.WriteLine("Introduzca el segundo string: "); // Muestra por consola el mensaje "Introduzca el segundo string".

            String segundoString = Console.ReadLine(); // Lee una línea de texto ingresada por el usuario y la almacena en la variable "segundoString".

            string resultado = primerString + " " + segundoString; // Concatena los dos strings ingresados y los almacena en la variable "resultado".

            Console.WriteLine(resultado); // Muestra el contenido de la variable "resultado" por pantalla.

            Console.ReadLine(); // Deja bloqueada la consola hasta que el usuario presione una tecla.
        }
    }
}
*/

/*
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml; // Asegúrate de importar el espacio de nombres necesario.
using System.IO;


class Program
{
    static void Main()
    {
        // Configurar la propiedad LicenseContext
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Definir la ruta del archivo Excel
        //string filePath = @"C:\Users\leonardo.arellano\source\repos\VISUAL BASIC\INTENTO TUTORIAL 1\NuestraPrimeraAplicacionDeConsola\EXCEL\archivo.xlsx";
        string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "MiAplicacion", "archivo.xlsx");


        // Crear un nuevo archivo Excel
        using (var package = new ExcelPackage())
        {
            // Agregar una hoja de trabajo al libro
            var worksheet = package.Workbook.Worksheets.Add("MiHoja");

            // Escribir datos en celdas
            worksheet.Cells["A1"].Value = "Nombre";
            worksheet.Cells["B1"].Value = "Edad";

            worksheet.Cells["A2"].Value = "Juan";
            worksheet.Cells["B2"].Value = 30;

            worksheet.Cells["A3"].Value = "María";
            worksheet.Cells["B3"].Value = 25;

            // Guardar el libro Excel en la ubicación especificada
            package.SaveAs(new System.IO.FileInfo(filePath));
        }

        // Mostrar un mensaje de éxito
        Console.WriteLine("Archivo de Excel creado exitosamente en: " + filePath);
    }
}


using System;
using System.Windows.Forms;
using System.IO;

class Program
{
    [STAThread] // Añade esta línea
    static void Main()
    {
        // Crear una instancia del cuadro de diálogo SaveFileDialog
        SaveFileDialog saveFileDialog = new SaveFileDialog();

        // Configurar propiedades del cuadro de diálogo
        saveFileDialog.Filter = "Archivos de Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*";
        saveFileDialog.Title = "Guardar archivo de Excel";
        saveFileDialog.DefaultExt = "xlsx";

        // Mostrar el cuadro de diálogo y obtener la ubicación de guardado
        if (saveFileDialog.ShowDialog() == DialogResult.OK)
        {
            // Obtener la ruta de acceso seleccionada por el usuario
            string filePath = saveFileDialog.FileName;

            // Obtener la carpeta de destino del archivo
            string folderPath = Path.GetDirectoryName(filePath);

            // Verificar si la carpeta existe, y si no, crearla
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            // Aquí puedes guardar tu archivo de Excel en la ubicación seleccionada (filePath)
            // Ejemplo:
            // GuardarArchivoDeExcel(filePath);
        }
        else
        {
            // El usuario canceló la operación
            Console.WriteLine("Operación cancelada por el usuario.");
        }
    }

    // Método para guardar el archivo de Excel en la ubicación seleccionada
    // Puedes implementar esta función según tus necesidades específicas
    // private static void GuardarArchivoDeExcel(string filePath)
    // {
    //     // Aquí colocas el código para guardar el archivo de Excel en filePath
    // }
}
*/


using System;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;


class Program
{
    [STAThread]
    static void Main()
    {
        // Crear una instancia del cuadro de diálogo SaveFileDialog
        SaveFileDialog saveFileDialog = new SaveFileDialog();

        // Configurar propiedades del cuadro de diálogo
        saveFileDialog.Filter = "Archivos de Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*";
        saveFileDialog.Title = "Guardar archivo de Excel";
        saveFileDialog.DefaultExt = "xlsx";

        // Mostrar el cuadro de diálogo y obtener la ubicación de guardado
        if (saveFileDialog.ShowDialog() == DialogResult.OK)
        {
            // Configurar la propiedad LicenseContext
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Obtener la ruta de acceso seleccionada por el usuario
            string filePath = saveFileDialog.FileName;

            using (var package = new ExcelPackage())
            {
                // Agregar una hoja de trabajo al libro
                var worksheet = package.Workbook.Worksheets.Add("MiHoja");

                // Escribir datos en celdas
                worksheet.Cells["A1"].Value = "Nombre";
                worksheet.Cells["B1"].Value = "Edad";

                worksheet.Cells["A2"].Value = "Juan";
                worksheet.Cells["B2"].Value = 30;

                worksheet.Cells["A3"].Value = "María";
                worksheet.Cells["B3"].Value = 25;

                // Guardar el libro Excel en la ubicación especificada
                package.SaveAs(new System.IO.FileInfo(filePath));
            }

            // Mostrar la ubicación del archivo en la consola
            Console.WriteLine("El archivo se guardó en: " + filePath);

            // Aquí puedes guardar tu archivo de Excel en la ubicación seleccionada (filePath)
            // Ejemplo:
            // GuardarArchivoDeExcel(filePath);
            Console.ReadLine();
        }
        else
        {
            // El usuario canceló la operación
            Console.WriteLine("Operación cancelada por el usuario.");
            
        }
    }

    // Método para guardar el archivo de Excel en la ubicación seleccionada
    // Puedes implementar esta función según tus necesidades específicas
    // private static void GuardarArchivoDeExcel(string filePath)
    // {
    //     // Aquí colocas el código para guardar el archivo de Excel en filePath
    // }
}
