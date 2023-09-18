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
using OfficeOpenXml; 
using System.IO;


class Program
{
    static void Main()
    {
        
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


        string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "MiAplicacion", "archivo.xlsx");

        using (var package = new ExcelPackage())
        {
      
            var worksheet = package.Workbook.Worksheets.Add("MiHoja");

   
            worksheet.Cells["A1"].Value = "Nombre";
            worksheet.Cells["B1"].Value = "Edad";

            worksheet.Cells["A2"].Value = "Juan";
            worksheet.Cells["B2"].Value = 30;

            worksheet.Cells["A3"].Value = "María";
            worksheet.Cells["B3"].Value = 25;

     
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
    [STAThread] 
    static void Main()
    {
        SaveFileDialog saveFileDialog = new SaveFileDialog();


        saveFileDialog.Filter = "Archivos de Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*";
        saveFileDialog.Title = "Guardar archivo de Excel";
        saveFileDialog.DefaultExt = "xlsx";

      
        if (saveFileDialog.ShowDialog() == DialogResult.OK)
        {
            string filePath = saveFileDialog.FileName;

            string folderPath = Path.GetDirectoryName(filePath);

        
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

        }
        else
        {

            Console.WriteLine("Operación cancelada por el usuario.");
        }
    }


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
   
        SaveFileDialog saveFileDialog = new SaveFileDialog();

        saveFileDialog.Filter = "Archivos de Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*";
        saveFileDialog.Title = "Guardar archivo de Excel";
        saveFileDialog.DefaultExt = "xlsx";

        if (saveFileDialog.ShowDialog() == DialogResult.OK)
        {
 
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string filePath = saveFileDialog.FileName;

            using (var package = new ExcelPackage())
            {
 
                var worksheet = package.Workbook.Worksheets.Add("MiHoja");

                worksheet.Cells["A1"].Value = "Nombre";
                worksheet.Cells["B1"].Value = "Edad";

                worksheet.Cells["A2"].Value = "Juan";
                worksheet.Cells["B2"].Value = 30;

                worksheet.Cells["A3"].Value = "María";
                worksheet.Cells["B3"].Value = 25;

                package.SaveAs(new System.IO.FileInfo(filePath));
            }

    
            Console.WriteLine("El archivo se guardó en: " + filePath);

    
            Console.ReadLine();
        }
        else
        {

            Console.WriteLine("Operación cancelada por el usuario.");
            
        }
    }

}
