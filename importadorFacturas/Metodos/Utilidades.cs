using System.Collections.Generic;
using System.IO;
using System.Text;

namespace importadorFacturas.Metodos
{
    public class Utilidades
    {
        //Metodo para quitar simbolos raros
        public string QuitaRaros(string cadena)
        {
            Dictionary<char, char> caracteresReemplazo = new Dictionary<char, char>
            {
                {'á', 'a'}, {'é', 'e'}, {'í', 'i'}, {'ó', 'o'}, {'ú', 'u'},
                {'Á', 'A'}, {'É', 'E'}, {'Í', 'I'}, {'Ó', 'O'}, {'Ú', 'U'}
                //{'\u00AA', '.'}, {'ª', '.'}, {'\u00BA', '.'}, {'°', '.' }
            };

            StringBuilder resultado = new StringBuilder(cadena.Length);
            foreach(char c in cadena)
            {
                if(caracteresReemplazo.TryGetValue(c, out char reemplazo))
                {
                    resultado.Append(reemplazo);
                }
                else
                {
                    resultado.Append(c);
                }
            }

            return resultado.ToString();
        }

        //Controla si existe el fichero para borrarlo
        public void ControlFicheros(string fichero)
        {
            if(File.Exists(fichero)) File.Delete(fichero);
        }

        //Metodo para grabar el fichero en la ruta que se pase
        public void GrabarFichero(string fichero, string texto)
        {
            File.WriteAllText(fichero, texto, Encoding.Default);
        }

        //Metodo para dividir una cadena por el divisor pasado y solo la divide en un maximo de 2 partes (divide desde el primer divisor que encuentra)
        public (string, string) DivideCadena(string cadena, char divisor)
        {
            string atributo = string.Empty;
            string valor = string.Empty;
            string[] partes = cadena.Split(new[] { divisor }, 2);
            if(partes.Length == 2)
            {
                atributo = partes[0].Trim();
                valor = partes[1].Trim();
            }

            return (atributo, valor);
        }
    }
}
