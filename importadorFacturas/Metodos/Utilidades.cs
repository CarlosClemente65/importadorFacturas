using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace importadorFacturas.Metodos
{
    public class Utilidades
    {
        public string quitaRaros(string cadena)
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

        public void ControlFicheros(string fichero)
        {
            if(File.Exists(fichero)) File.Delete(fichero);
        }

        public void GrabarFichero(string fichero, string texto)
        {
            File.WriteAllText(fichero, texto, Encoding.Default);
        }
    }
}
