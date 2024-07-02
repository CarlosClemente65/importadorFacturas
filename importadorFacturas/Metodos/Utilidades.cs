using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace importadorFacturas.Metodos
{
    public class Utilidades
    {
        public string quitaRaros(string cadena)
        {
            //Metodo para eliminar caracteres raros
            List<(string, string)> caracteresReemplazo = new List<(string, string)>
            {
                ("á", "a"),
                ("é", "e"),
                ("í", "i"),
                ("ó", "o"),
                ("ú", "u"),
                ("º", "."),
                ("ª", "."),
                ("ñ", "¤"),
                ("Á", "A"),
                ("É", "E"),
                ("Í", "I"),
                ("Ó", "O"),
                ("Ú", "U"),
                ("Ñ", "¤")
            };

            foreach (var tupla in caracteresReemplazo)
            {
                cadena = cadena.Replace(tupla.Item1, tupla.Item2);
            }
            return cadena;

        }
    }
}
