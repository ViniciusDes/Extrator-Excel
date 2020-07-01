using System;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;


namespace ConsoleLeitorArquivos
{
    class Program
    {
        static void Main(string[] args)
        {
            var Ext = new ExtratorEmp27();
            Ext.LerArquivoUnimed();


        }
    
    }
}
