using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMapWord
{
    internal class Program
    {
        static void Main(string[] args)
        {
			try
			{
				if(CmdInputArgument.CheckInputArgment(args, out CmdInputArgument inputArgument))
				{
					ExcelHelper excel = new ExcelHelper(inputArgument);
					excel.StartMapToWord();
				}
			}
			catch (Exception ex)
			{
                Console.WriteLine(ex.Message);
            }
        }
    }
}
