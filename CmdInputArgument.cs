using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMapWord
{
    internal class CmdInputArgument
    {
        public string ExcelPath { get; set; }

        public string WordSamplePath { get; set; }

        public string SheetName { get; set; }

        public string Row { get; set; }

        #region + public static bool CheckInputArgment(string[] args, out CmdInputArgument cmdInputArgument)
        public static bool CheckInputArgment(string[] args, out CmdInputArgument cmdInputArgument)
        {
            cmdInputArgument = null;

            if (args.Length != 8)
            {
                //error infomation
                StringBuilder sb = new StringBuilder();
                sb.AppendLine();
                sb.AppendLine("-excel excel文件路径,路径有空格要用英文引号括起来");
                sb.AppendLine("-sheet excel文件的sheet名");
                sb.AppendLine("-row sheet行范围, 例如: -row 2 或 -row 2-10");
                sb.AppendLine("-word word范本文件路径,路径有空格要用英文引号括起来");
                sb.AppendLine();
                sb.AppendLine("-------------");
                sb.AppendLine("Word文件范本参数格式: [?] , ? 表示excel文件 列符号, 例如: [A] , [B]");
                sb.AppendLine();

                Console.WriteLine(sb.ToString());
                return false;
            }

            cmdInputArgument = new CmdInputArgument();

            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].Trim().ToLower() == "-excel")
                {
                    cmdInputArgument.ExcelPath = args[i+1];
                }

                if (args[i].Trim().ToLower() == "-sheet")
                {
                    cmdInputArgument.SheetName = args[i + 1];
                }

                if (args[i].Trim().ToLower() == "-row")
                {
                    cmdInputArgument.SheetName = args[i + 1];
                }

                if (args[i].Trim().ToLower() == "-word")
                {
                    cmdInputArgument.WordSamplePath = args[i + 1];
                }
            }

            return true;
        }
        #endregion
    }
}
