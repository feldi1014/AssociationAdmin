using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOrRemoveGroupMember
{
    internal class CommandArgs
    {
        public CommandArgs(string[] args) 
        {
            // 2025-01-06_MitgliederErwachsene.xlsx -MG:Moos -By:SplitBestellungen.xlsx -ByTable:"Erwachsene Moos" -ByField:Wunschkarte -ByContent:Moos
            MemberFileName = args[0];
            foreach (var arg in args)
            {
                if (string.Compare("--membergroup", arg, true) == 0)
                {
                    Command = ToolCommandType.MemberGroup;
                }
                else if (string.Compare("--fees", arg, true) == 0)
                {
                    Command = ToolCommandType.Fees;
                }
                else if (arg.StartsWith("-MG:", StringComparison.CurrentCultureIgnoreCase))
                {
                    MemberGroupName = arg.Substring(4);
                }
                else if (arg.StartsWith("-By:", StringComparison.CurrentCultureIgnoreCase))
                {
                    ByExcelFileName = arg.Substring(4);
                }
                else if (arg.StartsWith("-ByTable:", StringComparison.CurrentCultureIgnoreCase))
                {
                    ByTableName = arg.Substring(9);
                }
                else if (arg.StartsWith("-ByField:", StringComparison.CurrentCultureIgnoreCase))
                {
                    ByFieldName = arg.Substring(9);
                }
                else if (arg.StartsWith("-ByContent:", StringComparison.CurrentCultureIgnoreCase))
                {
                    ByContent = arg.Substring(11);
                }
                else if (arg.StartsWith("-F:", StringComparison.CurrentCultureIgnoreCase))
                {
                    Fees = arg.Substring(3);
                }
            }
        }

        public ToolCommandType Command { get; } = ToolCommandType.MemberGroup;

        public string MemberFileName { get; } = string.Empty;
        public string MemberGroupName { get; } = string.Empty;
        public string ByExcelFileName { get; } = string.Empty;
        public string ByTableName { get; } = string.Empty;
        public string ByFieldName { get; } = string.Empty;
        public string ByContent { get; } = string.Empty;
        public string Fees { get; private set; } = string.Empty;
    }

    public enum ToolCommandType 
    {
        MemberGroup,
        Fees,
    }
}
