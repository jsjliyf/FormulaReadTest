using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FormulaReadTest
{
    class InitTables
    {
        public static int InitOriginalColumnCount(string standardReportName)
        {
            int columnNumMax;

            #region 原始表
            switch (standardReportName)
            {
                case "G0100": columnNumMax = 5; break;
                case "G0101": columnNumMax = 5; break;
                case "G0102": columnNumMax = 3; break;
                case "G0103": columnNumMax = 5; break;
                case "G0109": columnNumMax = 5; break;
                case "G0106": columnNumMax = 3; break;
                case "G0107": columnNumMax = 3; break;
                case "G0300": columnNumMax = 10; break;
                case "G0400": columnNumMax = 3; break;
                case "G4A00": columnNumMax = 3; break;
                case "G4A01a": columnNumMax = 3; break;
                case "G4B01": columnNumMax = 22; break;
                case "G4B02": columnNumMax = 9; break;
                case "G4D00": columnNumMax = 6; break;
                case "G0601": columnNumMax = 10; break;
                case "G0602": columnNumMax = 5; break;
                case "G1101": columnNumMax = 10; break;
                case "G1102": columnNumMax = 11; break;
                case "G1103": columnNumMax = 10; break;
                case "G1200": columnNumMax = 16; break;
                case "G1401": columnNumMax = 9; break;
                case "G1402": columnNumMax = 16; break;
                case "G1403": columnNumMax = 17; break;
                case "G1404": columnNumMax = 14; break;
                case "G1405": columnNumMax = 14; break;
                case "G1406": columnNumMax = 14; break;
                case "G14a00": columnNumMax = 23; break;
                case "G1500": columnNumMax = 18; break;
                case "G1700": columnNumMax = 5; break;
                case "G18": columnNumMax = 7; break;
                case "G18_III": columnNumMax = 14; break;
                case "G21": columnNumMax = 13; break;
                case "G2200": columnNumMax = 5; break;
                case "G2400": columnNumMax = 15; break;
                case "G2600": columnNumMax = 5; break;
                case "G3101": columnNumMax = 7; break;
                case "G3102": columnNumMax = 22; break;
                case "G34": columnNumMax = 24; break;
                case "G4000": columnNumMax = 3; break;
                case "G4400": columnNumMax = 3; break;
                case "G5301": columnNumMax = 37; break;
                case "G5302": columnNumMax = 37; break;
                case "G5303": columnNumMax = 37; break;
                case "G5304": columnNumMax = 37; break;
                case "S0300": columnNumMax = 5; break;
                case "S4500": columnNumMax = 4; break;
                case "S4700": columnNumMax = 3; break;
                case "S4800": columnNumMax = 14; break;
                case "S4900": columnNumMax = 3; break;
                case "S6301": columnNumMax = 9; break;
                case "S6302": columnNumMax = 9; break;
                case "S6303": columnNumMax = 8; break;
                case "S6304": columnNumMax = 12; break;
                case "S6401": columnNumMax = 10; break;
                case "S6402": columnNumMax = 10; break;
                case "S6501": columnNumMax = 22; break;
                case "S6502": columnNumMax = 10; break;
                case "S6600": columnNumMax = 18; break;
                case "S6700": columnNumMax = 22; break;
                case "S70_I科技金融基本情况表": columnNumMax = 9; break;
                case "S7002": columnNumMax = 9; break;
                case "S7101": columnNumMax = 33; break;
                case "S7102": columnNumMax = 26; break;
                case "S7103": columnNumMax = 20; break;
                default: columnNumMax = 0; break;
            }
            #endregion

            #region 区域特色表
            if (standardReportName == "QYTS29") //表一
                columnNumMax = 39;
            else if (standardReportName == "QYTS30") //表二
                columnNumMax = 25;
            else if (standardReportName == "QYTS31") //表三
                columnNumMax = 41;
            else if (standardReportName == "QYTS32") //表四
                columnNumMax = 14;
            else if (standardReportName == "QYTS28") //户均贷款
                columnNumMax = 12;
            else if (standardReportName.Contains("S41")) //农合机构S41-农村合作金融机构业务状况表
                columnNumMax = 9;
            #endregion

            return columnNumMax;
        }

        public static int InitAuditSumColumnCount(string standardReportName)
        {
            int columnNumMax;

            #region 审核或汇总表
            switch (standardReportName)
            {
                case "G0100": columnNumMax = 5 + 10; break;
                case "G0101": columnNumMax = 5 + 10; break;
                case "G0102": columnNumMax = 3 + 10; break;
                case "G0103": columnNumMax = 5 + 10; break;
                case "G0109": columnNumMax = 5 + 10; break;
                case "G0106": columnNumMax = 3 + 10; break;
                case "G0107": columnNumMax = 3 + 10; break;
                case "G0300": columnNumMax = 10 + 10; break;
                case "G0400": columnNumMax = 3 + 10; break;
                case "G4A00": columnNumMax = 3 + 10; break;
                case "G4A01a": columnNumMax = 3 + 10; break;
                case "G4B01": columnNumMax = 22 + 10; break;
                case "G4B02": columnNumMax = 9 + 10; break;
                case "G4D00": columnNumMax = 6 + 10; break;
                case "G0601": columnNumMax = 10 + 10; break;
                case "G0602": columnNumMax = 5 + 10; break;
                case "G1101": columnNumMax = 10+10; break;
                case "G1102": columnNumMax = 11+10; break;
                case "G1103": columnNumMax = 10+10; break;
                case "G1200": columnNumMax = 16 + 10; break;
                case "G1401": columnNumMax = 9 + 10; break;
                case "G1402": columnNumMax = 16 + 10; break;
                case "G1403": columnNumMax = 17 + 10; break;
                case "G1404": columnNumMax = 14 + 10; break;
                case "G1405": columnNumMax = 14 + 10; break;
                case "G1406": columnNumMax = 14 + 10; break;
                case "G14a00": columnNumMax = 23 + 10; break;
                case "G1500": columnNumMax = 18 + 10; break;
                case "G1700": columnNumMax = 5 + 10; break;
                case "G18": columnNumMax = 7 + 10; break;
                case "G1803": columnNumMax = 14 + 10; break;
                case "G21": columnNumMax = 13 + 10; break;
                case "G2200": columnNumMax = 5 + 10; break;
                case "G2400": columnNumMax = 15 + 10; break;
                case "G2600": columnNumMax = 5 + 10; break;
                case "G3101": columnNumMax = 7 + 10; break;
                case "G3102": columnNumMax = 22 + 10; break;
                case "G34": columnNumMax = 24 + 10; break;
                case "G4000": columnNumMax = 3 + 10; break;
                case "G4400": columnNumMax = 3 + 10; break;
                case "G5301": columnNumMax = 37 + 10; break;
                case "G5302": columnNumMax = 37 + 10; break;
                case "G5303": columnNumMax = 37 + 10; break;
                case "G5304": columnNumMax = 37 + 10; break;
                case "S0300": columnNumMax = 5 + 10; break;
                case "S4500": columnNumMax = 4 + 10; break;
                case "S4700": columnNumMax = 3 + 10; break;
                case "S4800": columnNumMax = 14 + 10; break;
                case "S4900": columnNumMax = 3 + 10; break;
                case "S6301": columnNumMax = 9 + 10; break;
                case "S6302": columnNumMax = 9 + 10; break;
                case "S6303": columnNumMax = 8 + 10; break;
                case "S6304": columnNumMax = 12 + 10; break;
                case "S6401": columnNumMax = 10 + 10; break;
                case "S6402": columnNumMax = 10 + 10; break;
                case "S6501": columnNumMax = 22 + 20; break;
                case "S6502": columnNumMax = 10 + 10; break;
                case "S6600": columnNumMax = 18 + 10; break;
                case "S6700": columnNumMax = 22 + 10; break;
                case "S7001": columnNumMax = 9 + 10; break;
                case "S7002": columnNumMax = 9 + 10; break;
                case "S7101": columnNumMax = 33 + 10; break;
                case "S7102": columnNumMax = 26 + 10; break;
                case "S7103": columnNumMax = 20 + 10; break;
                default: columnNumMax = 0; break;
            }
            #endregion

            #region 区域特色表
            if (standardReportName == "QYTS29") //表一
                columnNumMax = 39;
            else if (standardReportName == "QYTS30") //表二
                columnNumMax = 25;
            else if (standardReportName == "QYTS31") //表三
                columnNumMax = 41;
            else if (standardReportName == "QYTS32") //表四
                columnNumMax = 14;
            else if (standardReportName == "QYTS28") //户均贷款
                columnNumMax = 12;
            else if (standardReportName.Contains("S41")) //农合机构S41-农村合作金融机构业务状况表
                columnNumMax = 9;
            #endregion

            return columnNumMax;
        }

        public static string GenerateStandardSheetName(string originalReportName)
        {
            string standardReportName;

            switch (originalReportName)
            {
                case "G01": standardReportName = "G0100"; break;
                case "G01_I": standardReportName = "G0101"; break;
                case "G01_II": standardReportName = "G0102"; break;
                case "G01_III": standardReportName = "G0103"; break;
                case "G01_IX": standardReportName = "G0109"; break;
                case "G01_VI": standardReportName = "G0106"; break;
                case "G01_VII": standardReportName = "G0107"; break;
                case "G03": standardReportName = "G0300"; break;
                case "G04": standardReportName = "G0400"; break;
                case "G4A": standardReportName = "G4A00"; break;
                case "G4A-1(a)": standardReportName = "G4A01a"; break;
                case "G4B_I": standardReportName = "G4B01"; break;
                case "G4B_II": standardReportName = "G4B02"; break;
                case "G4D": standardReportName = "G4D00"; break;
                case "G06_I": standardReportName = "G0601"; break;
                case "G06_II": standardReportName = "G0602"; break;
                case "G11_I": standardReportName = "G1101"; break;
                case "G11_II": standardReportName = "G1102"; break;
                case "G11_III": standardReportName = "G1103"; break;
                case "G12": standardReportName = "G1200"; break;
                case "G14_I": standardReportName = "G1401"; break;
                case "G14a": standardReportName = "G14a00"; break;
                case "G15": standardReportName = "G1500"; break;
                case "G17": standardReportName = "G1700"; break;
                case "G18": standardReportName = "G18"; break;
                case "G18_III": standardReportName = "G18_III"; break;
                case "G21": standardReportName = "G21"; break;
                case "G22": standardReportName = "G2200"; break;
                case "G24": standardReportName = "G2400"; break;
                case "G26": standardReportName = "G2600"; break;
                case "G31_I": standardReportName = "G3101"; break;
                case "G31_II": standardReportName = "G3102"; break;
                case "G34": standardReportName = "G34"; break;
                case "G40": standardReportName = "G4000"; break;
                case "G44": standardReportName = "G4400"; break;
                case "G53_I": standardReportName = "G5301"; break;
                case "G53_II": standardReportName = "G5302"; break;
                case "G53_III": standardReportName = "G5303"; break;
                case "G53_IV": standardReportName = "G5304"; break;
                case "S03": standardReportName = "S0300"; break;
                case "S45": standardReportName = "S4500"; break;
                case "S47": standardReportName = "S4700"; break;
                case "S48": standardReportName = "S4800"; break;
                case "S49": standardReportName = "S4900"; break;
                case "S63_I": standardReportName = "S6301"; break;
                case "S63_II": standardReportName = "S6302"; break;
                case "S63_III": standardReportName = "S6303"; break;
                case "S63_IV": standardReportName = "S6304"; break;
                case "S64_I": standardReportName = "S6401"; break;
                case "S64_II": standardReportName = "S6402"; break;
                case "S65_I": standardReportName = "S6501"; break;
                case "S65_II": standardReportName = "S6502"; break;
                case "S66": standardReportName = "S6600"; break;
                case "S67": standardReportName = "S6700"; break;
                case "S70_I": standardReportName = "S7001"; break;
                case "S70_II": standardReportName = "S7002"; break;
                case "S71_I": standardReportName = "S7101"; break;
                case "S71_II": standardReportName = "S7102"; break;
                case "S71_III": standardReportName = "S7103"; break;

                default: standardReportName = "NotFound"; break;
            }

            return standardReportName;
        }
    }
}
