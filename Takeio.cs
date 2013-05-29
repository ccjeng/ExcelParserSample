using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using TakeIo.Spreadsheet;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ExcelParser
{
    class Takeio
    {
        static void Main(string[] args)
        {

            DateTime startTime = DateTime.Now;

            var inputFile = new FileInfo("..\\..\\File\\Book1500.xlsx"); // could be .xls or .xlsx too
            var sheet = Spreadsheet.Read(inputFile);

            sheet.RemoveEmptyRows();

            int start = 10; //start from row 10
            bool beginFind = false;

            List<Row> rowObj = new List<Row>();

            

            for (var i = 0; i < sheet.Count; i++)
            { 
                if (i == start)  //start parsing
                {
                    beginFind = true;
                } else if (sheet[i][0]=="Total") //end parsing
                {
                    break;
                }
                if (beginFind)
                {
                    rowObj.Add(new Row {
                            Header1  = sheet[i][0],
                            Header2  = sheet[i][1],
                            Header3  = sheet[i][2],
                            Header4  = sheet[i][3],
                            Header5  = sheet[i][4],
                            Header6  = sheet[i][5],
                            Header7  = sheet[i][6],
                            Header8  = sheet[i][7],
                            Header9  = sheet[i][8],
                            Header10 = sheet[i][9],
                            Header11 = sheet[i][10],
                            Header12 = sheet[i][11],
                            Header13 = sheet[i][12],
                            Header14 = sheet[i][13],
                            Header15 = sheet[i][14],
                            Header16 = sheet[i][15],
                            Header17 = sheet[i][16],
                            Header18 = sheet[i][17],
                            Header19 = sheet[i][18],
                            Header20 = sheet[i][19],
                            Header21 = sheet[i][20],
                            Header22 = sheet[i][21],
                            Header23 = sheet[i][22],
                            Header24 = sheet[i][23],
                            Header25 = sheet[i][24],
                            Header26 = sheet[i][25],
                            Header27 = sheet[i][26],
                            Header28 = sheet[i][27],
                            Header29 = sheet[i][28],
                            Header30 = sheet[i][29],
                            Header31 = sheet[i][30],
                            Header32 = sheet[i][31],
                            Header33 = sheet[i][32],
                            Header34 = sheet[i][33],
                            Header35 = sheet[i][34],
                            Header36 = sheet[i][35],
                            Header37 = sheet[i][36],
                            Header38 = sheet[i][37],
                            Header39 = sheet[i][38],
                            Header40 = sheet[i][39],
                            Header41 = sheet[i][40],
                            Header42 = sheet[i][41],
                            Header43 = sheet[i][42],
                            Header44 = sheet[i][43],
                            Header45 = sheet[i][44],
                            Header46 = sheet[i][45],
                            Header47 = sheet[i][46],
                            Header48 = sheet[i][47],
                            Header49 = sheet[i][48],
                            Header50 = sheet[i][49],
                            Header51 = sheet[i][50],
                            Header52 = sheet[i][51],
                            Header53 = sheet[i][52],
                            Header54 = sheet[i][53],
                            Header55 = sheet[i][54],
                            Header56 = sheet[i][55],
                            Header57 = sheet[i][56],
                            Header58 = sheet[i][57],
                            Header59 = sheet[i][58],
                            Header60 = sheet[i][59],
                            Header61 = sheet[i][60],
                            Header62 = sheet[i][61],
                            Header63 = sheet[i][62],
                            Header64 = sheet[i][63],
                            Header65 = sheet[i][64],
                            Header66 = sheet[i][65],
                            Header67 = sheet[i][66],
                            Header68 = sheet[i][67],
                            Header69 = sheet[i][68],
                            Header70 = sheet[i][69],
                            Header71 = sheet[i][70],
                            Header72 = sheet[i][71],
                            Header73 = sheet[i][72],
                            Header74 = sheet[i][73],
                            Header75 = sheet[i][74],
                            Header76 = sheet[i][75],
                            Header77 = sheet[i][76],
                            Header78 = sheet[i][77],
                            Header79 = sheet[i][78],
                            Header80 = sheet[i][79],
                            Header81 = sheet[i][80],
                            Header82 = sheet[i][81],
                            Header83 = sheet[i][82],
                            Header84 = sheet[i][83],
                            Header85 = sheet[i][84],
                            Header86 = sheet[i][85],
                            Header87 = sheet[i][86],
                            Header88 = sheet[i][87],
                            Header89 = sheet[i][88],
                            Header90 = sheet[i][89],
                            Header91 = sheet[i][90],
                            Header92 = sheet[i][91],
                            Header93 = sheet[i][92],
                            Header94 = sheet[i][93],
                            Header95 = sheet[i][94],
                            Header96 = sheet[i][95],
                            Header97 = sheet[i][96],
                            Header98 = sheet[i][97],
                            Header99 = sheet[i][98],
                            Header100= sheet[i][99]
                        }
                    );
                 
                }
            
            }

            DateTime endParseTime = DateTime.Now;

            double Parse_elapsedTime = (endParseTime - startTime).TotalMilliseconds;


            Console.Write(JsonConvert.SerializeObject(rowObj, Formatting.Indented));

            DateTime endTime = DateTime.Now;

            double elapsedTime = (endTime - startTime).TotalMilliseconds;

            Console.Write("Parse ElapsedTime:" + Parse_elapsedTime + " ms");

            Console.Write("Total ElapsedTime:" + elapsedTime + " ms");

            Console.Read();

        }
    }
}
