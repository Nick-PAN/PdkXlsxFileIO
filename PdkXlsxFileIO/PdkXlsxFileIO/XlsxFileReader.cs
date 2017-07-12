using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using System.IO;
using System.Xml;
using System.Globalization;
using System.Text.RegularExpressions;

namespace PdkXlsxFileIO
{
    public class XlsxFileReader
    {
        public delegate void OnProgressPercentChanged(int percent);

        /// <summary>
        /// 读取进度变化1个百分点事件
        /// </summary>
        public event OnProgressPercentChanged ProgressPercentChanged;

        /// <summary>
        /// 字母表
        /// </summary>
        private static string[] _letters = {"A", "B", "C", "D", "E", "F","G", "H", "I", "J", "K", "L","M", "N", "O", "P", "Q", "R","S", "T", "U", "V", "W", "X","Y", "Z" };

        public IList<string> DateTimeDataColumnName { get; set; }

        public XlsxFileReader()
        {
            this.DateTimeDataColumnName = new List<string>();
        }

        /// <summary>
        /// 数字26进制，转换成字母，用递归算法
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private static string Num2Letter(int value)
        {
            //此处判断输入的是否是正确的数字，略（正在表达式判断）
            if (value < 0) throw new Exception("无法将负数转为字母！");
            int remainder = value % 26;
            int front = (value - remainder) / 26;
            if (front > 0)
            {
                if (front < 26)
                {
                    return _letters[front - 1] + _letters[remainder];
                }
                else
                {
                    return Num2Letter(front) + _letters[remainder];
                }
            }
            else
            {
                return _letters[remainder];
            }            
        }

        /// <summary>
        /// 输入字符得到相应的数字，这是最笨的方法，还可用ASIICK编码；
        /// </summary>
        /// <param name="ch"></param>
        /// <returns></returns>
        private static int Char2Num(char ch)
        {
            switch (ch)
            {
                case 'A':
                    return 0;
                case 'B':
                    return 1;
                case 'C':
                    return 2;
                case 'D':
                    return 3;
                case 'E':
                    return 4;
                case 'F':
                    return 5;
                case 'G':
                    return 6;
                case 'H':
                    return 7;
                case 'I':
                    return 8;
                case 'J':
                    return 9;
                case 'K':
                    return 10;
                case 'L':
                    return 11;
                case 'M':
                    return 12;
                case 'N':
                    return 13;
                case 'O':
                    return 14;
                case 'P':
                    return 15;
                case 'Q':
                    return 16;
                case 'R':
                    return 17;
                case 'S':
                    return 18;
                case 'T':
                    return 19;
                case 'U':
                    return 20;
                case 'V':
                    return 21;
                case 'W':
                    return 22;
                case 'X':
                    return 23;
                case 'Y':
                    return 24;
                case 'Z':
                    return 25;
            }
            return -1;
        }

        /// <summary>
        /// 26进制字母转换成数字
        /// </summary>
        /// <param name="letter"></param>
        /// <returns></returns>
        private static int Leter2Num(string str)
        {
            //此处判断是否是由A-Z字母组成的字符串，略（正在表达式片段）
            char[] letter = str.ToCharArray(); //拆分字符串
            int reNum = 0;
            int power = 1; //用于次方算值
            int times = 1;  //最高位需要加1
            int num = letter.Length;//得到字符串个数
            //得到最后一个字母的尾数值
            reNum += Char2Num(letter[num - 1]);
            //得到除最后一个字母的所以值,多于两位才执行这个函数
            if (num >= 2)
            {
                for (int i = num - 1; i > 0; i--)
                {
                    power = 1;//致1，用于下一次循环使用次方计算
                    for (int j = 0; j < i; j++)           //幂，j次方，应该有函数
                    {
                        power *= 26;
                    }
                    reNum += (power * (Char2Num(letter[num - i - 1]) + times));  //最高位需要加1，中间位数不需要加一
                    times = 0;
                }
            }
            //Console.WriteLine(letter.Length);
            return reNum;
        }

        /// <summary>
        /// 比较列代号
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <returns></returns>
        private static int CompareColumnCodes(string x, string y)
        {
            int xIndex = Leter2Num(x);
            int yIndex = Leter2Num(y);
            return xIndex.CompareTo(yIndex);
        }


        private static IDictionary<int, string> ReadSheetNamesFromWorkbookFile(Stream input)
        {
            var sheetNames = new Dictionary<int, string>();
            using (var reader = XmlReader.Create(input))
            {
                for (reader.MoveToContent(); reader.Read(); )
                {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        if (reader.Name == "sheet")
                        {
                            string name = reader.GetAttribute("name");
                            int id = int.Parse(reader.GetAttribute("sheetId"));
                            sheetNames.Add(id, name);
                        }                        
                    }
                }
            }
            return sheetNames;
        }


        /// <summary>
        /// 读取工作表单文件
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="stringTable"></param>
        /// <param name="isFirstRowColumnNames"></param>
        /// <returns></returns>
        private DataTable ReadWorksheetFile(string filename, IList<string> stringTable, bool isFirstRowColumnNames)
        {
            FileInfo fi = new FileInfo(filename);
            if (fi.Extension.ToLower() != ".xml") throw new Exception("文件扩展名只能为.xml！");
            string sheetName = fi.Name.Substring(0, fi.Name.Length - 4);
            DataTable dt = new DataTable(sheetName);
            // Open XML file with worksheet data..
            using (var stream = new FileStream(filename, FileMode.Open, FileAccess.Read))
            // ..and call helper method that parses that XML and fills DataTable with values.
            {             
                using (var reader = XmlReader.Create(stream))
                {
                    DataRow row = null;
                    int columnStartSN = 0;
                    int columnFinalSN = 0;
                    int rowStartSN = 0;
                    int rowFinalSN = 0;
                    int rowSN = 0;
                    int columnIndex = 0;
                    string columnRow;
                    string columnCode;
                    string type = "";
                    int intValue;
                    double doubleValue;
                    string stringValue;
                    int oldPercent = 0;
                    int newPercent = 0;

                    for (reader.MoveToContent(); reader.Read(); )
                    {
                        if (reader.NodeType == XmlNodeType.Element)
                        {
                            switch (reader.Name)
                            {
                                case "dimension":
                                    //读取表的维度，用于构建表结构
                                    string r = reader.GetAttribute("ref");
                                    string[] fromTo = r.Split(':');
                                    if (fromTo.Length == 0) throw new Exception("文件已被损坏！");
                                    if (fromTo.Length > 0)
                                    {
                                        string columnFrom = Regex.Match(fromTo[0], "[a-zA-Z]+").Value;
                                        columnStartSN = Leter2Num(columnFrom);
                                        string rowFrom = Regex.Match(fromTo[0], @"\d+").Value;
                                        rowStartSN = int.Parse(rowFrom);
                                        rowSN = rowStartSN;
                                    }
                                    if (fromTo.Length > 1)
                                    {
                                        string columnTo = Regex.Match(fromTo[1], "[a-zA-Z]+").Value;
                                        columnFinalSN = Leter2Num(columnTo);
                                        string rowTo = Regex.Match(fromTo[1], @"\d+").Value;
                                        rowFinalSN = int.Parse(rowTo);
                                    }
                                    else
                                    {
                                        columnFinalSN = columnStartSN;
                                        rowFinalSN = rowStartSN;
                                    }
 
                                    for (int i = columnStartSN; i <= columnFinalSN; i++)
                                    {
                                        dt.Columns.Add(Num2Letter(i));
                                    }
                                    break;
                                case "row":
                                    row = dt.NewRow();                                    
                                    break;
                                case "c":
                                    columnRow = reader.GetAttribute("r");
                                    columnCode = Regex.Match(columnRow, "[a-zA-Z]+").Value;
                                    columnIndex = Leter2Num(columnCode) - columnStartSN;
                                    type = reader.GetAttribute("t");
                                    break;
                                case "v":
                                    //解析
                                    stringValue = reader.ReadElementString();
                                    if (type == "s")
                                    {
                                        if(int.TryParse(stringValue, out intValue)){
                                            stringValue = stringTable[intValue];
                                        }
                                    }                                        
                                    else if (this.DateTimeDataColumnName!=null && this.DateTimeDataColumnName.Contains(dt.Columns[columnIndex].ColumnName))
                                    {
                                        if (double.TryParse(stringValue, out doubleValue))
                                        {
                                            stringValue = DateTime.FromOADate(doubleValue).ToString();
                                        }                                        
                                    }
                                    //赋值
                                    if (rowSN >= (isFirstRowColumnNames ? rowStartSN + 1 : rowStartSN))
                                    {
                                        row[columnIndex] = stringValue;
                                    }
                                    else
                                    {
                                        dt.Columns[columnIndex].ColumnName = stringValue;
                                    }
                                    break;
                            }
                        }
                        else if (reader.NodeType == XmlNodeType.EndElement)
                        {
                            if (reader.Name == "row")
                            {
                                if (rowSN >= (isFirstRowColumnNames ? rowStartSN + 1 : rowStartSN))
                                {
                                    dt.Rows.Add(row);

                                    //发起进度变化事件
                                    newPercent=XlsxFileHelper.ProgressPercent(rowSN, rowStartSN, rowFinalSN);
                                    if ( newPercent> oldPercent)
                                    {
                                        oldPercent = newPercent;
                                        if (this.ProgressPercentChanged != null)
                                        {
                                            this.ProgressPercentChanged(oldPercent);
                                        }
                                    }
                                }
                                rowSN++;
                            }
                        }
                    }
                }
            }
            return dt;
            
        }
       
        /// <summary>
        /// 读取共享字符串表单
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        private static IList<string> ReadStringTable(Stream input)
        {
            var stringTable = new List<string>();
            using (var reader = XmlReader.Create(input))
            {
                string meta = "";
                bool added = true;
                
                for (reader.MoveToContent(); reader.Read(); )
                {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        if (reader.Name == "si")
                        {
                            if (!added)
                            {
                                stringTable.Add(meta);
                                added = true;
                            }
                            meta = "";
                            added = false;

                        }
                        if (reader.Name == "t")
                        {
                            meta += reader.ReadElementString();
                        }
                    }
                    else if (reader.NodeType == XmlNodeType.EndElement)
                    {
                        if (reader.Name == "si")
                        {
                            if (!added)
                            {
                                stringTable.Add(meta);
                                added = true;
                            }
                        }
                        else if (reader.Name == "sst")
                        {
                            if (!added)
                            {
                                stringTable.Add(meta);
                                added = true;
                            }
                        }
                    }                  
                }
            }
          
            return stringTable;
        }

        /// <summary>
        /// 由Xlsx文件读取数据集
        /// </summary>
        /// <param name="data"></param>
        /// <param name="fileName"></param>
        public DataSet Read(string fileName, bool isFirstRowColumnNames)
        {
            DataSet dataSet = null;
            if (string.IsNullOrEmpty(fileName)) throw new Exception("路径为空！");
            if (!File.Exists(fileName)) throw new Exception("不存在文件" + fileName);
            if (fileName.Length < 5) throw new Exception("无法读取文件" + fileName);
            if (fileName.Substring(fileName.Length - 5, 5).ToLower() != ".xlsx") throw new Exception("无法读取文件" + fileName);
            FileInfo fi = new FileInfo(fileName);
            string fullName = fi.FullName;
            string tempDir = fullName.Substring(0, fullName.Length - 5) + "/";
            // Delete contents of the temporary directory.
            XlsxFileHelper.DeleteDirectoryContents(tempDir);

            // Unzip input XLSX file to the temporary directory.
            XlsxFileHelper.UnzipFile(fileName, tempDir);

            IDictionary<int, string> sheetNames;
            using (var stream = new FileStream(Path.Combine(tempDir, @"xl\workbook.xml"), FileMode.Open, FileAccess.Read))
                sheetNames = ReadSheetNamesFromWorkbookFile(stream);

            IList<string> stringTable;
            // Open XML file with table of all unique strings used in the workbook..
            using (var stream = new FileStream(Path.Combine(tempDir, @"xl\sharedStrings.xml"), FileMode.Open, FileAccess.Read))
                // ..and call helper method that parses that XML and returns an array of strings.
                stringTable = ReadStringTable(stream);

            DirectoryInfo worksheetsDirInfo = new DirectoryInfo(Path.Combine(tempDir, @"xl\worksheets"));
            FileInfo[] sheetsFileInfoArr=worksheetsDirInfo.GetFiles();
            if (sheetsFileInfoArr.Length != 0) dataSet = new DataSet(fi.Name.Substring(0, fi.Name.Length-5));
            if (sheetNames.Count <= sheetsFileInfoArr.Count<FileInfo>())
            {
                for (int i = 1; i <= sheetNames.Count; i++)
                {
                    string sheetFileName = "sheet" + i + ".xml";
                    DataTable dt = ReadWorksheetFile(Path.Combine(worksheetsDirInfo.FullName, sheetFileName), stringTable, isFirstRowColumnNames);
                    dt.TableName = sheetNames[i];
                    dataSet.Tables.Add(dt);
                }
            }
            else
            {
                throw new Exception(string.Format("文件异常，无法读取{0}！", fileName));
            }

            // 删除临时目录
            Directory.Delete(tempDir, true);

            return dataSet;
        }  
    }
}
