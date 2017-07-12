using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using System.IO;
using System.Xml;
using System.Globalization;
using System.Reflection;

namespace PdkXlsxFileIO
{
    public class XlsxFileWriter
    {
        public delegate void OnProgressPercentChanged(int percent);

        /// <summary>
        /// 读取进度变化1个百分点事件
        /// </summary>
        public event OnProgressPercentChanged ProgressPercentChanged;

        private static IList<string> CreateStringTable(DataSet dataSet, out IDictionary<string, int> lookupTable)
        {
            var stringTable = new List<string>();
            lookupTable = new Dictionary<string, int>();

            foreach (DataTable data in dataSet.Tables)
            {
                //把列名添加到字符串表
                foreach (DataColumn column in data.Columns)
                {
                    object obj = column.ColumnName;
                    if (obj != null)
                    {
                        var value = obj.ToString();
                        if (!lookupTable.ContainsKey(value))
                        {
                            lookupTable.Add(value, stringTable.Count);
                            stringTable.Add(value);
                        }
                    }
                }
                //把数据添加到字符串表
                foreach (DataRow row in data.Rows)
                    foreach (DataColumn column in data.Columns)
                        if (column.DataType == typeof(string))
                        {
                            object obj = row[column];
                            if (obj != null)
                            {
                                var value = obj.ToString();
                                if (!lookupTable.ContainsKey(value))
                                {
                                    lookupTable.Add(value, stringTable.Count);
                                    stringTable.Add(value);
                                }
                            }
                        }
            }            

            return stringTable;
        }

        private static IList<string> CreateStringTable(DataTable dataTable, out IDictionary<string, int> lookupTable)
        {
            var stringTable = new List<string>();
            lookupTable = new Dictionary<string, int>();

            //把列名添加到字符串表
            foreach (DataColumn column in dataTable.Columns)
            {
                object obj = column.ColumnName;
                if (obj != null)
                {
                    var value = obj.ToString();
                    if (!lookupTable.ContainsKey(value))
                    {
                        lookupTable.Add(value, stringTable.Count);
                        stringTable.Add(value);
                    }
                }
            }
            //把数据添加到字符串表
            foreach (DataRow row in dataTable.Rows)
                foreach (DataColumn column in dataTable.Columns)
                    if (column.DataType == typeof(string))
                    {
                        object obj = row[column];
                        if (obj != null)
                        {
                            var value = obj.ToString();
                            if (!lookupTable.ContainsKey(value))
                            {
                                lookupTable.Add(value, stringTable.Count);
                                stringTable.Add(value);
                            }
                        }
                    }

            return stringTable;
        }

        private static void WriteStringTable(Stream output, IList<string> stringTable)
        {
            using (var writer = XmlWriter.Create(output))
            {
                writer.WriteStartDocument(true);

                writer.WriteStartElement("sst", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                writer.WriteAttributeString("count", stringTable.Count.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("uniqueCount", stringTable.Count.ToString(CultureInfo.InvariantCulture));

                foreach (var str in stringTable)
                {
                    writer.WriteStartElement("si");
                    writer.WriteElementString("t", str);
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
            }
        }

        private static void WriteSheetNamesToWorkbook(string workbookFilePath, IList<string> sheetNames)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(workbookFilePath);
            XmlNamespaceManager nsManager = new XmlNamespaceManager(doc.NameTable);
            nsManager.AddNamespace("default", doc.DocumentElement.NamespaceURI);
            XmlNode sheetsNode=doc.SelectSingleNode("//default:sheets", nsManager);
            XmlNodeList sheetNodesList = doc.SelectNodes("//default:sheets/default:sheet", nsManager);
            for (int i = 0; i != sheetNames.Count;i++)
            {
                bool found = false;
                foreach (XmlNode n in sheetNodesList)
                {
                    int id = int.Parse(n.Attributes["sheetId"].Value);
                    if (i + 1 == id)
                    {
                        n.Attributes["name"].Value = sheetNames[i];
                        found = true;
                    }
                }
                if (!found)
                {
                    //添加一个sheet节点
                    XmlNode sheetNode = doc.CreateNode("element", "sheet", "");
                    XmlElement sheetEle = sheetNode as XmlElement;
                    sheetEle.SetAttribute("name", sheetNames[i]);
                    sheetEle.SetAttribute("sheetId", i.ToString());
                    sheetEle.SetAttribute("r:id", "rId" + i.ToString());
                    sheetsNode.AppendChild(sheetsNode);
                }
            }


            doc.Save(workbookFilePath);            
        }

        private void WriteWorksheet(Stream output, DataTable data, IDictionary<string, int> lookupTable, bool isFirstRowColumnNames)
        {
            using (XmlTextWriter writer = new XmlTextWriter(output, Encoding.UTF8))
            {
                writer.WriteStartDocument(true);

                writer.WriteStartElement("worksheet");
                writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                writer.WriteAttributeString("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                writer.WriteStartElement("dimension");
                var lastCell = RowColumnToPosition(data.Rows.Count - 1, data.Columns.Count - 1);
                writer.WriteAttributeString("ref", "A1:" + lastCell);
                writer.WriteEndElement();

                writer.WriteStartElement("sheetViews");
                writer.WriteStartElement("sheetView");
                writer.WriteAttributeString("tabSelected", "1");
                writer.WriteAttributeString("workbookViewId", "0");
                writer.WriteEndElement();
                writer.WriteEndElement();

                writer.WriteStartElement("sheetFormatPr");
                writer.WriteAttributeString("defaultRowHeight", "15");
                writer.WriteEndElement();

                writer.WriteStartElement("sheetData");
                WriteWorksheetData(writer, data, lookupTable, isFirstRowColumnNames);
                writer.WriteEndElement();

                writer.WriteStartElement("pageMargins");
                writer.WriteAttributeString("left", "0.7");
                writer.WriteAttributeString("right", "0.7");
                writer.WriteAttributeString("top", "0.75");
                writer.WriteAttributeString("bottom", "0.75");
                writer.WriteAttributeString("header", "0.3");
                writer.WriteAttributeString("footer", "0.3");
                writer.WriteEndElement();

                writer.WriteEndElement();
            }
        }

        private static string ColumnIndexToName(int columnIndex)
        {
            var second = (char)(((int)'A') + columnIndex % 26);

            columnIndex /= 26;

            if (columnIndex == 0)
                return second.ToString();
            else
                return ((char)(((int)'A') - 1 + columnIndex)).ToString() + second.ToString();
        }

        private static string RowIndexToName(int rowIndex)
        {
            return (rowIndex + 1).ToString(CultureInfo.InvariantCulture);
        }

        private static string RowColumnToPosition(int row, int column)
        {
            return ColumnIndexToName(column) + RowIndexToName(row);
        }

        private void WriteWorksheetData(XmlTextWriter writer, DataTable data, IDictionary<string, int> lookupTable, bool isFirstRowColumnNames)
        {
            var rowsCount = data.Rows.Count;
            var columnsCount = data.Columns.Count;
            string relPos;
            int oldPercent = 0;
            int newPercent = 0;

            for (int row = 0; row < rowsCount; row++)
            {
                writer.WriteStartElement("row");
                relPos = RowIndexToName(row);
                writer.WriteAttributeString("r", relPos);
                writer.WriteAttributeString("spans", "1:" + columnsCount.ToString(CultureInfo.InvariantCulture));

                for (int column = 0; column < columnsCount; column++)
                {
                    object value = data.Rows[row][column];
                    if (isFirstRowColumnNames && row==0)
                    {
                        value = data.Columns[column].ColumnName;
                    }

                    writer.WriteStartElement("c");
                    relPos = RowColumnToPosition(row, column);
                    writer.WriteAttributeString("r", relPos);

                    var str = value as string;
                    if (str != null)
                    {
                        writer.WriteAttributeString("t", "s");
                        value = lookupTable[str];
                    }

                    writer.WriteElementString("v", value.ToString());

                    writer.WriteEndElement();
                }

                writer.WriteEndElement();

                //发起进度变化事件
                newPercent = XlsxFileHelper.ProgressPercent(row+1, 0, rowsCount-1);
                if (newPercent > oldPercent)
                {
                    oldPercent = newPercent;
                    if (this.ProgressPercentChanged != null)
                    {
                        this.ProgressPercentChanged(oldPercent);
                    }
                }
            }
        }

        public void Write(string fileName, DataTable data, bool isFirstRowColumnNames)
        {
            string rootDir = Path.GetDirectoryName(fileName);
            string tempDir = rootDir + "/" + Path.GetFileNameWithoutExtension(fileName);

            if (!Directory.Exists(tempDir))
            {
                Directory.CreateDirectory(tempDir);
            }

            // Delete contents of the temporary directory.
            XlsxFileHelper.DeleteDirectoryContents(tempDir);

            // Create template XLSX file from resource
            string dllpath = Assembly.GetExecutingAssembly().CodeBase;
            dllpath = dllpath.Substring(8);
            string templateFile = Path.GetDirectoryName(dllpath) + "/template.xlsx";
            if (!Directory.Exists(Path.GetDirectoryName(templateFile)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(templateFile));
            }
            if (!File.Exists(templateFile))
            {
                File.WriteAllBytes(templateFile, Resource.template);
            }

            // Unzip template XLSX file to the temporary directory.
            XlsxFileHelper.UnzipFile(templateFile, tempDir);

            //删除模板文件
            if (File.Exists(templateFile))
            {
                File.Delete(templateFile);
            }

            //将表名写入workbook.xml文件
            IList<string> tableNames = new List<string>();
            tableNames.Add(data.TableName);
            WriteSheetNamesToWorkbook(Path.Combine(tempDir, @"xl\workbook.xml"), tableNames);

            //using (var stream = new FileStream(Path.Combine(tempDir, @"xl\workbook.xml"), FileMode.Open, FileAccess.ReadWrite))
            //    // ..and fill it with unique strings used in the workbook
            //    WriteSheetNamesToWorkbook(stream, sheetNames);

            // We will need two string tables; a lookup IDictionary<string, int> for fast searching 
            // an ordinary IList<string> where items are sorted by their index.
            IDictionary<string, int> lookupTable;

            // Call helper methods which creates both tables from input data.
            var stringTable = CreateStringTable(data, out lookupTable);

            // Create XML file..
            using (var stream = new FileStream(Path.Combine(tempDir, @"xl\sharedStrings.xml"), FileMode.Create))
                // ..and fill it with unique strings used in the workbook
                WriteStringTable(stream, stringTable);

            // Create XML file..
            using (var stream = new FileStream(Path.Combine(tempDir, @"xl\worksheets\sheet1.xml"), FileMode.Create))
                // ..and fill it with rows and columns of the DataTable.
                WriteWorksheet(stream, data, lookupTable, isFirstRowColumnNames);

            // ZIP temporary directory to the XLSX file.
            XlsxFileHelper.ZipDirectory(tempDir, fileName);

            if (Directory.Exists(tempDir))
            {
                Directory.Delete(tempDir, true);
            }

        }

        public void Write(string fileName, DataSet dataSet, bool isFirstRowColumnNames)
        {
            if (Path.GetExtension(fileName).ToLower() != ".xlsx") throw new ArgumentException("不能写入扩展名非.xlsx的文件！");
            string rootDir = Path.GetDirectoryName(fileName);
            string tempDir = rootDir + "/" + Path.GetFileNameWithoutExtension(fileName);

            if (!Directory.Exists(tempDir))
            {
                Directory.CreateDirectory(tempDir);
            }

            // Delete contents of the temporary directory.
            XlsxFileHelper.DeleteDirectoryContents(tempDir);

            // Create template XLSX file from resource
            string dllpath = Assembly.GetExecutingAssembly().CodeBase;
            dllpath = dllpath.Substring(8);
            string templateFile = Path.GetDirectoryName(dllpath) + "/template.xlsx";
            if (!Directory.Exists(Path.GetDirectoryName(templateFile)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(templateFile));
            }
            if (!File.Exists(templateFile))
            {
                File.WriteAllBytes(templateFile, Resource.template);
            }

            // Unzip template XLSX file to the temporary directory.
            XlsxFileHelper.UnzipFile(templateFile, tempDir);

            //删除模板文件
            if (File.Exists(templateFile))
            {
                File.Delete(templateFile);
            }

            //将表名写入workbook.xml文件
            IList<string> tableNames = new List<string>();
            foreach (DataTable dt in dataSet.Tables)
            {
                tableNames.Add(dt.TableName);
            }
            WriteSheetNamesToWorkbook(Path.Combine(tempDir, @"xl\workbook.xml"), tableNames);

            // We will need two string tables; a lookup IDictionary<string, int> for fast searching 
            // an ordinary IList<string> where items are sorted by their index.
            IDictionary<string, int> lookupTable;

            // Call helper methods which creates both tables from input data.
            var stringTable = CreateStringTable(dataSet, out lookupTable);

            // Create XML file..
            using (var stream = new FileStream(Path.Combine(tempDir, @"xl\sharedStrings.xml"), FileMode.Create))
                // ..and fill it with unique strings used in the workbook
                WriteStringTable(stream, stringTable);

            for (int i = 0; i != dataSet.Tables.Count;i++ )
            {
                string sheetFileName = @"xl\worksheets\sheet" + (i+1).ToString() + ".xml";
                // Create XML file..
                using (var stream = new FileStream(Path.Combine(tempDir, sheetFileName), FileMode.Create))
                    // ..and fill it with rows and columns of the DataTable.
                    WriteWorksheet(stream, dataSet.Tables[i], lookupTable, isFirstRowColumnNames);
            }

            // ZIP temporary directory to the XLSX file.
            XlsxFileHelper.ZipDirectory(tempDir, fileName);

            if (Directory.Exists(tempDir))
            {
                Directory.Delete(tempDir, true);
            }
        }
    }
}
