using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using PdkXlsxFileIO;

namespace CA4ExcelIO
{
    class Program
    {
        static void Main(string[] args)
        {
            DataSet ds=null;
            XlsxFileReader reader = new XlsxFileReader();
            reader.ProgressPercentChanged += reader_ProgressPercentChanged;

            reader.DateTimeDataColumnName.Add("时间");
            ds = reader.Read("t1.xlsx", true);
            if (ds != null)
            {
                DataTable dt = ds.Tables[0];
                DataRow dr = dt.Rows[8];
                for (int i = 0; i != dr.ItemArray.Length; i++)
                {
                    Console.WriteLine("["+dt.TableName+"].("+dt.Columns[i].ColumnName + " : " + dr.ItemArray[i]+")");
                }
            }
            Console.WriteLine("Reading Over!");


            XlsxFileWriter writer = new XlsxFileWriter();
            writer.ProgressPercentChanged += writer_ProgressPercentChanged;
            writer.Write("output.xlsx", ds, true);
            Console.WriteLine("Writing Over!");

            ds = reader.Read("output.xlsx", true);
            Console.WriteLine("Rereading Over!");

            Console.ReadKey();

        }

        static void writer_ProgressPercentChanged(int percent)
        {
            //throw new NotImplementedException();
            Console.WriteLine(percent + "%");
        }

        static void reader_ProgressPercentChanged(int percent)
        {
            //throw new NotImplementedException();
            Console.WriteLine(percent + "%");
        }
    }
}
