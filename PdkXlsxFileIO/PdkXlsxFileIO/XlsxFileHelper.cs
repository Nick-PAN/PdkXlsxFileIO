using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.IO.Compression;

namespace PdkXlsxFileIO
{
    class XlsxFileHelper
    {
        /// <summary>
        /// 删除目录内容
        /// </summary>
        /// <param name="directory"></param>
        public static void DeleteDirectoryContents(string directory)
        {
            var info = new DirectoryInfo(directory);
            if (!info.Exists) return;
            foreach (var file in info.GetFiles())
            {
                file.Delete();
            }

            foreach (var dir in info.GetDirectories())
            {
                dir.Delete(true);
            }
        }

        /// <summary>
        /// 解压zip文件到指定的文件夹
        /// </summary>
        /// <param name="zipFileName"></param>
        /// <param name="targetDirectory"></param>
        public static void UnzipFile(string zipFileName, string targetDirectory)
        {
            //new FastZip().ExtractZip(zipFileName, targetDirectory, null);

            ZipFile.ExtractToDirectory(zipFileName, targetDirectory);
        }

        public static void ZipDirectory(string sourceDirectory, string zipFileName)
        {
            //new FastZip().CreateZip(zipFileName, sourceDirectory, true, null);
            if (File.Exists(zipFileName))
            {
                File.Delete(zipFileName);
            }
            ZipFile.CreateFromDirectory(sourceDirectory, zipFileName);
        }

        public static int ProgressPercent(int currentRowS, int startRowSn, int finialRowSn)
        {
            return (currentRowS - startRowSn) * 100 / (finialRowSn - startRowSn+1);
        }
    }
}
