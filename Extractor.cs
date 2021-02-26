using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using LibGit2Sharp;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Win32;
using System.Diagnostics;

namespace ReversionExtract
{
    public abstract class Extractor
    {
        public string path;
        public string outputDir;
        public List<List<object>> InputMatrix;
        public int RowNum;
        public static FastZip fz = new FastZip();

        public int total = 30;
        public int current { set; get; }
        public uint starttime { set; get; }
        public uint endtime { set; get; }
        public abstract void doExtract();
    }
    class GitExtractor : Extractor
    {
        [DllImport("winmm")]
        static extern uint timeGetTime();
        public GitExtractor(string path, string outputDir, List<List<object>> InputMatrix, int RowNum)
        {
            this.path = path;
            this.outputDir = outputDir;
            this.InputMatrix = InputMatrix;
            this.RowNum = RowNum;
        }
        public override void doExtract()
        {
            current = 0;
            String projectName = System.IO.Path.GetFileNameWithoutExtension(path);
            using (var repo = new Repository(path))
            {
                starttime = timeGetTime();
                for (int i = 1; i < RowNum; i++)
                {
                    String hashCode = Convert.ToString(InputMatrix[i][0]);
                    //git archive
                    Commit cm = repo.Lookup<Commit>(hashCode);
                    String tarpath = outputDir + "\\" + projectName;
                    String untarpath = outputDir + "\\" + projectName + "\\" + hashCode;
                    string tarfile = outputDir + "\\" + projectName + "\\" + hashCode+ ".zip";
                    String tarname = hashCode + ".zip";
                    if (!Directory.Exists(untarpath))
                    {
                        //创建
                        Directory.CreateDirectory(untarpath);
                    }
                    repo.ObjectDatabase.Archive(cm, tarfile);
                    UnRAR(untarpath, tarpath, tarname);
                    current++;
                }
                endtime = timeGetTime();
            }
        }
        /// <summary>
        /// 利用 WinRAR 进行解压缩
        /// </summary>
        /// <param name="path">文件解压路径（绝对）</param>
        /// <param name="rarPath">将要解压缩的 .rar 文件的存放目录（绝对路径）</param>
        /// <param name="rarName">将要解压缩的 .rar 文件名（包括后缀）</param>
        public void UnRAR(string path, string rarPath, string rarName)
        {
            ProcessStartInfo startinfo = new ProcessStartInfo(); 
            Process process = new Process();
            String rarexe = @"F:\WinRAR\WinRAR.exe"; //WinRAR安装位置
            //解压缩命令，相当于在要压缩文件(rarName)上点右键->WinRAR->解压到当前文件夹
            String cmd = string.Format("x {0} {1} -y",rarName,path);
            startinfo.FileName = rarexe;
            startinfo.Arguments = cmd;
            startinfo.WindowStyle = ProcessWindowStyle.Hidden;
            startinfo.WorkingDirectory = rarPath;
            process.StartInfo = startinfo;
            process.Start();
            process.WaitForExit();
            process.Dispose();
            process.Close();
            bool isdeletezip = true;
            if (isdeletezip) 
            {
                DeleteFile(rarPath + "\\"+rarName);
            }
        }
        ///
        /// 根据文件路径删除
        ///
        public static bool DeleteFile(string fPath)
        {
            if (File.Exists(fPath))
            {
                FileInfo fi = new FileInfo(fPath);
                fi.Attributes = FileAttributes.Normal;
                fi.Delete();               
            }
            return true;
        }
    }
}
