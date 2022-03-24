using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Timers;
using System.Configuration;
using System.Threading;

namespace CopyFileService
{
    public partial class CopyFileService : ServiceBase
    {

        #region field
        private string LogFileName =System.AppDomain.CurrentDomain.BaseDirectory+ @"OperationLog.txt";
        private string SourcePath = @"C:\Source";
        private string DestinationPath = @"C:\Distination";
        private OperatoionType OpType = OperatoionType.FileCopy;
        private FileSystemWatcher FSW;

        #endregion
        public CopyFileService()
        {
            InitializeComponent();

            //initial FileSystemWatcher
            FSW=new FileSystemWatcher(SourcePath);
            FSW.EnableRaisingEvents = true;
            FSW.Created += FSW_Created;
        }

        private void FSW_Created(object sender, FileSystemEventArgs e)
        {
            //检测文件是否被占用,如果文件被占用,等待300ms重试。
            int timeout = 180000;
            int occupiedTime = 0;
            while (IsFileUse(e.FullPath))
            {
                Thread.Sleep(300);
                occupiedTime += 300;
                if (occupiedTime==timeout)
                {
                    WriteLogs($"复制的文件:{e.FullPath}被占用，操作无法完成", DateTime.Now);
                }
            }

            try
            {
                if (OpType == OperatoionType.FileCopy)
                {
                    //复制源文件到目标文件夹
                    CopyFile(e.FullPath);
                }
                else
                {
                    //移动目标文件到目标文件夹
                    File.Move(e.FullPath, $"{DestinationPath}\\{e.Name}");
                }

                //log message
                string logContent = $"copy file {e.FullPath} →{DestinationPath}\\{e.Name}";
                WriteLogs(logContent, DateTime.Now);
            }
            catch (Exception ex)
            {

                WriteLogs( ex.Message +'\t'+"请重新启动服务", DateTime.Now);
            }
            
        }
        public static bool IsFileUse(string fileName)
        {
            bool isUsed = true;
            FileStream Fs = null;
            try
            {
                Fs=new FileStream(fileName, FileMode.Open, FileAccess.Read,FileShare.None);
                isUsed = false;
                Fs.Dispose();
            }
            catch (Exception)
            {
                
            }
            finally
            {
                if (Fs != null)
                {
                    Fs.Close(); 
                }
            }
            return isUsed;
        }

        protected override void  OnStart(string[] args)
        {
            //判断log记录文件是否存在,如果不存在则创建
            if (File.Exists(LogFileName)==false)
            {
              StreamWriter SW= File.CreateText(LogFileName);
              SW.Close();
            }

            //read app configuration file
            Configuration config = GetConfiguration();
            if (config.AppSettings.Settings.AllKeys.Contains<string>("SourcePath"))
            {
                SourcePath = config.AppSettings.Settings["SourcePath"].Value;
            }
            if (config.AppSettings.Settings.AllKeys.Contains<string>("DestinationPath"))
            {
                DestinationPath = config.AppSettings.Settings["DestinationPath"].Value;
            }
            if (config.AppSettings.Settings.AllKeys.Contains<string>("OperationType"))
            {
                OpType = (config.AppSettings.Settings["OperationType"].Value) == "FileCopy" ? OperatoionType.FileCopy : OperatoionType.FileMove;
            }
            config.Save();
            config=null;

            //设置监视文件夹筛选条件
            FSW.NotifyFilter = NotifyFilters.FileName
                                        | NotifyFilters.Attributes
                                        | NotifyFilters.CreationTime;

            //Write Log Service Startup
            WriteLogs("CopyFileService Startup", DateTime.Now);

        }
        protected override void OnShutdown()
        {
            this.OnStop();
            base.OnShutdown();
        }

        protected override void OnStop()
        {
            //Write config to configuration document
            string[] settings = new string[] { "SourcePath","DestinationPath","OperationType"};
            string[] settingvalue = new string[3];
            settingvalue[0] = SourcePath;
            settingvalue[1] = DestinationPath;
            settingvalue[2] = OpType.ToString();
            WriteConfiguration(settings, settingvalue);
            WriteLogs("CopyFileService Stop",DateTime.Now);
        }


        #region CustomizedMethod
        /// <summary>
        /// 从源文件夹复制文件到目标文件夹
        /// </summary>
        /// <param name="fileName">文件名（包含路径）</param>
        private void CopyFile(string fileName)
        {
            if (File.Exists(fileName))
                File.Copy(fileName, $"{DestinationPath}\\{Path.GetFileName(fileName)}",true);
        }
        /// <summary>
        /// 写入log文件
        /// </summary>
        /// <param name="Message">写入log文件的消息</param>
        /// <param name="opTime">操作的时间</param>
        private void WriteLogs(string Message, DateTime opTime)
        {
            File.AppendAllText(LogFileName, opTime.ToString("G") + '\t' + $"{Message}\n");
        }

        /// <summary>
        /// 向配置文件中添加指定的键值对数据
        /// </summary>
        /// <param name="key">要添加的键字符串数组</param>
        /// <param name="value">要添加的值字符串数组</param>
        private void WriteConfiguration(string[] key ,string[] value)
        {
            //Write config to config file
            Configuration currentExeConfig = GetConfiguration();
            var configkey = currentExeConfig.AppSettings.Settings; 
           for(int i = 0; i < key.Length; i++)
            {
                if (configkey.AllKeys.Contains<String>(key[i]))
                {
                    configkey[key[i]].Value= value[i];

                }
                else
                {
                    configkey.Add(key[i], value[i]);
                }
            }
           currentExeConfig.Save();
            currentExeConfig=null;

        }

        private Configuration GetConfiguration()
        {
            return ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        }

    #endregion
}
    public enum OperatoionType
    {
        FileCopy,
        FileMove
    }
}
