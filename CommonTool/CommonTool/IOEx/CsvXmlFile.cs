using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Runtime.Serialization;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace V1000Option
{
    public class CsvXmlFile
    {

        // コンストラクション
        public CsvXmlFile()
        {
               
        }

        private StreamWriter _Files;
        private int CSV_FLUSH_COUNT = 1000;
        private int nWriteNum = 0;
        public bool OpenFile(string filePath)
        {
            try
            {
                _Files = new StreamWriter(filePath, false, System.Text.Encoding.GetEncoding("UTF-8"));
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        public bool CloseFile()
        {
            try
            {
                if (null != _Files)
                {
                    _Files.Flush();
                    _Files.Close();
                    _Files.Dispose();
                }
                _Files = null;
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        public bool WriteLine(string[] strItemArray)
        {
            StringBuilder stringBuild = new StringBuilder("");
            foreach (string sData in strItemArray)
            {
                if (stringBuild.Length > 0) stringBuild.Append(",");

                if (!(System.Text.RegularExpressions.Regex.Match(sData, "^[0-9\\-+.]+$")).Success)
                {
                    string sTemp = sData.Replace(@"""", @"""""");
                    if(sTemp != "")
                        sTemp = @"""" + sTemp + @"""";
                    stringBuild.Append(sTemp);
                }
                else
                {
                    stringBuild.Append(sData);
                }                                
                        
            }
            try
            {
                // CSVファイル書込
                _Files.Write(stringBuild.ToString());
                _Files.Write(Environment.NewLine);

                nWriteNum++;
                if (nWriteNum > CSV_FLUSH_COUNT)
                {
                    nWriteNum = 0;
                    _Files.Flush();
                }
            }
            catch (Exception)
            {
                
                return false;
            }
            return true;
        }

        public void OutputFile(string filePath, DataTable dt, List<string> itmList)
        {

            if (filePath.Length < 1) return;

            if (System.IO.Path.GetExtension(filePath) == ".csv")
            {
                // CSVファイルで出力
                CreateCsvFile(filePath, dt, itmList);
            }
        }

        private void CreateCsvFile(string filePath, DataTable dt, List<string> itmList)
        {
            StreamWriter sw = null;
            try
            {
                // CSVファイルオープン
                sw = new StreamWriter(filePath, false, System.Text.Encoding.GetEncoding("UTF-8"));

                // ヘッダー追記                
                StringBuilder header = new StringBuilder("");
                for (int i = 0; i < itmList.Count; i++)
                {
                    if (header.Length > 0) header.Append(",");

                    if (!(System.Text.RegularExpressions.Regex.Match(itmList[i], "^[0-9\\-+.]+$")).Success)
                    {
                        string sTemp = itmList[i].Replace(@"""", @"""""");
                        if (sTemp != "")
                            sTemp = @"""" + sTemp + @"""";
                        header.Append(sTemp);
                    }
                    else {
                        header.Append(itmList[i]);
                    }
                }
                // CSVファイル書込
                sw.Write(header.ToString());
                sw.Write(Environment.NewLine);
                Stopwatch sw1 = new Stopwatch();
                sw1.Start();
                
                for (int r = 0; r < dt.Rows.Count; r++)
                {
                    StringBuilder stringBuild = new StringBuilder("");
                        for (int c = 0; c < dt.Columns.Count; c++)
                        {
                                if (stringBuild.Length > 0) stringBuild.Append(",");

                                string sData = dt.Rows[r][c].ToString();
                                if (!(System.Text.RegularExpressions.Regex.Match(sData, "^[0-9\\-+.]+$")).Success)
                                {
                                    string sTemp = sData.Replace(@"""", @"""""");
                                    if(sTemp != "")
                                        sTemp = @"""" + sTemp + @"""";
                                    stringBuild.Append(sTemp);
                                }
                                else
                                {
                                    stringBuild.Append(sData);
                                }                                
                        
                        }
           
                    // CSVファイル書込
                    sw.Write(stringBuild.ToString());
                    sw.Write(Environment.NewLine);

                }
                sw1.Stop();
                TimeSpan ts = sw1.Elapsed;
                Console.WriteLine("---------------------------------- CreateCsvFile {0}", ts.TotalMilliseconds);
                // CSVファイルクローズ
                sw.Close();
                sw.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                // CSVファイルクローズ
                if (sw != null) sw.Close();
            }
        }



        /// <summary>
        /// configurationファイル(.xml)にセーブ
        /// </summary>
        /// <returns>true:正常にロード, false:ロード失敗</returns>
        public bool SaveConfigFile<T>(T t, string configFileName)
        {
            try
            {
                DirectoryInfo di =
                    new DirectoryInfo(
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                            @"Sony\Topss"));
                if (!di.Exists)
                {
                    Directory.CreateDirectory(
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                            @"Sony\Topss"));
                }

                string configfile =
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                        @"Sony\Topss", configFileName);

                Stream stream = File.Open(configfile, FileMode.Create, FileAccess.Write, FileShare.Write);

                var serializer = new XmlSerializer(t.GetType());
                serializer.Serialize(stream, t);
                stream.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
            return true;
        }

        /// <summary>
        /// configurationファイル(.xml)をロード
        /// </summary>
        /// <returns>true:正常にロード, false:ロード失敗</returns>
        public bool LoadConfigFile<T>(ref T t, string configFileName)
        {

            string configfile = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), @"Sony\Topss", configFileName);
            FileInfo fi = new FileInfo(configfile);
            if (fi.Exists)
            {
                Stream stream = File.Open(configfile, FileMode.Open, FileAccess.Read, FileShare.Read); // サービス間の競合を防止。
                var serializer = new XmlSerializer(t.GetType());
                try
                {
                    t = (T)serializer.Deserialize(stream);
                }
                catch (Exception ex)
                {
                    return false;
                }
                finally
                {
                    stream.Close();
                }
            }
            else
            {
                // デフォールト値の設定ファイルを作成
                SaveConfigFile(t, configFileName);
                return true;
            }
            return true;
        }
    }

    /// <summary>
    /// </summary>
    /// <typeparam name="TKey"></typeparam>
    /// <typeparam name="TValue"></typeparam>
    [XmlRoot("SerializableDictionary")]
    public class SerializableDictionary<TKey, TValue>
        : Dictionary<TKey, TValue>, IXmlSerializable
    {
        #region 

        public SerializableDictionary()
            : base()
        {
        }

        public SerializableDictionary(IDictionary<TKey, TValue> dictionary)
            : base(dictionary)
        {
        }

        public SerializableDictionary(IEqualityComparer<TKey> comparer)
            : base(comparer)
        {
        }

        public SerializableDictionary(int capacity)
            : base(capacity)
        {
        }

        public SerializableDictionary(int capacity, IEqualityComparer<TKey> comparer)
            : base(capacity, comparer)
        {
        }

        protected SerializableDictionary(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }

        #endregion 

        #region IXmlSerializable Members

        public System.Xml.Schema.XmlSchema GetSchema()
        {
            return null;
        }

        /// <summary>
        /// </summary>
        /// <param name="reader"></param>
        public void ReadXml(System.Xml.XmlReader reader)
        {
            XmlSerializer keySerializer = new XmlSerializer(typeof (TKey));
            XmlSerializer valueSerializer = new XmlSerializer(typeof (TValue));
            bool wasEmpty = reader.IsEmptyElement;
            reader.Read();
            if (wasEmpty)
                return;
            while (reader.NodeType != System.Xml.XmlNodeType.EndElement)
            {
                reader.ReadStartElement("item");
                reader.ReadStartElement("key");
                TKey key = (TKey) keySerializer.Deserialize(reader);
                reader.ReadEndElement();
                reader.ReadStartElement("value");
                TValue value = (TValue) valueSerializer.Deserialize(reader);
                reader.ReadEndElement();
                this.Add(key, value);
                reader.ReadEndElement();
                reader.MoveToContent();
            }
            reader.ReadEndElement();
        }

        /**/

        /// <summary>
        /// </summary>
        /// <param name="writer"></param>
        public void WriteXml(System.Xml.XmlWriter writer)
        {
            XmlSerializer keySerializer = new XmlSerializer(typeof (TKey));
            XmlSerializer valueSerializer = new XmlSerializer(typeof (TValue));
            foreach (TKey key in this.Keys)
            {
                writer.WriteStartElement("item");
                writer.WriteStartElement("key");
                keySerializer.Serialize(writer, key);
                writer.WriteEndElement();
                writer.WriteStartElement("value");
                TValue value = this[key];
                valueSerializer.Serialize(writer, value);
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
        }

        #endregion IXmlSerializable Members

    }

}
