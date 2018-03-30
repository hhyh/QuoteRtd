using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Net;
using System.Xml;

using ExcelDna.Integration.Rtd;
using ExcelDna.Logging;

namespace QuoteRtd
{
    // Need call regsvr32 "QuoteRtd-AddIn-packed.xll" to activate this COM server
    [Guid("9EADD693-9F57-4403-B56C-5BD48B12DB5B")]
    [ComVisible(true)]
    public class QuoteServer : IRtdServer
    {
        private Dictionary<int, Topic> m_topics;
        private IRTDUpdateEvent m_callback;
        private Timer m_timer;
        private WebClient m_webClient;
        private XmlDocument m_xml;
        private bool m_saveXml = false;

        // Script will be loaded when creating the ScriptLoader
        private ScriptLoader m_scriptLoader = new ScriptLoader();

        #region IRtdServer Members

        public object ConnectData(int topicId, ref Array Strings, ref bool newValues)
        {
            /*
             * Strings is an array contains two strings
             * The first is the quotation id, such as sh000300...
             * The second is the item of the quotation, such as "close", "volume"
             */
            if (Strings.Length != 3)
            {
                throw new Exception("Invalid topic parameters");
            }

            string id = Strings.GetValue(0) as string;
            if (id == null || id.Length == 0)
            {
                throw new Exception("Invalid quotation ID");
            }
            string idx = Strings.GetValue(1) as string;
            string item = Strings.GetValue(2) as string;
            if (idx == null || idx.Length == 0)
            {
                throw new Exception("Invalid idx");
            }
            if (item == null || item.Length == 0)
            {
                throw new Exception("Invalid item name");
            }

            m_topics.Add(topicId, new Topic(id, idx, item));

            // Load saved data from xml
            XmlElement rootNode = m_xml.DocumentElement;
            XmlElement idNode = rootNode.SelectSingleNode("/data/"+id) as XmlElement;
            if (idNode == null)
            {
                idNode = m_xml.CreateElement(id);
                rootNode.AppendChild(idNode);
            }
            XmlElement itemNode = idNode.SelectSingleNode(item) as XmlElement;
            if (itemNode == null)
            {
                itemNode = m_xml.CreateElement(item);
                itemNode.InnerText = "Queued";
                idNode.AppendChild(itemNode);
            }
            return itemNode.InnerText;
        }

        public void DisconnectData(int topicID)
        {
            m_topics.Remove(topicID);
        }

        public int Heartbeat()
        {
            return 1;
        }

        public Array RefreshData(ref int TopicCount)
        {
            object[,] results = new object[2, m_topics.Count];
            TopicCount = 0;

            foreach (int topicID in m_topics.Keys)
            {
                if (m_topics[topicID].Updated == true)
                {
                    results[0, TopicCount] = topicID;
                    results[1, TopicCount] = m_topics[topicID].Value;
                    TopicCount++;
                }
            }

            object[,] temp = new object[2, TopicCount];
            for (int i = 0; i < TopicCount; i++)
            {
                temp[0, i] = results[0, i];
                temp[1, i] = results[1, i];
            }

            return temp;
        }

        public int ServerStart(IRTDUpdateEvent CallbackObject)
        {
            m_topics = new Dictionary<int, Topic>();
            m_callback = CallbackObject;
            m_webClient = new WebClient();

            // Load stored data
            m_xml = new XmlDocument();
            try
            {
                m_xml.Load("data.xml");
            }
            catch (Exception e)
            {
                m_xml.LoadXml("<data></data>");
            }

            // Create timer
            m_timer = new Timer();
            m_timer.Tick += Callback;
            m_timer.Interval = GlobalConfig.refreshInterval;
            m_timer.Start();

            return 1;
        }

        public void ServerTerminate()
        {
            if (m_saveXml)
                m_xml.Save("data.xml");
            m_timer.Dispose();
            m_topics = null;
        }

        #endregion

        //======================================================
        // Timer callback
        // Get data from server and update every topic.
        //======================================================
        private void Callback(object sender, EventArgs e)
        {
            if (!GlobalConfig.refreshData)
                return;

            // Stop the timer to prevent re-enter
            m_timer.Stop();

            lock (m_topics)
            {
                try
                {
                    // Multiple topics can use same quotation id
                    // Put them together in a list and put all lists in a dictionary
                    // which use quotation id as key
                    Dictionary<string, List<Topic>> ids =
                        new Dictionary<string, List<Topic>>();
                    foreach (KeyValuePair<int, Topic> x in m_topics)
                    {
                        if (!ids.ContainsKey(x.Value.Id))
                        {
                            // New quotation ID
                            // Allocate a new list for this id
                            ids.Add(x.Value.Id, new List<Topic>());
                        }
                        // Put the topic to corresponding list
                        ids[x.Value.Id].Add(x.Value);
                    }

                    // Fetch data
                    string[] keys = new string[ids.Keys.Count];
                    ids.Keys.CopyTo(keys, 0);
                    GetDataFromServer(keys);
                }
                catch (Exception ex)
                {
                    // Show error in logging window
                    LogDisplay.WriteLine(ex.ToString());
                }
            }

            m_callback.UpdateNotify();

            // Restart the timer
            m_timer.Interval = GlobalConfig.refreshInterval;
            m_timer.Start();
        }

        //===================================================
        // Fetch data from sina source
        //===================================================
        private void GetDataFromServer(string[] ids)
        {
            StringBuilder url = new StringBuilder("http://hq.sinajs.cn/list=");

            // Build the request URL
            foreach (string id in ids)
            {
                url.AppendFormat("{0},", id);
            }
            // Remove the last ','
            url.Remove(url.Length - 1, 1);
            if (GlobalConfig.logEnable)
                LogDisplay.WriteLine(url.ToString());

            // Submit request
            m_webClient.Encoding = Encoding.Default;
            string result = m_webClient.DownloadString(url.ToString());
            if (GlobalConfig.logEnable)
                LogDisplay.WriteLine(result);

            // Parse the result
            string[] lines = result.Split('\n');
            Regex regex = new Regex("var hq_str_(.*)=\"(.*?)\"");

            foreach (string line in lines)
            {
                // Parse the result line
                Match match = regex.Match(line);

                // Parse successfully
                if (match.Groups.Count > 1)
                {
                    // Split the result by ','
                    string[] items = match.Groups[2].Value.Split(',');

                    // Search topics which are interested in this result line
                    foreach (KeyValuePair<int, Topic> x in m_topics)
                    {
                        Topic t = x.Value;

                        // If this line belongs to this topic, update the topic value
                        if (t.Id == match.Groups[1].Value)
                        {
                            // Update value
                            int idx = Int32.Parse(t.Idx);
                            t.Value = items[idx];

                            // Save the data in XML
                            XmlElement node = m_xml.DocumentElement.
                                SelectSingleNode("/data/" + t.Id + "/" + t.Item) as XmlElement;
                            node.InnerText = t.Value;
                            m_saveXml = true;
                        }
                    }
                }
            }
        }
    }

    //========================================
    // Topic represents an RTD cell in Excel
    //========================================
    public class Topic
    {
        private string m_ID;    // m_ID is the quote id, such as "sh000001"
        private string m_Idx;   // m_Idx is the index of the quote item
        private string m_Item;  // m_Item is the name of quote item, such as "name", "open"
        private string m_Value;     // m_Value the quote value
        private bool m_bUpdated;    // If the topic value has changed

        // Attributes
        public string Id
        {
            get { return m_ID; }
            set { m_ID = value; }
        }
        public string Idx
        {
            get { return m_Idx; }
            set { m_Idx = value; }
        }
        public string Item
        {
            get { return m_Item; }
            set { m_Item = value; }
        }
        public string Value
        {
            get { return m_Value; }
            set
            {
                if (m_Value != value)
                {
                    m_Value = value;
                    m_bUpdated = true;
                }
                else
                {
                    m_bUpdated = false;
                }
            }
        }
        public bool Updated
        {
            get { return m_bUpdated; }
            set { m_bUpdated = value; }
        }

        // Constructor
        public Topic(string Id, string Idx, string Item)
        {
            m_ID = Id;
            m_Idx = Idx;
            m_Item = Item;
            m_Value = "";
        }
    }
}
