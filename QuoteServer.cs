﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Net;
using System.Diagnostics;
using ExcelDna.Integration;
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

        #region IRtdServer Members

        public object ConnectData(int topicId, ref Array Strings, ref bool newValues)
        {
            /*
             * Strings is an array contains two strings
             * The first is the quotation id, such as sh000300...
             * The second is the item of the quotation, such as "close", "volume"
             */
            if (Strings.Length != 2)
            {
                throw new Exception("Invalid topic parameters");
            }

            string id = Strings.GetValue(0) as string;
            if (id == null || id.Length == 0)
            {
                throw new Exception("Invalid quotation ID");
            }
            string item = Strings.GetValue(1) as string;
            if (item == null || item.Length == 0)
            {
                throw new Exception("Invalid item");
            }

            m_topics.Add(topicId, new Topic(id, item));

            return "Queued";
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
            m_timer = new Timer();
            m_timer.Tick += Callback;
            m_timer.Interval = 2000;
            m_timer.Start();

            return 1;
        }

        public void ServerTerminate()
        {
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
            Debug.Print(url.ToString());

            // Submit request
            m_webClient.Encoding = Encoding.Default;
            string result = m_webClient.DownloadString(url.ToString());
            Debug.Print(result);

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
                            int idx = Int32.Parse(t.Item);
                            t.Value = items[idx];
                        }
                    }
                }
            }
        }
    }

    //========================================
    // Topic represents a RTD call in Excel
    //========================================
    public class Topic
    {
        private string m_ID;    // m_ID is the quote id, such as "sh000001"
        private string m_Item;  // m_Item is a quote item, currently it only supports integer
                                // e.g. "1" means the second quote item (0-based index)
        private string m_Value;     // m_Value the quote value
        private bool m_bUpdated;    // If the topic value has changed

        // Attribute operations
        public string Id
        {
            get { return m_ID; }
            set { m_ID = value; }
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
        public Topic(string Id, string Item)
        {
            m_ID = Id;
            m_Item = Item;
            m_Value = "";
        }
    }
}