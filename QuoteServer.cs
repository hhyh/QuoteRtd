using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;

namespace QuoteRtd
{
    // Need call regsvr32 "QuoteRtd-AddIn-packed.xll" to activate this COM server
    [Guid("9EADD693-9F57-4403-B56C-5BD48B12DB5B")]
    [ComVisible(true)]
    public class QuoteServer : IRtdServer
    {
        private Dictionary<int, Topic> _topics;
        private IRTDUpdateEvent _callback;
        private System.Windows.Forms.Timer _timer;

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

            _topics.Add(topicId, new Topic(id, item));

            return "Queued";
        }

        public void DisconnectData(int topicID)
        {
            _topics.Remove(topicID);
        }

        public int Heartbeat()
        {
            return 1;
        }

        public Array RefreshData(ref int TopicCount)
        {
            object[,] results = new object[2, _topics.Count];
            TopicCount = 0;

            foreach (int topicID in _topics.Keys)
            {
                if (_topics[topicID].Updated == true)
                {
                    results[0, TopicCount] = topicID;
                    results[1, TopicCount] = _topics[topicID].Value + " : " + Convert.ToString(DateTime.Now);
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
            _topics = new Dictionary<int, Topic>();
            _callback = CallbackObject;
            _timer = new System.Windows.Forms.Timer();
            _timer.Tick += Callback;
            _timer.Interval = 2000;
            _timer.Start();

            return 1;
        }

        public void ServerTerminate()
        {
            _timer.Dispose();
            _topics = null;
        }

        #endregion

        private void Callback(object sender, EventArgs e)
        {

            lock (_topics)
            {
                try
                {
                    /*
                     * Multiple topics can use same quotation id
                     * Put them together in a list and put all lists in a dictionary
                     * which use quotation id as key
                     */
                    Dictionary<string, List<Topic>> ids =
                        new Dictionary<string, List<Topic>>();
                    foreach (KeyValuePair<int, Topic> x in _topics)
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

                    string[] keys = new string[ids.Keys.Count];
                    ids.Keys.CopyTo(keys, 0);
                    GetDataFromServer(keys);
                }
                catch (Exception)
                {
                }
            }

            _callback.UpdateNotify();
        }

        private void GetDataFromServer(string[] ids)
        {
            string url = "http://hq.sinajs.cn/list=";
        }
    }

    public class Topic
    {
        private string iId;
        private string iItem;
        private string iValue;
        private bool iUpdated;
        public string Id
        {
            get { return iId; }
            set { iId = value; }
        }
        public string Item
        {
            get { return iItem; }
            set { iItem = value; }
        }
        public string Value
        {
            get { return iValue; }
            set
            {
                if (iValue != value)
                {
                    iValue = value;
                    iUpdated = true;
                }
                else
                {
                    iUpdated = false;
                }
            }
        }
        public bool Updated
        {
            get { return iUpdated; }
            set { iUpdated = value; }
        }

        public Topic(string Id, string Item)
        {
            iId = Id;
            iItem = Item;
            iValue = "";
        }
    }
}
