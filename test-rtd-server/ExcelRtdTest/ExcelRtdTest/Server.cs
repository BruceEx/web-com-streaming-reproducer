using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace ExcelRtdTest
{
    [Guid(Constants.ServerGuid)]
    [ComVisible(true)]
    [ProgId(Constants.ServerProgId)]
    public class Server : IRtdServer
    {
        private IRTDUpdateEvent m_callback;
        private Timer m_timer;
        private List<int> m_topics;
        private static int m_count;

        public int ServerStart(IRTDUpdateEvent callback)
        {
            m_callback = callback;

            m_timer = new Timer();
            m_timer.Tick += new EventHandler(TimerEventHandler);
            m_timer.Interval = 2000;

            m_topics = new List<int>();

            return 1;
        }

        public void ServerTerminate()
        {
            if (null != m_timer)
            {
                m_timer.Dispose();
                m_timer = null;
            }
        }

        public object ConnectData(int topicId,
                                  ref Array strings,
                                  ref bool newValues)
        {
            m_topics.Add(topicId);
            m_timer.Start();
            return GetCount();
        }

        public void DisconnectData(int topicId)
        {
            m_topics.Remove(topicId);
        }

        public Array RefreshData(ref int topicCount)
        {
            object[,] data = new object[2, m_topics.Count];

            int index = 0;

            foreach (int topicId in m_topics)
            {
                data[0, index] = topicId;
                data[1, index] = GetCount();

                ++index;
            }

            topicCount = m_topics.Count;

            m_timer.Start();
            return data;
        }

        public int Heartbeat()
        {
            return 1;
        }

        private void TimerEventHandler(object sender,
                                       EventArgs args)
        {
            m_timer.Stop();
            m_callback.UpdateNotify();
        }

        private static string GetCount()
        {
            return $"COM: {++m_count}";
        }
    }
}
