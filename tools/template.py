CSHARP_TEMPLATE = """
//Building

using System.Collections.Generic;

namespace SKS.Model
{
    /// <summary> ##EXCEL_FILENAME## </summary>
    public class ##SHEET_NAME##:ConfigBase
    {
        public class Record
        {
##KEY_DEFINE##

            public void Decode(StreamTool st)
            {
                var In = st.In;
##DECODE##
            }
        }
        public Dictionary<int, Record> map = new Dictionary<int, Record>();
        public List<Record> list = new List<Record>();

        public override void Decode(StreamTool st)
        {
            var In = st.In;
            map.Clear();
            list.Clear();
            var n_map = In.ReadInt32();
            //UnityEngine.Debug.Log("n_map:" + n_map);
            for (int i = 0, n = n_map; i < n; ++i)
            {
                var key = In.ReadInt32();
                var record = new Record();
                record.Decode(st);
                list.Add(record);
                map.Add(key, record);
            }
        }
        /*
        public void Encode(StreamTool st)
        {
            var Out = st.Out;
            var n = list.Count;
            Out.Write(n);
            for(int i=0;i<n;++i)
            {
                var record = list[i];
                Out.Write(record.id);
                record.Encode(st);
            }
        } */

        public Record this[int id]
        {
            get
            {
                Record ret = null;
                map.TryGetValue(id, out ret);
                if(null == ret)
                {
                    UnityEngine.Debug.LogError(string.Format("{0} sheet has not id:{1}", "##SHEET_NAME##", id));
                }
                return ret;
            }
        }//[]
    }//class Building	
}//namespace SKS

"""