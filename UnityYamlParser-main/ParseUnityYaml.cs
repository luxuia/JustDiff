using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using VYaml.Serialization;

namespace UnityYamlParser
{
    public class GameObject
    {
        public string name;
        public List<string> comps = new List<string>();

        public List<GameObject> childs = new List<GameObject>();

        public dynamic data;
    };

    public class GameObjectComparer : IEqualityComparer<GameObject>
    {
        public bool Equals(GameObject a, GameObject b)
        {
            return a.name == b.name;
        }

        public int GetHashCode(GameObject a)
        {
            return a.name.GetHashCode();
        }
    }


    public class Scene
    {
        public List<GameObject> roots = new List<GameObject>();
    };
    public class ParseUnityYaml
    {
   
        public static Scene ParseYaml(string path)
        {
            var str = File.ReadAllText(path);

            var idlist = new List<long>();
            var fileid_map = new Dictionary<long, int>();
            foreach (var item in Regex.Matches(str, @"---.+&(\d+)"))
            {
                var match = item as Match;
                if (match != null)
                {
                    //Console.WriteLine(item.ToString());
                    var id = long.Parse(match.Groups[1].Value);

                    idlist.Add(id);
                    fileid_map[id] = idlist.Count - 1;
                }
            }

            var bytes = System.Text.Encoding.UTF8.GetBytes(str);
            var yaml = YamlSerializer.DeserializeMultipleDocuments<dynamic>(bytes).ToArray();

            var all_gos = new Dictionary<int, GameObject>();
            var ref_gos = new Dictionary<int, GameObject>();
            var scene = new Scene();
            for (var i = 0; i < yaml.Length; i++)
            {
                var item = yaml[i];
                if (item.ContainsKey("GameObject"))
                {
                    all_gos[i] = new GameObject() { data = item["GameObject"] };
                }
                else if (item.ContainsKey("PrefabInstance"))
                {
                    ref_gos[i] = new GameObject() { data = item["PrefabInstance"] };
                }
            }
            var trans2go = new Dictionary<int, GameObject>();
            var wait_trans = new List<Tuple<dynamic, GameObject>>();

            foreach (var v in all_gos)
            {
                var go_fileid = idlist[v.Key];
                var go = v.Value;

                var data = go.data;
                var comps = data["m_Component"] as List<dynamic>;

                go.name = data["m_Name"];

                foreach (var com in comps)
                {
                    var com_fileid = com["component"]["fileID"];

                    if (fileid_map.ContainsKey(com_fileid))
                    {
                        var com_idx = fileid_map[com_fileid];
                        var comdata = yaml[com_idx];

                        string key = string.Empty;
                        foreach (var kk in comdata.Keys)
                        {
                            key = kk;
                        }

                        go.comps.Add(key);

                        if (key == "Transform" || key == "RectTransform")
                        {
                            // 认为com的idx和go的idx一直
                            trans2go[com_idx] = go;

                            var trans = comdata[key];
                            wait_trans.Add(new Tuple<dynamic, GameObject>(trans, go));

                            
                        }
                    }
                }
            }
            foreach (var v in wait_trans)
            {
                var trans = v.Item1;
                var go = v.Item2;

                var father = trans["m_Father"];
                var father_id = father["fileID"];
                if (father_id == 0)
                {
                    scene.roots.Add(go);
                }
                else
                {
                    if (fileid_map.ContainsKey(father_id))
                    {
                        var trans_id = fileid_map[father_id];
                        var parentgo = trans2go[trans_id];
                        parentgo.childs.Add(go);
                    }
                }
            }

            foreach (var v in ref_gos)
            {
                var go_fileid = idlist[v.Key];
                var go = v.Value;

                var data = go.data;
                var modify = data["m_Modification"]["m_Modifications"];
                foreach (var m in modify)
                {
                    if (m["propertyPath"] == "m_Name")
                    {
                        go.name = m["value"] + "[REF]";
                    }
                }
                var father = data["m_Modification"]["m_TransformParent"];
                if (fileid_map.ContainsKey(father["fileID"]))
                {
                    var parentgo = trans2go[fileid_map[father["fileID"]]];
                    parentgo.childs.Add(go);
                }
            }
            return scene;
        }
    }
}
