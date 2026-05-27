using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using VYaml.Serialization;

namespace ExcelMerge
{
    public class UnityGameObject
    {
        public string name;
        public List<string> comps = new List<string>();
        public List<UnityGameObject> childs = new List<UnityGameObject>();
        public dynamic data;
    }

    public class UnityGameObjectComparer : IEqualityComparer<UnityGameObject>
    {
        public bool Equals(UnityGameObject a, UnityGameObject b)
        {
            return a.name == b.name;
        }

        public int GetHashCode(UnityGameObject a)
        {
            return a.name.GetHashCode();
        }
    }

    public class UnityScene
    {
        public List<UnityGameObject> roots = new List<UnityGameObject>();
    }

    public class ParseUnityYaml
    {
        public static UnityScene ParseYaml(string path)
        {
            var str = File.ReadAllText(path);

            var idlist = new List<long>();
            var fileid_map = new Dictionary<long, int>();
            foreach (var item in Regex.Matches(str, @"---.+&(-?\d+)"))
            {
                var match = item as Match;
                if (match != null)
                {
                    var id = long.Parse(match.Groups[1].Value);
                    idlist.Add(id);
                    fileid_map[id] = idlist.Count - 1;
                }
            }

            var bytes = System.Text.Encoding.UTF8.GetBytes(str);
            var yaml = YamlSerializer.DeserializeMultipleDocuments<dynamic>(bytes).ToArray();

            var all_gos = new Dictionary<int, UnityGameObject>();
            var ref_gos = new Dictionary<int, UnityGameObject>();
            var scene = new UnityScene();
            for (var i = 0; i < yaml.Length; i++)
            {
                var item = yaml[i];
                if (item.ContainsKey("GameObject"))
                {
                    all_gos[i] = new UnityGameObject() { data = item["GameObject"] };
                }
                else if (item.ContainsKey("PrefabInstance"))
                {
                    ref_gos[i] = new UnityGameObject() { data = item["PrefabInstance"] };
                }
            }
            var trans2go = new Dictionary<int, UnityGameObject>();
            var wait_trans = new List<Tuple<dynamic, UnityGameObject>>();

            foreach (var v in all_gos)
            {
                var go = v.Value;
                var data = go.data;
                go.name = data.ContainsKey("m_Name") ? data["m_Name"]?.ToString() ?? "(unnamed)" : "(unnamed)";

                var comps = data.ContainsKey("m_Component") ? data["m_Component"] as List<dynamic> : null;
                if (comps == null) continue;

                foreach (var com in comps)
                {
                    long com_fileid = Convert.ToInt64(com["component"]["fileID"]);

                    if (fileid_map.ContainsKey(com_fileid))
                    {
                        var com_idx = fileid_map[com_fileid];
                        var comdata = yaml[com_idx];

                        string key = string.Empty;
                        foreach (var kk in comdata.Keys)
                        {
                            key = kk?.ToString() ?? "";
                        }

                        go.comps.Add(key);

                        if (key == "Transform" || key == "RectTransform")
                        {
                            trans2go[com_idx] = go;
                            var trans = comdata[key];
                            wait_trans.Add(new Tuple<dynamic, UnityGameObject>(trans, go));
                        }
                    }
                }
            }
            foreach (var v in wait_trans)
            {
                var trans = v.Item1;
                var go = v.Item2;

                var father = trans["m_Father"];
                long father_id = Convert.ToInt64(father["fileID"]);
                if (father_id == 0)
                {
                    scene.roots.Add(go);
                }
                else
                {
                    if (fileid_map.ContainsKey(father_id))
                    {
                        var trans_id = fileid_map[father_id];
                        if (trans2go.ContainsKey(trans_id))
                        {
                            var parentgo = trans2go[trans_id];
                            parentgo.childs.Add(go);
                        }
                        else
                        {
                            scene.roots.Add(go);
                        }
                    }
                    else
                    {
                        scene.roots.Add(go);
                    }
                }
            }

            foreach (var v in ref_gos)
            {
                var go = v.Value;
                var data = go.data;
                var modify = data["m_Modification"]["m_Modifications"];
                foreach (var m in modify)
                {
                    if (m["propertyPath"]?.ToString() == "m_Name")
                    {
                        go.name = m["value"]?.ToString() + "[REF]";
                    }
                }
                var father = data["m_Modification"]["m_TransformParent"];
                long fatherFileId = Convert.ToInt64(father["fileID"]);
                if (fileid_map.ContainsKey(fatherFileId))
                {
                    var tidx = fileid_map[fatherFileId];
                    if (trans2go.ContainsKey(tidx))
                    {
                        var parentgo = trans2go[tidx];
                        parentgo.childs.Add(go);
                    }
                    else
                    {
                        scene.roots.Add(go);
                    }
                }
                else
                {
                    scene.roots.Add(go);
                }
            }
            return scene;
        }
    }
}
