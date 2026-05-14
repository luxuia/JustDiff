using VYaml;
using VYaml.Annotations;
using VYaml.Serialization;
using System.Text.RegularExpressions;
using UnityYamlParser;


class Program
{
    static void ParseUnity(string path)
    {
        Console.WriteLine("Begin " + path);
        Console.WriteLine("----------- ");

        var scene = ParseUnityYaml.ParseYaml(path);

        Action<GameObject, string> Dump = null;
        Dump = (GameObject go, string indent) =>
        {
            Console.WriteLine(indent + go.name + "=>" + string.Join("|", go.comps));

            foreach (var child in go.childs)
            {
                Dump(child, indent + "  ");
            }
        };
        foreach (var go in scene.roots)
        {
            Dump(go, "");
        }

        Console.WriteLine("----------- ");
        Console.WriteLine();
    }


    static void Main(string[] args)
    {
        ParseUnity("../../../test/Cube.prefab");
        ParseUnity("../../../test/SampleScene.unity");
    }
}