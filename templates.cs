using System.Collections.Generic;
using System.Linq;

public class TemplatesManager {

    public List<Template> List = new List<Template>();

    public string[] GetGroupNames() {
        return this.GetGroupDic().Keys.ToArray();
    }

    public Dictionary<string, List<Template>> GetGroupDic(string searchKeyword = "") {
        
        Dictionary<string, List<Template>> dic = new Dictionary<string, List<Template>>();

        foreach (Template temp in this.List) {
            if (!dic.ContainsKey(temp.Group)) dic.Add(temp.Group, new List<Template>());

            if (searchKeyword.Length == 0 || (searchKeyword.Length > 0 && (temp.Name.Contains(searchKeyword) || (temp.Category.Contains(searchKeyword)) || (temp.Body.Contains(searchKeyword))))) {
                dic[temp.Group].Add(temp);
            }
        }

        return dic;
    }
}

public class Template {
    public string Category { get; set; }
    public string Name { get; set; }
    public string Body { get; set; }
    public string Group { get; set; } 

    public Template(string group, string category, string name, string body) {
        this.Group = group;
        this.Category = category;
        this.Name = name;
        this.Body = body;
    }
}