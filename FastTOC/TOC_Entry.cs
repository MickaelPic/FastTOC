using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace FastTOC
{
  public class TOC_Entry
  {
    private string _title = "";

    private string _line = "";

    private string _link = "";
    public string Depth { get; set; }

    public string NewLine { get; set; }

    public List<TOC_Entry> Subs { get; set; }

    public string TOC_Line { get; set; }

    public TOC_Entry(string depth, string line)
    {
      Depth = depth;
      this._line = line;
      this._title = ExtractTitle(line, line.IndexOf(" "));
      Subs = new List<TOC_Entry>();

      GenerateLink();
      Generate_ToC_Line();
      GenerateNewLine();
    }

    public override string ToString()
    {
      return String.Format("{0} {1}", Depth, this._title);
    }

    private string ExtractTitle(
      string line,
      int index)
    {
      line = line.Substring(index + 1);
      Regex r = new Regex("[0-9]*\\.");
      Match m = r.Match(line);
      while (m.Success)
      {
        line = line.Replace(m.Value, "");
        m = m.NextMatch();
      }
      line = line.Trim();

      return line;
    }

    private void GenerateLink()
    {
      string text = String.Format("{0} {1}", Depth, this._title.ToLower());
      text = text.Replace(" ", "-");
      StringBuilder sb = new StringBuilder();
      foreach(char c in text)
      {
        if ((c >= '0' && c <= '9') ||
          (c >= 'a' && c <= 'z') ||
          c == '-')
        {
          sb.Append(c);
        }
      }

      this._link = "#" + sb.ToString();
    }

    private void Generate_ToC_Line()
    {
      string spaces = "";
      string[] splitted = Depth.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
      for (int i = 0; i < splitted.Length - 1; i ++)
      {
        spaces += "    ";
      }
      TOC_Line = String.Format("{0}- [{1} {2}]({3})", spaces, Depth, this._title, this._link);
    }

    private void GenerateNewLine()
    {
      string marks = "";
      int count = Depth.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries).Length;
      for (int i = 0; i < count; i ++)
      {
        marks += "#";
      }
      this.NewLine = String.Format("{0} {1} {2}", marks, Depth, this._title);
    }
  }
}