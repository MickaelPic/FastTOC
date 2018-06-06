using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

// <!-- TOC -->
// - [1. A](#1-A)
//     - [1.1 AA](#11-AA)
//         - [1.1.1 AAA](#111-AAA)
// <!-- /TOC -->
// # 1. A
// ## 1.1. AA
// ### 1.1.1. AAA

namespace FastTOC
{
  /// <summary>
  /// Command handler
  /// </summary>
  internal sealed class Command
  {
    /// <summary>
    /// Command ID.
    /// </summary>
    public const int CommandId = 0x0100;

    /// <summary>
    /// Command menu group (command set GUID).
    /// </summary>
    public static readonly Guid CommandSet = new Guid("01af178a-5980-47e8-932c-eff36f6259b2");

    /// <summary>
    /// VS Package that provides this command, not null.
    /// </summary>
    private readonly Package package;

    /// <summary>
    /// Initializes a new instance of the <see cref="Command"/> class.
    /// Adds our command handlers for menu (commands must exist in the command table file)
    /// </summary>
    /// <param name="package">Owner package, not null.</param>
    private Command(Package package)
    {
      if (package == null)
      {
        throw new ArgumentNullException("package");
      }

      this.package = package;

      OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
      if (commandService != null)
      {
        var menuCommandID = new CommandID(CommandSet, CommandId);
        var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
        commandService.AddCommand(menuItem);
      }
    }

    /// <summary>
    /// Gets the instance of the command.
    /// </summary>
    public static Command Instance
    {
      get;
      private set;
    }

    /// <summary>
    /// Gets the service provider from the owner package.
    /// </summary>
    private IServiceProvider ServiceProvider
    {
      get
      {
        return this.package;
      }
    }

    /// <summary>
    /// Initializes the singleton instance of the command.
    /// </summary>
    /// <param name="package">Owner package, not null.</param>
    public static void Initialize(Package package)
    {
      Instance = new Command(package);
    }

    /// <summary>
    /// This function is the callback used to execute the command when the menu item is clicked.
    /// See the constructor to see how the menu item is associated with this function using
    /// OleMenuCommandService service and MenuCommand class.
    /// </summary>
    /// <param name="sender">Event sender.</param>
    /// <param name="e">Event args.</param>
    private void MenuItemCallback(object sender, EventArgs e)
    {
      try
      {
        var dte = this.ServiceProvider.GetService(typeof(_DTE)) as _DTE;
        if (dte == null)
        {
          return;
        }

        var activeDocument = dte.ActiveDocument;
        if (activeDocument == null)
        {
          return;
        }

        var textDocument = activeDocument.Object() as TextDocument;
        if (textDocument == null)
        {
          return;
        }

        string editorContent = textDocument.CreateEditPoint(textDocument.StartPoint).GetText(textDocument.EndPoint);
        string[] lines = editorContent.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

        if (lines.Length == 0)
        {
          return;
        }

        // Search for the new TOC
        List<TOC_Entry> entries = new List<TOC_Entry>();
        List<TOC_Entry> currentDepth = new List<TOC_Entry>();
        int index;
        string depth;
        TOC_Entry current;
        bool hasChanged = false;
        foreach (string line in lines)
        {
          if (line.StartsWith("#") == false)
          {
            // Not a marker, don't care
            continue;
          }

          // Length of marker
          index = line.IndexOf(" ");

          if (index == 1)
          {
            // Root level
            depth = String.Format("{0}.", entries.Count + 1);

            current = new TOC_Entry(depth, line);
            entries.Add(current);
            currentDepth.Clear();
            currentDepth.Add(current);
          }
          else if (index > currentDepth.Count)
          {
            // Jumped to a sub level, check for potentially missing levels
            depth = String.Format("{0}", currentDepth[currentDepth.Count - 1].Depth);

            // Fill missing with 0
            for (int i = currentDepth.Count; i < index - 1; i++)
            {
              depth += "0.";
            }

            // Down level
            depth = String.Format("{0}{1}.", depth, currentDepth[currentDepth.Count - 1].Subs.Count + 1);
            current = new TOC_Entry(depth, line);
            currentDepth[currentDepth.Count - 1].Subs.Add(current);
            currentDepth.Add(current);
          }
          else
          {
            // Up from previous level
            currentDepth.RemoveRange(index - 1, currentDepth.Count - index + 1);
            depth = String.Format("{0}{1}.", currentDepth[currentDepth.Count - 1].Depth, currentDepth[currentDepth.Count - 1].Subs.Count + 1);
            current = new TOC_Entry(depth, line);
            currentDepth[currentDepth.Count - 1].Subs.Add(current);
            currentDepth.Add(current);
          }

          // Checks if the line has changed
          if (String.Compare(line, current.NewLine) != 0)
          {
            // Replace the line with its updated version
            if (textDocument.ReplacePattern(line + Environment.NewLine, current.NewLine + Environment.NewLine) == false)
            {
              textDocument.ReplacePattern(line, current.NewLine);

              hasChanged = true;
            }
          }
        }

        if (hasChanged)
        {
          // Generate the new TOC
          string newTOC = "";
          newTOC += "<!-- TOC -->";
          newTOC += Environment.NewLine;
          newTOC += GenerateNewTOC(entries);
          newTOC += "<!-- /TOC -->";
          newTOC += Environment.NewLine;

          // Search for the old TOC and replace it
          bool inOldTOC = false;
          bool foundOldTOC = false;
          foreach (string line in lines)
          {
            if (string.IsNullOrEmpty(line) == true)
            {
              continue;
            }
            if (line.IndexOf("<!-- TOC -->") >= 0)
            {
              // Beginning of Old TOC
              inOldTOC = true;
              foundOldTOC = true;
            }
            if (line.IndexOf("<!-- /TOC -->") >= 0)
            {
              // End of Old TOC
              inOldTOC = false;
              textDocument.ReplacePattern(line + Environment.NewLine, newTOC);
            }
            if (inOldTOC == true)
            {
              // Remove content of old TOC
              textDocument.ReplacePattern(line + Environment.NewLine, "");
            }
          }
          if (foundOldTOC == false)
          {
            // If no found old TOC, just place the new one at the top
            string line = lines[0];
            newTOC = newTOC + line + Environment.NewLine;
            textDocument.ReplacePattern(line + Environment.NewLine, newTOC);
          }
        }
      }
      catch (Exception ex)
      {
        VsShellUtilities.ShowMessageBox(
            this.ServiceProvider,
            String.Format("Failed to generate the new TOC.{0}System reports:{0}{1}", Environment.NewLine, ex.Message),
            "Error",
            OLEMSGICON.OLEMSGICON_INFO,
            OLEMSGBUTTON.OLEMSGBUTTON_OK,
            OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
      }
    }

    /// <summary>
    /// Generates the content of the new TOC recursively
    /// </summary>
    /// <param name="entries">List of entries to generate</param>
    /// <returns>A string containing the TOC generated for "entries"</returns>
    private string GenerateNewTOC(List<TOC_Entry> entries)
    {
      string newTOC = "";
      foreach (TOC_Entry entry in entries)
      {
        newTOC += entry.TOC_Line + Environment.NewLine;
        newTOC += GenerateNewTOC(entry.Subs);
      }
      return newTOC;
    }
  }
}