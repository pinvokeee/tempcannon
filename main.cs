using System;
using System.Windows.Forms;
using Exloader;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.IO;
using System.Runtime.InteropServices;

class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        string excelPath = args.Length > 0 ? args[0] : "テンプレート.xlsx";
        bool useClipboard = (args.Length > 1) && args[1].ToUpper() == "-USECLIPBOARD";
        Application.EnableVisualStyles();
        Application.Run(new Form1(excelPath, useClipboard));

    }
}

class Form1 : Form
{
    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int x, int y, int cx, int cy, uint flags);
    [DllImport("user32.dll")]
    private static extern bool SetActiveWindow(IntPtr hWnd);

    private const int HWND_TOPMOST = -1;
    private enum SWP : int
    {
        NOSIZE = 0x0001,
        NOMOVE = 0x0002,
        NOZORDER = 0x0004,
        NOREDRAW = 0x0008,
        NOACTIVATE = 0x0010,
        FRAMECHANGED = 0x0020,
        SHOWWINDOW = 0x0040,
        HIDEWINDOW = 0x0080,
        NOCOPYBITS = 0x0100,
        NOOWNERZORDER = 0x0200,
        NOSENDCHANGING = 0x400
    }


    [DllImport("user32.dll")]
    private static extern UInt32 GetWindowLong(IntPtr hWnd,GWL index);
    [DllImport("user32.dll")]
    private static extern UInt32 SetWindowLong(IntPtr hWnd,GWL index, UInt32 unValue);

    const UInt32 WS_EX_NOACTIVATE = 0x8000000;
    private enum GWL : int
    {
        WINDPROC = -4,
        HINSTANCE = -6,
        HWNDPARENT = -8,
        ID = -12,
        STYLE = -16,
        EXSTYLE = -20,
        USERDATA = -21,
    }

    private void SetNoActive(IntPtr hWnd)
    {
        UInt32 unStyle = GetWindowLong(hWnd, GWL.EXSTYLE);
        unStyle = (unStyle | WS_EX_NOACTIVATE);
        SetWindowLong(hWnd, GWL.EXSTYLE, unStyle);
    }

    private void SetActive(IntPtr hWnd)
    {
        UInt32 unStyle = GetWindowLong(hWnd, GWL.EXSTYLE);
        unStyle = (unStyle);
        SetWindowLong(hWnd, GWL.EXSTYLE, unStyle);
    }

    public string ExcelPath { get;set; }
    public bool UseClipboard { get; set; }

    public Form1(string path, bool useClipboard)
    {
        this.Text = "TextCannon";
        this.TopMost = true;
        this.Load += new EventHandler(Form1_Load);
        this.ShowInTaskbar = false;
        this.ShowIcon = false;
        this.MaximizeBox = false;

        if (!File.Exists(path)) {
            MessageBox.Show("テンプレートファイルが存在しません\n終了します");
            Environment.Exit(0);
        }

        this.ExcelPath = path;
        this.UseClipboard = useClipboard;
    }

    public TemplatesManager Templates { get; set; }

    private TemplatesManager LoadTemplatesData() {
        
        TemplatesManager templates = new TemplatesManager();

        string path = this.ExcelPath; 

        using (ExcelLoader excel = new ExcelLoader()){

            using (Workbook w = excel.OpenWorkbook(path)) {
                
                WorkSheet[] sheets = w.GetWorkSheets();

                foreach (WorkSheet sheet in sheets) {

                    using (sheet) {

                        object[,] range = sheet.GetUsedRange();

                        //ヘッダーは無視
                        for (int y = 1; y < range.GetLength(0); y++) {
                            
                            string category = range[y, 0].ToString();
                            string name = range[y, 1].ToString();
                            string body = range[y, 2].ToString();

                            templates.List.Add(new Template(sheet.Name, category, name, body));
                        }
                    }
                }
            }
        }

        return templates;
    }

    private Panel TempSelecter = new Panel();
    private Panel GroupSelecter = new Panel();

    private string SelectedGroup = "";
    private string SearchKeyword { get; set; }

    private TextBox SearchTextBox { get; set; }

    private void Form1_Load(object sender, System.EventArgs e)
    {
        // SetWindowPos(this.Handle, HWND_TOPMOST,0, 0, 0, 0,(uint)(SWP.NOMOVE | SWP.NOSIZE |SWP.NOOWNERZORDER | SWP.FRAMECHANGED |SWP.NOSENDCHANGING | SWP.NOACTIVATE |SWP.SHOWWINDOW));
        SetNoActive(this.Handle);

       this.Templates = this.LoadTemplatesData();

        Panel p = new Panel ();
        p.Dock = DockStyle.Top;
        
        this.SearchTextBox = new TextBox();
        this.SearchTextBox.Dock = DockStyle.Fill;
        this.SearchTextBox.GotFocus += new EventHandler(SearchTextBox_GotFocus);
        this.SearchTextBox.LostFocus += new EventHandler(SearchTextBox_LostFocus);

        p.Height = this.SearchTextBox.Height;
        // p.Controls.Add(this.SearchTextBox);
        

        TempSelecter.Dock = DockStyle.Fill;
        TempSelecter.AutoScroll = true;

        GroupSelecter.Width = 100;
        GroupSelecter.Dock = DockStyle.Right;
        GroupSelecter.Padding = new Padding(3);
        GroupSelecter.AutoScroll = true;
        // GroupSelecter.BackColor  = Color.Red;


        this.Controls.Add(TempSelecter);
        this.Controls.Add(GroupSelecter);
        this.Controls.Add(p);

        this.SearchKeyword = "";

        this.AppendTemplateButtons();
    }

    private void SearchTextBox_GotFocus(object sender, EventArgs e)
    {
        SetActive(this.Handle);
    }

    private void SearchTextBox_LostFocus(object sender, EventArgs e)
    {
        SetNoActive(this.Handle);
    }

    private void AppendTemplateButtons() {
        
        this.GroupSelecter.Controls.Clear();

        Dictionary<string, List<Template>> dic = this.Templates.GetGroupDic(this.SearchKeyword);
        string[] keys = dic.Keys.ToArray();
        Array.Reverse(keys);

        RadioButton groupButton = null;

        foreach (string key in keys) {
            
            groupButton = new RadioButton();
            groupButton.Text = key;
            groupButton.Dock = DockStyle.Top;
            groupButton.Appearance = Appearance.Button;
            groupButton.Padding = new Padding(7);
            groupButton.Height = 35;
            groupButton.TextAlign = ContentAlignment.MiddleCenter;
            groupButton.CheckedChanged += new EventHandler(GroupButton_CheckedChanged);

            // Panel panel = new Panel();
            // panel.Dock = DockStyle.Top;
            // panel.Padding = new Padding(2);
            // panel.Height = 40;
            // panel.Controls.Add(groupButton);

            this.GroupSelecter.Controls.Add(groupButton);
        }

        if (groupButton != null) {
            groupButton.PerformClick();
        }

        if (keys.Length > 0) {
            this.SelectedGroup = keys[keys.Length - 1];
            Template[] temps = dic[this.SelectedGroup].ToArray();
            // ApplyTemplatesButton(temps);
        }

    }  

    private void ApplyTemplatesButton(Template[] templates) {

        Template[] temps = templates;
        Array.Reverse(temps);

        this.TempSelecter.Controls.Clear();

        List<DBPanel> buttons = new List<DBPanel>();

        Button tempButton = null;

        foreach (Template temp in temps) {

            tempButton = new Button();
            tempButton.Text = temp.Category + "_" + temp.Name;
            tempButton.Dock = DockStyle.Fill;
            tempButton.Tag = temp;
            tempButton.Click += (object sender, System.EventArgs e) => {

                Button b = (Button)(sender);
                string message = ((Template)(b.Tag)).Body;
                
                if (!this.UseClipboard) {
                    string escaped =  Regex.Replace(message, @"[+^%~(){}\[\]]", "{$0}");
                    SendKeys.Send(escaped);
                }
                else {
                    Clipboard.SetText(message);
                    SendKeys.Send("^v");
                }

            };

            DBPanel bpanel = new DBPanel();
            bpanel.Dock = DockStyle.Top;
            bpanel.Padding = new Padding(2);
            bpanel.Height = 40;
            bpanel.Controls.Add(tempButton);

            buttons.Add(bpanel);
        }

        this.TempSelecter.Controls.AddRange(buttons.ToArray());

        if (tempButton != null) { 
            this.ActiveControl = tempButton;
            tempButton.Focus();
        }
    }

    public void GroupButton_CheckedChanged(object sender, System.EventArgs e) {
        string key = ((RadioButton)sender).Text;
        Dictionary<string, List<Template>> dic = this.Templates.GetGroupDic(this.SearchKeyword);
        this.ApplyTemplatesButton(dic[key].ToArray());
    }

    // private const int WS_EX_NOACTIVATE = 0x8000000;
    // protected override CreateParams CreateParams
    // {
    //     get
    //     {
    //         Console.WriteLine((SearchTextBox != null && this.SearchTextBox.Focused));
    //         CreateParams p = base.CreateParams;

    //         if (!DesignMode || (SearchTextBox != null && this.SearchTextBox.Focused))
    //         {
    //             p.ExStyle |= (WS_EX_NOACTIVATE);
    //         }

    //         return (p);
    //     }
    // }


}

public class DBPanel : Panel
{
    public DBPanel()
    {
        this.DoubleBuffered = true;
    }
}
