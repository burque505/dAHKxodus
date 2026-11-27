// Printing pages whose paper is not the printer's default (an envelope, a document of mixed sizes): Word sends each
// section's paper size with the job, so an envelope prints on its envelope form without the printer being changed.
// The browser's print dialog takes one paper size for the whole job, its own, so a document whose pages are not all
// the default paper prints from here: a small dialog (printer, copies, pages), then a job for each run of pages of
// one size, each with that paper size (WebView2's print settings: the driver picks its form of that size). Loaded by
// dAHKxodus.cs (#load).
using System.Drawing;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public static class PrintPages {
    /// The pages' sizes in points, from the page's "612x792,684x297".
    public static List<(double w, double h)> Parse(string sizes) {
        var list = new List<(double, double)>();
        foreach (var s in sizes.Split(',', StringSplitOptions.RemoveEmptyEntries)) {
            var wh = s.Split('x');
            if (wh.Length == 2 && double.TryParse(wh[0], System.Globalization.CultureInfo.InvariantCulture, out var w)
                && double.TryParse(wh[1], System.Globalization.CultureInfo.InvariantCulture, out var h)) list.Add((w, h));
        }
        return list;
    }

    static bool Same(double a, double b) => Math.Abs(a - b) <= 3;

    /// Whether a page is a sheet of this paper (points), either way round (a landscape page prints on it turned).
    public static bool Fits((double w, double h) page, (double w, double h) paper) =>
        (Same(page.w, paper.w) && Same(page.h, paper.h)) || (Same(page.w, paper.h) && Same(page.h, paper.w));

    /// The default printer's default paper in points; null with no printer.
    public static (double w, double h)? DefaultPaper() {
        try {
            var settings = new PrinterSettings();
            if (!settings.IsValid) return null;
            var paper = settings.DefaultPageSettings.PaperSize;
            return (paper.Width * 0.72, paper.Height * 0.72);
        } catch (Exception) { return null; }
    }

    /// Whether every page is the default printer's paper: then the browser's print dialog prints them right.
    public static bool FitDefaultPaper(List<(double w, double h)> pages) =>
        DefaultPaper() is not { } paper || pages.All(p => Fits(p, paper));

    /// The pages from `first` to `last` (1-based) in runs of one size: each run's first and last page and its size.
    public static List<(int first, int last, double w, double h)> Runs(List<(double w, double h)> pages, int first, int last) {
        var runs = new List<(int, int, double, double)>();
        for (var i = Math.Max(1, first); i <= Math.Min(last, pages.Count); i++) {
            var p = pages[i - 1];
            if (runs.Count > 0 && runs[^1].Item2 == i - 1 && Same(runs[^1].Item3, p.w) && Same(runs[^1].Item4, p.h))
                runs[^1] = (runs[^1].Item1, i, runs[^1].Item3, runs[^1].Item4);
            else runs.Add((i, i, p.w, p.h));
        }
        return runs;
    }

    /// The printer's own paper form of a size (points, either way round): its paper kind (DMPAPER_*, e.g. 20 for
    /// Envelope #10) and name; null when it has none (then a custom size is all there is to ask for).
    public static (short kind, string name)? FormOf(string printer, (double w, double h) size) {
        try {
            var settings = new PrinterSettings { PrinterName = printer };
            if (!settings.IsValid) return null;
            double shortSide = Math.Min(size.w, size.h) / 0.72, longSide = Math.Max(size.w, size.h) / 0.72;    // (hundredths of an inch)
            foreach (PaperSize paper in settings.PaperSizes) {
                if (paper.Kind == PaperKind.Custom || paper.RawKind <= 0 || paper.RawKind > short.MaxValue) continue;
                double a = Math.Min(paper.Width, paper.Height), b = Math.Max(paper.Width, paper.Height);
                if (Math.Abs(a - shortSide) <= 8 && Math.Abs(b - longSide) <= 8) return ((short)paper.RawKind, paper.PaperName);
            }
        } catch (Exception) { }
        return null;
    }

    // ---- The printer's paper for this user, for one job (what choosing it in the printer's settings does by hand)

    [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool OpenPrinter(string name, out IntPtr printer, IntPtr defaults);
    [DllImport("winspool.drv", SetLastError = true)]
    static extern bool ClosePrinter(IntPtr printer);
    [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool GetPrinter(IntPtr printer, int level, IntPtr buffer, int size, out int needed);
    [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool SetPrinter(IntPtr printer, int level, IntPtr buffer, int command);
    [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern int DocumentProperties(IntPtr window, IntPtr printer, string name, IntPtr output, IntPtr input, int mode);

    const int DM_OUT_BUFFER = 2, DM_IN_BUFFER = 8;
    const int DM_ORIENTATION = 0x1, DM_PAPERSIZE = 0x2;
    // DEVMODEW: dmDeviceName (64 bytes), dmSpecVersion, dmDriverVersion, dmSize, dmDriverExtra, dmFields (at 72), then
    // dmOrientation (76), dmPaperSize (78)
    const int OffFields = 72, OffOrientation = 76, OffPaperSize = 78;

    /// Make paper `kind` (turned landscape or not) this user's default on `printer` until the returned value is
    /// disposed, which puts the user's own default back. Null when the printer can't be set so.
    public static IDisposable? UsePaper(string printer, short kind, bool landscape) {
        if (!OpenPrinter(printer, out var handle, IntPtr.Zero)) { MainForm.Log($"print: OpenPrinter {printer} failed ({Marshal.GetLastWin32Error()})"); return null; }
        IntPtr saved = IntPtr.Zero, devmode = IntPtr.Zero, info = IntPtr.Zero;
        try {
            // the user's own default now (PRINTER_INFO_9: a DEVMODE, or none), to put back
            GetPrinter(handle, 9, IntPtr.Zero, 0, out var needed);
            if (needed > 0) {
                saved = Marshal.AllocHGlobal(needed);
                if (!GetPrinter(handle, 9, saved, needed, out _)) { Marshal.FreeHGlobal(saved); saved = IntPtr.Zero; }
            }
            IntPtr savedMode = saved == IntPtr.Zero ? IntPtr.Zero : Marshal.ReadIntPtr(saved);
            byte[]? savedCopy = null;
            if (savedMode != IntPtr.Zero) {
                var size = Marshal.ReadInt16(savedMode, 68) + Marshal.ReadInt16(savedMode, 70);
                savedCopy = new byte[size];
                Marshal.Copy(savedMode, savedCopy, 0, size);
            }
            // the settings in effect, with the paper changed, checked by the driver
            var length = DocumentProperties(IntPtr.Zero, handle, printer, IntPtr.Zero, IntPtr.Zero, 0);
            if (length <= 0) return null;
            devmode = Marshal.AllocHGlobal(length);
            if (DocumentProperties(IntPtr.Zero, handle, printer, devmode, IntPtr.Zero, DM_OUT_BUFFER) < 0) return null;
            Marshal.WriteInt32(devmode, OffFields, Marshal.ReadInt32(devmode, OffFields) | DM_PAPERSIZE | DM_ORIENTATION);
            Marshal.WriteInt16(devmode, OffPaperSize, kind);
            Marshal.WriteInt16(devmode, OffOrientation, (short)(landscape ? 2 : 1));
            if (DocumentProperties(IntPtr.Zero, handle, printer, devmode, devmode, DM_IN_BUFFER | DM_OUT_BUFFER) < 0) return null;
            info = Marshal.AllocHGlobal(IntPtr.Size);
            Marshal.WriteIntPtr(info, devmode);
            if (!SetPrinter(handle, 9, info, 0)) { MainForm.Log($"print: SetPrinter {printer} failed ({Marshal.GetLastWin32Error()})"); return null; }
            var keep = (handle, savedCopy, devmode, info);
            handle = IntPtr.Zero; devmode = IntPtr.Zero; info = IntPtr.Zero;    // (the restorer owns them now)
            return new Restore(() => {
                var (h, copy, mode, block) = keep;
                try {
                    IntPtr back = IntPtr.Zero;
                    if (copy != null) { back = Marshal.AllocHGlobal(copy.Length); Marshal.Copy(copy, 0, back, copy.Length); }
                    Marshal.WriteIntPtr(block, back);
                    if (!SetPrinter(h, 9, block, 0)) MainForm.Log($"print: putting {printer}'s paper back failed ({Marshal.GetLastWin32Error()})");
                    if (back != IntPtr.Zero) Marshal.FreeHGlobal(back);
                } finally {
                    Marshal.FreeHGlobal(mode);
                    Marshal.FreeHGlobal(block);
                    ClosePrinter(h);
                }
            });
        } finally {
            if (saved != IntPtr.Zero) Marshal.FreeHGlobal(saved);
            if (devmode != IntPtr.Zero) Marshal.FreeHGlobal(devmode);
            if (info != IntPtr.Zero) Marshal.FreeHGlobal(info);
            if (handle != IntPtr.Zero) ClosePrinter(handle);
        }
    }

    sealed class Restore(Action undo) : IDisposable {
        Action? undo = undo;
        public void Dispose() { undo?.Invoke(); undo = null; }
    }

    /// A size as the dialog names it: "9.5 x 4.13 in".
    public static string Name((double w, double h) size) =>
        $"{Math.Round(size.w / 72, 2)} x {Math.Round(size.h / 72, 2)} in";
}

/// The print dialog for pages of their own paper sizes: the printer, copies and pages, and which sizes the pages are.
public sealed class PrintPagesDialog : Form {
    readonly ComboBox printer = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 320 };
    readonly NumericUpDown copies = new() { Minimum = 1, Maximum = 99, Value = 1, Width = 60 };
    readonly RadioButton all = new() { AutoSize = true, Checked = true };
    readonly RadioButton some = new() { AutoSize = true };
    readonly NumericUpDown from = new() { Minimum = 1, Width = 60 }, to = new() { Minimum = 1, Width = 60 };

    public string Printer => printer.SelectedItem as string ?? "";
    public int Copies => (int)copies.Value;
    public int First => all.Checked ? 1 : (int)from.Value;
    public int Last => all.Checked ? int.MaxValue : (int)to.Value;

    public PrintPagesDialog(string documentName, List<(double w, double h)> pages) {
        Text = Lang.T("printPages.title", documentName);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MinimizeBox = MaximizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Padding = new Padding(12);
        var defaultName = new PrinterSettings().PrinterName;
        foreach (string name in PrinterSettings.InstalledPrinters) printer.Items.Add(name);
        printer.SelectedItem = printer.Items.Contains(defaultName) ? defaultName : printer.Items.Count > 0 ? printer.Items[0] : null;
        from.Maximum = to.Maximum = Math.Max(1, pages.Count);
        to.Value = Math.Max(1, pages.Count);
        all.Text = Lang.T("printPages.all", pages.Count);
        some.Text = Lang.T("printPages.from");
        from.ValueChanged += (s, e) => some.Checked = true;
        to.ValueChanged += (s, e) => some.Checked = true;

        // which pages are which size: "Pages 1: 9.5 x 4.13 in; pages 2-3: 8.5 x 11 in"
        var sizes = string.Join("\n", PrintPages.Runs(pages, 1, pages.Count).Select(r =>
            (r.first == r.last ? Lang.T("printPages.page", r.first) : Lang.T("printPages.pages", r.first, r.last)) + ": " + PrintPages.Name((r.w, r.h))));
        var note = new Label { AutoSize = true, MaximumSize = new Size(440, 0), Text = Lang.T("printPages.note") + "\n" + sizes, Margin = new Padding(0, 0, 0, 10) };

        var grid = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Fill };
        Label L(string key) => new() { Text = Lang.T(key), AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 8, 0) };
        grid.Controls.Add(note, 0, 0);
        grid.SetColumnSpan(note, 2);
        grid.Controls.Add(L("printPages.printer"), 0, 1);
        grid.Controls.Add(printer, 1, 1);
        grid.Controls.Add(L("printPages.copies"), 0, 2);
        grid.Controls.Add(copies, 1, 2);
        var range = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = Padding.Empty };
        range.Controls.AddRange([all, some, from, new Label { Text = "-", AutoSize = true, Margin = new Padding(4, 6, 4, 0) }, to]);
        grid.Controls.Add(L("printPages.range"), 0, 3);
        grid.Controls.Add(range, 1, 3);
        var ok = new Button { Text = Lang.T("printPages.print"), DialogResult = DialogResult.OK, AutoSize = true, MinimumSize = new Size(88, 0) };
        var cancel = new Button { Text = Lang.T("settings.cancel"), DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(88, 0) };
        var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 10, 0, 0) };
        buttons.Controls.AddRange([cancel, ok]);
        grid.Controls.Add(buttons, 0, 4);
        grid.SetColumnSpan(buttons, 2);
        Controls.Add(grid);
        AcceptButton = ok;
        CancelButton = cancel;
        ok.Enabled = printer.Items.Count > 0;
        ok.Click += (s, e) => { if (some.Checked && from.Value > to.Value) { DialogResult = DialogResult.None; to.Focus(); } };
    }
}
