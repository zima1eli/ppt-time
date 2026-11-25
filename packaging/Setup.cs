using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

[assembly: AssemblyTitle("ppt time Setup")]
[assembly: AssemblyProduct("ppt time")]

internal static class Setup
{
    private const string Product = "ppt time";

    private static void ExtractMsi(string target)
    {
        using (Stream input = Assembly.GetExecutingAssembly().GetManifestResourceStream("ppt-time.msi"))
        using (FileStream output = new FileStream(target, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            input.CopyTo(output);
    }

    private static bool ConfirmInstallation(out bool enablePowerPoint, out bool enableWps)
    {
        using (Form form = new Form())
        using (TableLayoutPanel panel = new TableLayoutPanel())
        {
            form.Text = Product;
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MaximizeBox = form.MinimizeBox = false;
            form.AutoSize = true;
            form.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            panel.AutoSize = true;
            panel.ColumnCount = 1;
            panel.Padding = new Padding(14);
            var description = new Label { AutoSize = true, Text = "安装到当前用户，并添加开始菜单快捷方式。" };
            var powerPoint = new CheckBox { AutoSize = true, Checked = true, Text = "启用 PowerPoint 页码联动" };
            var wps = new CheckBox { AutoSize = true, Checked = true, Text = "启用 WPS 演示页码联动" };
            var hint = new Label { AutoSize = true, Text = "首次启用后，请保存文稿并重新打开对应的演示软件。" };
            var buttons = new FlowLayoutPanel { AutoSize = true, Anchor = AnchorStyles.Right };
            var install = new Button { Text = "安装", DialogResult = DialogResult.OK, AutoSize = true };
            var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, AutoSize = true };
            buttons.Controls.Add(install);
            buttons.Controls.Add(cancel);
            panel.Controls.Add(description);
            panel.Controls.Add(powerPoint);
            panel.Controls.Add(wps);
            panel.Controls.Add(hint);
            panel.Controls.Add(buttons);
            form.Controls.Add(panel);
            form.AcceptButton = install;
            form.CancelButton = cancel;
            bool confirmed = form.ShowDialog() == DialogResult.OK;
            enablePowerPoint = confirmed && powerPoint.Checked;
            enableWps = confirmed && wps.Checked;
            return confirmed;
        }
    }

    [STAThread]
    private static int Main(string[] args)
    {
        bool quiet = args.Length >= 1 && args[0].Equals("/quiet", StringComparison.OrdinalIgnoreCase);
        bool enablePowerPoint = false, enableWps = false;
        bool extract = args.Length == 2 && args[0].Equals("/extract", StringComparison.OrdinalIgnoreCase);
        string temporary = null;
        try
        {
            if (extract)
            {
                ExtractMsi(Path.GetFullPath(args[1]));
                return 0;
            }
            if (args.Length != 0 && !quiet)
                return 87;
            for (int i = 1; i < args.Length; i++)
            {
                if (args[i].Equals("/enable-powerpoint", StringComparison.OrdinalIgnoreCase))
                    enablePowerPoint = true;
                else if (args[i].Equals("/enable-wps", StringComparison.OrdinalIgnoreCase))
                    enableWps = true;
                else
                    return 87;
            }
            if (!quiet)
            {
                Application.EnableVisualStyles();
                if (!ConfirmInstallation(out enablePowerPoint, out enableWps))
                    return 1602;
            }
            temporary = Path.Combine(Path.GetTempPath(), "ppt-time-" + Guid.NewGuid().ToString("N") + ".msi");
            ExtractMsi(temporary);
            string engine = Path.Combine(Environment.SystemDirectory, "msiexec.exe");
            string options = "/i \"" + temporary + "\" " + (quiet ? "/qn" : "/qb") + " /norestart";
            if (enablePowerPoint)
                options += " PPTTIME_ENABLE_POWERPOINT=1";
            if (enableWps)
                options += " PPTTIME_ENABLE_WPS=1";
            else if (!quiet)
                options += " PPTTIME_ENABLE_WPS=0";
            using (Process process = Process.Start(new ProcessStartInfo(engine, options) { UseShellExecute = false }))
            {
                process.WaitForExit();
                int result = process.ExitCode;
                if (!quiet && (result == 0 || result == 3010))
                    MessageBox.Show("安装完成。可从 Windows 开始菜单打开 ppt time。", Product);
                else if (!quiet && result != 1602)
                    MessageBox.Show("安装未完成。Windows Installer 返回 " + result.ToString() + "。", Product,
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return result;
            }
        }
        catch (Exception error)
        {
            if (!quiet && !extract)
                MessageBox.Show(error.Message, Product, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1603;
        }
        finally
        {
            if (temporary != null)
            {
                try { File.Delete(temporary); }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }
    }
}
