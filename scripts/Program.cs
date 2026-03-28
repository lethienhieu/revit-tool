using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;

class Program
{
    static string outputDir;

    static void Main()
    {
        outputDir = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "THBIM_Core", "Resources", "MEP"));
        if (!Directory.Exists(outputDir))
            Directory.CreateDirectory(outputDir);

        Console.WriteLine($"Output: {outputDir}");

        // Bloom 16 + 32
        SaveIcon("Bloom", 16, DrawBloom);
        SaveIcon("Bloom", 32, DrawBloom);

        // Route 16 + 32
        SaveIcon("Route", 16, DrawRoute);
        SaveIcon("Route", 32, DrawRoute);

        // Reroute 16
        SaveIcon("Reroute", 16, DrawReroute);

        // Elbow variants 16
        SaveIcon("ElbowUp45", 16, DrawElbowUp45);
        SaveIcon("ElbowDown45", 16, DrawElbowDown45);
        SaveIcon("ElbowUp90", 16, DrawElbowUp90);
        SaveIcon("ElbowDown90", 16, DrawElbowDown90);

        // Connect, Disconnect, AlignBranch, MoveConnect 16
        SaveIcon("Connect", 16, DrawConnect);
        SaveIcon("Disconnect", 16, DrawDisconnect);
        SaveIcon("AlignBranch", 16, DrawAlignBranch);
        SaveIcon("MoveConnect", 16, DrawMoveConnect);

        // ParaPusher 16 + 32
        SaveIcon("ParaPusher", 16, DrawParaPusher);
        SaveIcon("ParaPusher", 32, DrawParaPusher);

        // SumParam 16
        SaveIcon("SumParam", 16, DrawSumParam);

        Console.WriteLine("Done! All icons generated.");
    }

    static void SaveIcon(string name, int size, Action<Graphics, int> draw)
    {
        using var bmp = new Bitmap(size, size);
        using var g = Graphics.FromImage(bmp);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
        g.Clear(Color.Transparent);
        draw(g, size);
        string path = Path.Combine(outputDir, $"{name}_{size}.png");
        bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
        Console.WriteLine($"  {name}_{size}.png");
    }

    static void DrawBloom(Graphics g, int s)
    {
        int cx = s / 2, cy = s / 2, r = s / 6;
        float pw = Math.Max(1.5f, s / 10f);
        using var pen = new Pen(Color.FromArgb(33, 150, 243), pw);
        using var brush = new SolidBrush(Color.FromArgb(33, 150, 243));
        g.FillEllipse(brush, cx - r, cy - r, r * 2, r * 2);
        for (int i = 0; i < 6; i++)
        {
            double a = i * Math.PI / 3;
            float x1 = cx + (float)(r * 1.4 * Math.Cos(a));
            float y1 = cy + (float)(r * 1.4 * Math.Sin(a));
            float x2 = cx + (float)(r * 2.8 * Math.Cos(a));
            float y2 = cy + (float)(r * 2.8 * Math.Sin(a));
            g.DrawLine(pen, x1, y1, x2, y2);
        }
    }

    static void DrawRoute(Graphics g, int s)
    {
        float pw = Math.Max(2f, s / 8f);
        using var pen = new Pen(Color.FromArgb(76, 175, 80), pw);
        pen.StartCap = LineCap.Round;
        pen.EndCap = LineCap.ArrowAnchor;
        float m = s / 6f;
        g.DrawLine(pen, m, s - m, s / 2f, s / 2f);
        g.DrawLine(pen, s / 2f, s / 2f, s - m, m);
    }

    static void DrawReroute(Graphics g, int s)
    {
        float pw = Math.Max(1.5f, s / 8f);
        using var pen = new Pen(Color.FromArgb(255, 152, 0), pw);
        pen.StartCap = LineCap.Round; pen.EndCap = LineCap.Round;
        float m = s / 5f;
        var pts = new PointF[] {
            new(m, s/2f), new(s/3f, s/2f),
            new(s/3f, m), new(2*s/3f, m),
            new(2*s/3f, s/2f), new(s-m, s/2f)
        };
        g.DrawLines(pen, pts);
    }

    static void DrawElbowUp45(Graphics g, int s)
    {
        float pw = Math.Max(1.5f, s / 8f);
        using var pen = new Pen(Color.FromArgb(33, 150, 243), pw);
        pen.EndCap = LineCap.ArrowAnchor;
        float m = s / 5f;
        g.DrawLine(pen, m, s - m, s / 2f, s / 2f);
        g.DrawLine(pen, s / 2f, s / 2f, s - m, m);
    }

    static void DrawElbowDown45(Graphics g, int s)
    {
        float pw = Math.Max(1.5f, s / 8f);
        using var pen = new Pen(Color.FromArgb(33, 150, 243), pw);
        pen.EndCap = LineCap.ArrowAnchor;
        float m = s / 5f;
        g.DrawLine(pen, m, m, s / 2f, s / 2f);
        g.DrawLine(pen, s / 2f, s / 2f, s - m, s - m);
    }

    static void DrawElbowUp90(Graphics g, int s)
    {
        float pw = Math.Max(1.5f, s / 8f);
        using var pen = new Pen(Color.FromArgb(0, 150, 136), pw);
        pen.EndCap = LineCap.ArrowAnchor;
        float m = s / 5f;
        g.DrawLine(pen, m, s - m, m, m);
        g.DrawLine(pen, m, m, s - m, m);
    }

    static void DrawElbowDown90(Graphics g, int s)
    {
        float pw = Math.Max(1.5f, s / 8f);
        using var pen = new Pen(Color.FromArgb(0, 150, 136), pw);
        pen.EndCap = LineCap.ArrowAnchor;
        float m = s / 5f;
        g.DrawLine(pen, m, m, m, s - m);
        g.DrawLine(pen, m, s - m, s - m, s - m);
    }

    static void DrawConnect(Graphics g, int s)
    {
        float pw = Math.Max(1.5f, s / 8f);
        using var pen = new Pen(Color.FromArgb(76, 175, 80), pw);
        float m = s / 5f;
        g.DrawLine(pen, m, s / 2f, s / 2f - 2, s / 2f);
        g.DrawLine(pen, s / 2f + 2, s / 2f, s - m, s / 2f);
        using var brush = new SolidBrush(Color.FromArgb(76, 175, 80));
        float dr = Math.Max(2, s / 6f);
        g.FillEllipse(brush, s / 2f - dr / 2, s / 2f - dr / 2, dr, dr);
    }

    static void DrawDisconnect(Graphics g, int s)
    {
        float pw = Math.Max(1.5f, s / 8f);
        using var pen = new Pen(Color.FromArgb(244, 67, 54), pw);
        float m = s / 5f;
        g.DrawLine(pen, m, s / 2f, s / 2f - 3, s / 2f);
        g.DrawLine(pen, s / 2f + 3, s / 2f, s - m, s / 2f);
        float xr = s / 7f;
        using var xPen = new Pen(Color.FromArgb(244, 67, 54), Math.Max(1f, s / 12f));
        g.DrawLine(xPen, s / 2f - xr, s / 2f - xr, s / 2f + xr, s / 2f + xr);
        g.DrawLine(xPen, s / 2f + xr, s / 2f - xr, s / 2f - xr, s / 2f + xr);
    }

    static void DrawAlignBranch(Graphics g, int s)
    {
        float pw = Math.Max(1.5f, s / 8f);
        using var pen = new Pen(Color.FromArgb(156, 39, 176), pw);
        float m = s / 5f;
        g.DrawLine(pen, m, s / 2f, s - m, s / 2f);
        g.DrawLine(pen, s / 2f, s / 2f, s / 2f, m);
    }

    static void DrawMoveConnect(Graphics g, int s)
    {
        float pw = Math.Max(1.5f, s / 8f);
        using var pen = new Pen(Color.FromArgb(33, 150, 243), pw);
        pen.EndCap = LineCap.ArrowAnchor;
        float m = s / 5f;
        g.DrawLine(pen, m, s - m, s - m, m);
        using var dotBrush = new SolidBrush(Color.FromArgb(76, 175, 80));
        float dr = Math.Max(3, s / 5f);
        g.FillEllipse(dotBrush, s - m - dr / 2, m - dr / 2, dr, dr);
    }

    static void DrawParaPusher(Graphics g, int s)
    {
        float pw = Math.Max(1f, s / 12f);
        using var pen = new Pen(Color.FromArgb(76, 175, 80), pw);
        float m = s / 5f;
        g.DrawRectangle(pen, m, m, s - 2 * m, s - 2 * m);
        g.DrawLine(pen, s / 2f, m, s / 2f, s - m);
        g.DrawLine(pen, m, s / 2f, s - m, s / 2f);
        float apw = Math.Max(1.5f, s / 8f);
        using var arrowPen = new Pen(Color.FromArgb(33, 150, 243), apw);
        arrowPen.EndCap = LineCap.ArrowAnchor;
        g.DrawLine(arrowPen, s / 4f, 3 * s / 4f, 3 * s / 4f, s / 4f);
    }

    static void DrawSumParam(Graphics g, int s)
    {
        float fs = Math.Max(7, s * 0.55f);
        using var font = new Font("Arial", fs, FontStyle.Bold);
        using var brush = new SolidBrush(Color.FromArgb(76, 175, 80));
        var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        g.DrawString("\u03A3", font, brush, s / 2f, s / 2f, sf);
    }
}
