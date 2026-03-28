// C# Script to generate simple MEP icons as PNG files
// Run with: dotnet-script generate_mep_icons.csx
// Or manually copy the output PNGs

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;

string outputDir = Path.Combine(Path.GetDirectoryName(Args.Count > 0 ? Args[0] : "."),
    @"E:\THBIM-CODE\2025\CODE\THBIM\THBIM_Core\Resources\MEP");

if (!Directory.Exists(outputDir))
    Directory.CreateDirectory(outputDir);

void SaveIcon(string name, int size, Action<Graphics, int> draw)
{
    using var bmp = new Bitmap(size, size);
    using var g = Graphics.FromImage(bmp);
    g.SmoothingMode = SmoothingMode.AntiAlias;
    g.Clear(Color.Transparent);
    draw(g, size);
    string path = Path.Combine(outputDir, $"{name}_{size}.png");
    bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
    Console.WriteLine($"Created: {path}");
}

// === BLOOM ===
void DrawBloom(Graphics g, int s)
{
    int cx = s / 2, cy = s / 2;
    int r = s / 6;
    using var pen = new Pen(Color.FromArgb(33, 150, 243), Math.Max(1, s / 10));
    using var brush = new SolidBrush(Color.FromArgb(33, 150, 243));
    g.FillEllipse(brush, cx - r, cy - r, r * 2, r * 2); // center circle
    // radiating lines
    for (int i = 0; i < 6; i++)
    {
        double angle = i * Math.PI / 3;
        int x1 = cx + (int)(r * 1.3 * Math.Cos(angle));
        int y1 = cy + (int)(r * 1.3 * Math.Sin(angle));
        int x2 = cx + (int)(r * 2.8 * Math.Cos(angle));
        int y2 = cy + (int)(r * 2.8 * Math.Sin(angle));
        g.DrawLine(pen, x1, y1, x2, y2);
    }
}
SaveIcon("Bloom", 16, DrawBloom);
SaveIcon("Bloom", 32, DrawBloom);

// === ROUTE ===
void DrawRoute(Graphics g, int s)
{
    using var pen = new Pen(Color.FromArgb(76, 175, 80), Math.Max(2, s / 8));
    pen.StartCap = LineCap.Round;
    pen.EndCap = LineCap.ArrowAnchor;
    int m = s / 6;
    g.DrawLine(pen, m, s - m, s / 2, s / 2);
    g.DrawLine(pen, s / 2, s / 2, s - m, m);
}
SaveIcon("Route", 16, DrawRoute);
SaveIcon("Route", 32, DrawRoute);

// === REROUTE ===
void DrawReroute(Graphics g, int s)
{
    using var pen = new Pen(Color.FromArgb(255, 152, 0), Math.Max(2, s / 8));
    pen.StartCap = LineCap.Round; pen.EndCap = LineCap.Round;
    int m = s / 5;
    var pts = new Point[] {
        new Point(m, s/2), new Point(s/3, s/2),
        new Point(s/3, m), new Point(2*s/3, m),
        new Point(2*s/3, s/2), new Point(s-m, s/2)
    };
    g.DrawLines(pen, pts);
}
SaveIcon("Reroute", 16, DrawReroute);

// === ELBOW UP 45 ===
void DrawElbowUp45(Graphics g, int s)
{
    using var pen = new Pen(Color.FromArgb(33, 150, 243), Math.Max(2, s / 8));
    pen.EndCap = LineCap.ArrowAnchor;
    int m = s / 5;
    g.DrawLine(pen, m, s - m, s / 2, s / 2);
    g.DrawLine(pen, s / 2, s / 2, s - m, m);
}
SaveIcon("ElbowUp45", 16, DrawElbowUp45);

// === ELBOW DOWN 45 ===
void DrawElbowDown45(Graphics g, int s)
{
    using var pen = new Pen(Color.FromArgb(33, 150, 243), Math.Max(2, s / 8));
    pen.EndCap = LineCap.ArrowAnchor;
    int m = s / 5;
    g.DrawLine(pen, m, m, s / 2, s / 2);
    g.DrawLine(pen, s / 2, s / 2, s - m, s - m);
}
SaveIcon("ElbowDown45", 16, DrawElbowDown45);

// === ELBOW UP 90 ===
void DrawElbowUp90(Graphics g, int s)
{
    using var pen = new Pen(Color.FromArgb(0, 150, 136), Math.Max(2, s / 8));
    pen.EndCap = LineCap.ArrowAnchor;
    int m = s / 5;
    g.DrawLine(pen, m, s - m, m, m);
    g.DrawLine(pen, m, m, s - m, m);
}
SaveIcon("ElbowUp90", 16, DrawElbowUp90);

// === ELBOW DOWN 90 ===
void DrawElbowDown90(Graphics g, int s)
{
    using var pen = new Pen(Color.FromArgb(0, 150, 136), Math.Max(2, s / 8));
    pen.EndCap = LineCap.ArrowAnchor;
    int m = s / 5;
    g.DrawLine(pen, m, m, m, s - m);
    g.DrawLine(pen, m, s - m, s - m, s - m);
}
SaveIcon("ElbowDown90", 16, DrawElbowDown90);

// === CONNECT ===
void DrawConnect(Graphics g, int s)
{
    using var pen = new Pen(Color.FromArgb(76, 175, 80), Math.Max(2, s / 8));
    int m = s / 5;
    g.DrawLine(pen, m, s / 2, s / 2 - 2, s / 2);
    g.DrawLine(pen, s / 2 + 2, s / 2, s - m, s / 2);
    using var brush = new SolidBrush(Color.FromArgb(76, 175, 80));
    g.FillEllipse(brush, s / 2 - 3, s / 2 - 3, 6, 6);
}
SaveIcon("Connect", 16, DrawConnect);

// === ALIGN BRANCH ===
void DrawAlignBranch(Graphics g, int s)
{
    using var pen = new Pen(Color.FromArgb(156, 39, 176), Math.Max(2, s / 8));
    int m = s / 5;
    g.DrawLine(pen, m, s / 2, s - m, s / 2); // main
    g.DrawLine(pen, s / 2, s / 2, s / 2, m); // branch up
}
SaveIcon("AlignBranch", 16, DrawAlignBranch);

// === DISCONNECT ===
void DrawDisconnect(Graphics g, int s)
{
    using var pen = new Pen(Color.FromArgb(244, 67, 54), Math.Max(2, s / 8));
    int m = s / 5;
    g.DrawLine(pen, m, s / 2, s / 2 - 3, s / 2);
    g.DrawLine(pen, s / 2 + 3, s / 2, s - m, s / 2);
    // X mark
    using var xPen = new Pen(Color.FromArgb(244, 67, 54), Math.Max(1, s / 12));
    int xr = s / 8;
    g.DrawLine(xPen, s / 2 - xr, s / 2 - xr, s / 2 + xr, s / 2 + xr);
    g.DrawLine(xPen, s / 2 + xr, s / 2 - xr, s / 2 - xr, s / 2 + xr);
}
SaveIcon("Disconnect", 16, DrawDisconnect);

// === MOVE CONNECT ===
void DrawMoveConnect(Graphics g, int s)
{
    using var pen = new Pen(Color.FromArgb(33, 150, 243), Math.Max(2, s / 8));
    pen.EndCap = LineCap.ArrowAnchor;
    int m = s / 5;
    g.DrawLine(pen, m, s - m, s - m, m); // arrow
    using var dotBrush = new SolidBrush(Color.FromArgb(76, 175, 80));
    g.FillEllipse(dotBrush, s - m - 3, m - 3, 6, 6); // connect dot
}
SaveIcon("MoveConnect", 16, DrawMoveConnect);

// === PARAPUSHER ===
void DrawParaPusher(Graphics g, int s)
{
    using var pen = new Pen(Color.FromArgb(76, 175, 80), Math.Max(1, s / 12));
    using var brush = new SolidBrush(Color.FromArgb(76, 175, 80));
    int m = s / 5;
    // Grid lines
    g.DrawRectangle(pen, m, m, s - 2 * m, s - 2 * m);
    g.DrawLine(pen, s / 2, m, s / 2, s - m);
    g.DrawLine(pen, m, s / 2, s - m, s / 2);
    // Arrow
    using var arrowPen = new Pen(Color.FromArgb(33, 150, 243), Math.Max(2, s / 8));
    arrowPen.EndCap = LineCap.ArrowAnchor;
    g.DrawLine(arrowPen, s / 4, 3 * s / 4, 3 * s / 4, s / 4);
}
SaveIcon("ParaPusher", 16, DrawParaPusher);
SaveIcon("ParaPusher", 32, DrawParaPusher);

// === SUM PARAMETERS ===
void DrawSumParam(Graphics g, int s)
{
    using var font = new Font("Arial", Math.Max(6, s * 0.5f), FontStyle.Bold);
    using var brush = new SolidBrush(Color.FromArgb(76, 175, 80));
    var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
    g.DrawString("\u03A3", font, brush, s / 2f, s / 2f, sf); // Sigma symbol
}
SaveIcon("SumParam", 16, DrawSumParam);

Console.WriteLine("All icons generated!");
