using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using XtendedRefs.ComAddIn.Domain.Documents;
using XtendedRefs.ComAddIn.Infrastructure.OfficeCom;
using FontStyle = System.Drawing.FontStyle;

namespace XtendedRefs.ComAddIn;

public sealed partial class Connect
{
    public void setList(List<LevelStyleItem> list)
    {
        _styles = new List<LevelStyleItem>();
        for (int i = 0; i < 10; i++)
        {
            _styles.Add(new LevelStyleItem { StyleLevel = i, StyleName = $"i={i}" });
        }
    }

    public void UpdateStyles(List<LevelStyleItem> newStyles)
    {
        _styles = newStyles ?? new List<LevelStyleItem>();
        _startIndex = 0;
        _ribbon?.Invalidate();
    }

    public bool GetStyleButtonVisible(IRibbonControl control)
    {
        int globalIndex = _startIndex + GetSlotIndex(control.Id);
        return globalIndex < _styles.Count;
    }

    public string GetStyleButtonLabel(IRibbonControl control)
    {
        int globalIndex = _startIndex + GetSlotIndex(control.Id);
        return globalIndex < _styles.Count ? _styles[globalIndex].StyleName : string.Empty;
    }

    public stdole.IPictureDisp GetStyleButtonImage(IRibbonControl control)
    {
        int globalIndex = _startIndex + GetSlotIndex(control.Id);
        if (globalIndex >= _styles.Count)
        {
            return null;
        }

        return null;
    }

    public void OnStyleButtonAction(IRibbonControl control)
    {
        int globalIndex = _startIndex + GetSlotIndex(control.Id);
        if (globalIndex < _styles.Count)
        {
            ApplyStyle(_styles[globalIndex]);
        }
    }

    public bool GetNavUpVisible(IRibbonControl control) => _startIndex > 0;

    public bool GetNavDownVisible(IRibbonControl control) => _startIndex + PageSize < _styles.Count;

    public void OnNavUp(IRibbonControl control)
    {
        _startIndex = Math.Max(0, _startIndex - PageSize);
        _ribbon?.Invalidate();
    }

    public void OnNavDown(IRibbonControl control)
    {
        if (_startIndex + PageSize < _styles.Count)
        {
            _startIndex += PageSize;
            _ribbon?.Invalidate();
        }
    }

    public Bitmap CreateIconFromNumber(string number)
    {
        int size = 32;
        Bitmap bmp = new Bitmap(size, size);

        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            g.Clear(Color.FromArgb(240, 240, 240));

            using (Pen pen = new Pen(Color.DarkGray, 1))
            {
                g.DrawRectangle(pen, 0, 0, size - 1, size - 1);
            }

            float fontSize = number.Length > 2 ? 8 : 12;
            using (Font font = new Font("Segoe UI", fontSize, FontStyle.Bold))
            {
                StringFormat sf = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                };

                g.DrawString(number, font, Brushes.RoyalBlue, new RectangleF(0, 0, size, size), sf);
            }
        }

        return bmp;
    }

    private int GetSlotIndex(string id) => int.Parse(id.Substring("btnStyle".Length));

    private void ApplyStyle(LevelStyleItem item)
    {
    }
}
