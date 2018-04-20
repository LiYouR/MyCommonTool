using System;
using System.Drawing;
using System.Linq;

namespace CommonMath.DataClass
{
  public static class ColorHSV
  {
    public static Color FromHsv(int h, int s, int v)
    {
      while (h >= 360) h -= 360;
      while (h < 0) h += 360;
      if (s > 255) s = 255;
      if (s < 0) s = 0;
      if (v > 255) v = 255;
      if (v < 0) v = 0;

      return FromHsv((float)h, (float)s / 255.0f, (float)v / 255.0f);
    }

    private static Color FromHsv(float h, float s, float v)
    {
      if (h > 360.0) throw new ArgumentOutOfRangeException("h", h, "0~360で指定してください");
      if (h < 0.0) throw new ArgumentOutOfRangeException("h", h, "0~360で指定してください");
      if (s > 1.0) throw new ArgumentOutOfRangeException("s", s, "0.0~1.0で指定してください");
      if (s < 0.0) throw new ArgumentOutOfRangeException("s", s, "0.0~1.0で指定してください");
      if (v > 1.0) throw new ArgumentOutOfRangeException("v", v, "0.0~1.0で指定してください");
      if (v < 0.0) throw new ArgumentOutOfRangeException("v", v, "0.0~1.0で指定してください");

      Color resColor = Color.Transparent;

      if (s == 0.0)//gray
      {
        int rgb = Convert.ToInt16((float)(v * 255));
        resColor = Color.FromArgb(rgb, rgb, rgb);
      }
      else
      {
        int Hi = (int)(Math.Floor(h / 60.0f) % 6.0f);
        float f = (h / 60.0f) - Hi;

        float p = v * (1 - s);
        float q = v * (1 - f * s);
        float t = v * (1 - (1 - f) * s);

        float r = 0.0f;
        float g = 0.0f;
        float b = 0.0f;

        switch (Hi)
        {
          case 0: r = v; g = t; b = p; break;
          case 1: r = q; g = v; b = p; break;
          case 2: r = p; g = v; b = t; break;
          case 3: r = p; g = q; b = v; break;
          case 4: r = t; g = p; b = v; break;
          case 5: r = v; g = p; b = q; break;
          default: break;
        }

        r *= 255;
        g *= 255;
        b *= 255;

        resColor = Color.FromArgb((int)r, (int)g, (int)b);
      }
      return resColor;
    }
  }

  public static class ColorExtension
  {
    /// <summary>
    /// 色相(Hue) 0-360
    /// </summary>
    /// <param name="c"></param>
    /// <returns></returns>
    public static int h(this Color c)
    {
      float min = (new float[] { c.R, c.G, c.B }).Min();
      float max = (new float[] { c.R, c.G, c.B }).Max();

      if (max == 0) return 0;

      float h = 0;

      if (max == c.R) h = 60 * (c.G - c.B) / (max - min) + 0;
      else if (max == c.G) h = 60 * (c.B - c.R) / (max - min) + 120;
      else if (max == c.B) h = 60 * (c.R - c.G) / (max - min) + 240;

      if (h < 0) h += 360;

      return (int)Math.Round(h);
    }

    /// <summary>
    /// 彩度 0-255
    /// </summary>
    /// <param name="c"></param>
    /// <returns></returns>
    public static int s(this Color c)
    {
      float min = (new float[] { c.R, c.G, c.B }).Min();
      float max = (new float[] { c.R, c.G, c.B }).Max();

      if (max == 0) return 0;
      return (int)(255 * (max - min) / max);
    }

    /// <summary>
    /// 明度
    /// </summary>
    /// <param name="c"></param>
    /// <returns></returns>
    public static int v(this Color c)
    {
      return (int)(255.0f * (new float[] { c.R, c.G, c.B }).Max());
    }

    /// <summary>
    /// 現在の色を基準にHSV色空間を移動します
    /// </summary>
    /// <param name="c"></param>
    /// <param name="offsetH"></param>
    /// <param name="offsetS"></param>
    /// <param name="offsetV"></param>
    /// <returns></returns>
    public static Color Offset(this Color c, int offsetH, int offsetS, int offsetV)
    {
      int newH = (int)(c.h() + offsetH);
      int newS = (int)(c.s() + offsetS);
      int newV = (int)(c.v() + offsetV);

      return ColorHSV.FromHsv(newH, newS, newV);
    }

    /// <summary>
    /// 現在の色を文字列として返します
    /// </summary>
    /// <param name="c"></param>
    /// <returns></returns>
    public static string ToString2(this Color c)
    {
      return string.Format("R={0}, G={1}, B={2}, H={3}, S={4}, V={5}",
        c.R,
        c.G,
        c.B,
        c.h(),
        c.s(),
        c.v());
    }
  }
}
