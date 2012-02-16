#if(SILVERLIGHT || MONOTOUCH)
#if (MONOTOUCH)
using Color = MonoTouch.UIKit.UIColor;
#else
using System.Windows.Media;
#endif
namespace FlexCel.Core
{
#region Colors
    internal static class Colors
    {
        public static Color Transparent = ColorUtil.FromArgb(0x0, 0xFF, 0xFF, 0xFF);
        public static Color AliceBlue = ColorUtil.FromArgb(0xFF, 0xF0, 0xF8, 0xFF);
        public static Color AntiqueWhite = ColorUtil.FromArgb(0xFF, 0xFA, 0xEB, 0xD7);
        public static Color Aqua = ColorUtil.FromArgb(0xFF, 0x0, 0xFF, 0xFF);
        public static Color Aquamarine = ColorUtil.FromArgb(0xFF, 0x7F, 0xFF, 0xD4);
        public static Color Azure = ColorUtil.FromArgb(0xFF, 0xF0, 0xFF, 0xFF);
        public static Color Beige = ColorUtil.FromArgb(0xFF, 0xF5, 0xF5, 0xDC);
        public static Color Bisque = ColorUtil.FromArgb(0xFF, 0xFF, 0xE4, 0xC4);
        public static Color Black = ColorUtil.FromArgb(0xFF, 0x0, 0x0, 0x0);
        public static Color BlanchedAlmond = ColorUtil.FromArgb(0xFF, 0xFF, 0xEB, 0xCD);
        public static Color Blue = ColorUtil.FromArgb(0xFF, 0x0, 0x0, 0xFF);
        public static Color BlueViolet = ColorUtil.FromArgb(0xFF, 0x8A, 0x2B, 0xE2);
        public static Color Brown = ColorUtil.FromArgb(0xFF, 0xA5, 0x2A, 0x2A);
        public static Color BurlyWood = ColorUtil.FromArgb(0xFF, 0xDE, 0xB8, 0x87);
        public static Color CadetBlue = ColorUtil.FromArgb(0xFF, 0x5F, 0x9E, 0xA0);
        public static Color Chartreuse = ColorUtil.FromArgb(0xFF, 0x7F, 0xFF, 0x0);
        public static Color Chocolate = ColorUtil.FromArgb(0xFF, 0xD2, 0x69, 0x1E);
        public static Color Coral = ColorUtil.FromArgb(0xFF, 0xFF, 0x7F, 0x50);
        public static Color CornflowerBlue = ColorUtil.FromArgb(0xFF, 0x64, 0x95, 0xED);
        public static Color Cornsilk = ColorUtil.FromArgb(0xFF, 0xFF, 0xF8, 0xDC);
        public static Color Crimson = ColorUtil.FromArgb(0xFF, 0xDC, 0x14, 0x3C);
        public static Color Cyan = ColorUtil.FromArgb(0xFF, 0x0, 0xFF, 0xFF);
        public static Color DarkBlue = ColorUtil.FromArgb(0xFF, 0x0, 0x0, 0x8B);
        public static Color DarkCyan = ColorUtil.FromArgb(0xFF, 0x0, 0x8B, 0x8B);
        public static Color DarkGoldenrod = ColorUtil.FromArgb(0xFF, 0xB8, 0x86, 0xB);
        public static Color DarkGray = ColorUtil.FromArgb(0xFF, 0xA9, 0xA9, 0xA9);
        public static Color DarkGreen = ColorUtil.FromArgb(0xFF, 0x0, 0x64, 0x0);
        public static Color DarkKhaki = ColorUtil.FromArgb(0xFF, 0xBD, 0xB7, 0x6B);
        public static Color DarkMagenta = ColorUtil.FromArgb(0xFF, 0x8B, 0x0, 0x8B);
        public static Color DarkOliveGreen = ColorUtil.FromArgb(0xFF, 0x55, 0x6B, 0x2F);
        public static Color DarkOrange = ColorUtil.FromArgb(0xFF, 0xFF, 0x8C, 0x0);
        public static Color DarkOrchid = ColorUtil.FromArgb(0xFF, 0x99, 0x32, 0xCC);
        public static Color DarkRed = ColorUtil.FromArgb(0xFF, 0x8B, 0x0, 0x0);
        public static Color DarkSalmon = ColorUtil.FromArgb(0xFF, 0xE9, 0x96, 0x7A);
        public static Color DarkSeaGreen = ColorUtil.FromArgb(0xFF, 0x8F, 0xBC, 0x8B);
        public static Color DarkSlateBlue = ColorUtil.FromArgb(0xFF, 0x48, 0x3D, 0x8B);
        public static Color DarkSlateGray = ColorUtil.FromArgb(0xFF, 0x2F, 0x4F, 0x4F);
        public static Color DarkTurquoise = ColorUtil.FromArgb(0xFF, 0x0, 0xCE, 0xD1);
        public static Color DarkViolet = ColorUtil.FromArgb(0xFF, 0x94, 0x0, 0xD3);
        public static Color DeepPink = ColorUtil.FromArgb(0xFF, 0xFF, 0x14, 0x93);
        public static Color DeepSkyBlue = ColorUtil.FromArgb(0xFF, 0x0, 0xBF, 0xFF);
        public static Color DimGray = ColorUtil.FromArgb(0xFF, 0x69, 0x69, 0x69);
        public static Color DodgerBlue = ColorUtil.FromArgb(0xFF, 0x1E, 0x90, 0xFF);
        public static Color Firebrick = ColorUtil.FromArgb(0xFF, 0xB2, 0x22, 0x22);
        public static Color FloralWhite = ColorUtil.FromArgb(0xFF, 0xFF, 0xFA, 0xF0);
        public static Color ForestGreen = ColorUtil.FromArgb(0xFF, 0x22, 0x8B, 0x22);
        public static Color Fuchsia = ColorUtil.FromArgb(0xFF, 0xFF, 0x0, 0xFF);
        public static Color Gainsboro = ColorUtil.FromArgb(0xFF, 0xDC, 0xDC, 0xDC);
        public static Color GhostWhite = ColorUtil.FromArgb(0xFF, 0xF8, 0xF8, 0xFF);
        public static Color Gold = ColorUtil.FromArgb(0xFF, 0xFF, 0xD7, 0x0);
        public static Color Goldenrod = ColorUtil.FromArgb(0xFF, 0xDA, 0xA5, 0x20);
        public static Color Gray = ColorUtil.FromArgb(0xFF, 0x80, 0x80, 0x80);
        public static Color Green = ColorUtil.FromArgb(0xFF, 0x0, 0x80, 0x0);
        public static Color GreenYellow = ColorUtil.FromArgb(0xFF, 0xAD, 0xFF, 0x2F);
        public static Color Honeydew = ColorUtil.FromArgb(0xFF, 0xF0, 0xFF, 0xF0);
        public static Color HotPink = ColorUtil.FromArgb(0xFF, 0xFF, 0x69, 0xB4);
        public static Color IndianRed = ColorUtil.FromArgb(0xFF, 0xCD, 0x5C, 0x5C);
        public static Color Indigo = ColorUtil.FromArgb(0xFF, 0x4B, 0x0, 0x82);
        public static Color Ivory = ColorUtil.FromArgb(0xFF, 0xFF, 0xFF, 0xF0);
        public static Color Khaki = ColorUtil.FromArgb(0xFF, 0xF0, 0xE6, 0x8C);
        public static Color Lavender = ColorUtil.FromArgb(0xFF, 0xE6, 0xE6, 0xFA);
        public static Color LavenderBlush = ColorUtil.FromArgb(0xFF, 0xFF, 0xF0, 0xF5);
        public static Color LawnGreen = ColorUtil.FromArgb(0xFF, 0x7C, 0xFC, 0x0);
        public static Color LemonChiffon = ColorUtil.FromArgb(0xFF, 0xFF, 0xFA, 0xCD);
        public static Color LightBlue = ColorUtil.FromArgb(0xFF, 0xAD, 0xD8, 0xE6);
        public static Color LightCoral = ColorUtil.FromArgb(0xFF, 0xF0, 0x80, 0x80);
        public static Color LightCyan = ColorUtil.FromArgb(0xFF, 0xE0, 0xFF, 0xFF);
        public static Color LightGoldenrodYellow = ColorUtil.FromArgb(0xFF, 0xFA, 0xFA, 0xD2);
        public static Color LightGray = ColorUtil.FromArgb(0xFF, 0xD3, 0xD3, 0xD3);
        public static Color LightGreen = ColorUtil.FromArgb(0xFF, 0x90, 0xEE, 0x90);
        public static Color LightPink = ColorUtil.FromArgb(0xFF, 0xFF, 0xB6, 0xC1);
        public static Color LightSalmon = ColorUtil.FromArgb(0xFF, 0xFF, 0xA0, 0x7A);
        public static Color LightSeaGreen = ColorUtil.FromArgb(0xFF, 0x20, 0xB2, 0xAA);
        public static Color LightSkyBlue = ColorUtil.FromArgb(0xFF, 0x87, 0xCE, 0xFA);
        public static Color LightSlateGray = ColorUtil.FromArgb(0xFF, 0x77, 0x88, 0x99);
        public static Color LightSteelBlue = ColorUtil.FromArgb(0xFF, 0xB0, 0xC4, 0xDE);
        public static Color LightYellow = ColorUtil.FromArgb(0xFF, 0xFF, 0xFF, 0xE0);
        public static Color Lime = ColorUtil.FromArgb(0xFF, 0x0, 0xFF, 0x0);
        public static Color LimeGreen = ColorUtil.FromArgb(0xFF, 0x32, 0xCD, 0x32);
        public static Color Linen = ColorUtil.FromArgb(0xFF, 0xFA, 0xF0, 0xE6);
        public static Color Magenta = ColorUtil.FromArgb(0xFF, 0xFF, 0x0, 0xFF);
        public static Color Maroon = ColorUtil.FromArgb(0xFF, 0x80, 0x0, 0x0);
        public static Color MediumAquamarine = ColorUtil.FromArgb(0xFF, 0x66, 0xCD, 0xAA);
        public static Color MediumBlue = ColorUtil.FromArgb(0xFF, 0x0, 0x0, 0xCD);
        public static Color MediumOrchid = ColorUtil.FromArgb(0xFF, 0xBA, 0x55, 0xD3);
        public static Color MediumPurple = ColorUtil.FromArgb(0xFF, 0x93, 0x70, 0xDB);
        public static Color MediumSeaGreen = ColorUtil.FromArgb(0xFF, 0x3C, 0xB3, 0x71);
        public static Color MediumSlateBlue = ColorUtil.FromArgb(0xFF, 0x7B, 0x68, 0xEE);
        public static Color MediumSpringGreen = ColorUtil.FromArgb(0xFF, 0x0, 0xFA, 0x9A);
        public static Color MediumTurquoise = ColorUtil.FromArgb(0xFF, 0x48, 0xD1, 0xCC);
        public static Color MediumVioletRed = ColorUtil.FromArgb(0xFF, 0xC7, 0x15, 0x85);
        public static Color MidnightBlue = ColorUtil.FromArgb(0xFF, 0x19, 0x19, 0x70);
        public static Color MintCream = ColorUtil.FromArgb(0xFF, 0xF5, 0xFF, 0xFA);
        public static Color MistyRose = ColorUtil.FromArgb(0xFF, 0xFF, 0xE4, 0xE1);
        public static Color Moccasin = ColorUtil.FromArgb(0xFF, 0xFF, 0xE4, 0xB5);
        public static Color NavajoWhite = ColorUtil.FromArgb(0xFF, 0xFF, 0xDE, 0xAD);
        public static Color Navy = ColorUtil.FromArgb(0xFF, 0x0, 0x0, 0x80);
        public static Color OldLace = ColorUtil.FromArgb(0xFF, 0xFD, 0xF5, 0xE6);
        public static Color Olive = ColorUtil.FromArgb(0xFF, 0x80, 0x80, 0x0);
        public static Color OliveDrab = ColorUtil.FromArgb(0xFF, 0x6B, 0x8E, 0x23);
        public static Color Orange = ColorUtil.FromArgb(0xFF, 0xFF, 0xA5, 0x0);
        public static Color OrangeRed = ColorUtil.FromArgb(0xFF, 0xFF, 0x45, 0x0);
        public static Color Orchid = ColorUtil.FromArgb(0xFF, 0xDA, 0x70, 0xD6);
        public static Color PaleGoldenrod = ColorUtil.FromArgb(0xFF, 0xEE, 0xE8, 0xAA);
        public static Color PaleGreen = ColorUtil.FromArgb(0xFF, 0x98, 0xFB, 0x98);
        public static Color PaleTurquoise = ColorUtil.FromArgb(0xFF, 0xAF, 0xEE, 0xEE);
        public static Color PaleVioletRed = ColorUtil.FromArgb(0xFF, 0xDB, 0x70, 0x93);
        public static Color PapayaWhip = ColorUtil.FromArgb(0xFF, 0xFF, 0xEF, 0xD5);
        public static Color PeachPuff = ColorUtil.FromArgb(0xFF, 0xFF, 0xDA, 0xB9);
        public static Color Peru = ColorUtil.FromArgb(0xFF, 0xCD, 0x85, 0x3F);
        public static Color Pink = ColorUtil.FromArgb(0xFF, 0xFF, 0xC0, 0xCB);
        public static Color Plum = ColorUtil.FromArgb(0xFF, 0xDD, 0xA0, 0xDD);
        public static Color PowderBlue = ColorUtil.FromArgb(0xFF, 0xB0, 0xE0, 0xE6);
        public static Color Purple = ColorUtil.FromArgb(0xFF, 0x80, 0x0, 0x80);
        public static Color Red = ColorUtil.FromArgb(0xFF, 0xFF, 0x0, 0x0);
        public static Color RosyBrown = ColorUtil.FromArgb(0xFF, 0xBC, 0x8F, 0x8F);
        public static Color RoyalBlue = ColorUtil.FromArgb(0xFF, 0x41, 0x69, 0xE1);
        public static Color SaddleBrown = ColorUtil.FromArgb(0xFF, 0x8B, 0x45, 0x13);
        public static Color Salmon = ColorUtil.FromArgb(0xFF, 0xFA, 0x80, 0x72);
        public static Color SandyBrown = ColorUtil.FromArgb(0xFF, 0xF4, 0xA4, 0x60);
        public static Color SeaGreen = ColorUtil.FromArgb(0xFF, 0x2E, 0x8B, 0x57);
        public static Color SeaShell = ColorUtil.FromArgb(0xFF, 0xFF, 0xF5, 0xEE);
        public static Color Sienna = ColorUtil.FromArgb(0xFF, 0xA0, 0x52, 0x2D);
        public static Color Silver = ColorUtil.FromArgb(0xFF, 0xC0, 0xC0, 0xC0);
        public static Color SkyBlue = ColorUtil.FromArgb(0xFF, 0x87, 0xCE, 0xEB);
        public static Color SlateBlue = ColorUtil.FromArgb(0xFF, 0x6A, 0x5A, 0xCD);
        public static Color SlateGray = ColorUtil.FromArgb(0xFF, 0x70, 0x80, 0x90);
        public static Color Snow = ColorUtil.FromArgb(0xFF, 0xFF, 0xFA, 0xFA);
        public static Color SpringGreen = ColorUtil.FromArgb(0xFF, 0x0, 0xFF, 0x7F);
        public static Color SteelBlue = ColorUtil.FromArgb(0xFF, 0x46, 0x82, 0xB4);
        public static Color Tan = ColorUtil.FromArgb(0xFF, 0xD2, 0xB4, 0x8C);
        public static Color Teal = ColorUtil.FromArgb(0xFF, 0x0, 0x80, 0x80);
        public static Color Thistle = ColorUtil.FromArgb(0xFF, 0xD8, 0xBF, 0xD8);
        public static Color Tomato = ColorUtil.FromArgb(0xFF, 0xFF, 0x63, 0x47);
        public static Color Turquoise = ColorUtil.FromArgb(0xFF, 0x40, 0xE0, 0xD0);
        public static Color Violet = ColorUtil.FromArgb(0xFF, 0xEE, 0x82, 0xEE);
        public static Color Wheat = ColorUtil.FromArgb(0xFF, 0xF5, 0xDE, 0xB3);
        public static Color White = ColorUtil.FromArgb(0xFF, 0xFF, 0xFF, 0xFF);
        public static Color WhiteSmoke = ColorUtil.FromArgb(0xFF, 0xF5, 0xF5, 0xF5);
        public static Color Yellow = ColorUtil.FromArgb(0xFF, 0xFF, 0xFF, 0x0);
        public static Color YellowGreen = ColorUtil.FromArgb(0xFF, 0x9A, 0xCD, 0x32);
    }
}
#endregion

#if(SILVERLIGHT)

#endif
#endif
