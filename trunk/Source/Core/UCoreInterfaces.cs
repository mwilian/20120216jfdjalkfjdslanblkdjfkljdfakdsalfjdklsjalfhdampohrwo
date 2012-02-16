using System;
#if (MONOTOUCH)
    using Color = MonoTouch.UIKit.UIColor;
#endif

#if (WPF)
using System.Windows.Media;
using real = System.Double;
#else
using System.Drawing;
using real = System.Single;
#endif

namespace FlexCel.Core
{
    /// <summary>
    /// Interface for passing font information. Internal use.
    /// </summary>
    public interface IFlexCelFontList
    {
        /// <summary>
        /// Returns the font definition for a given font index.
        /// </summary>
        /// <param name="fontIndex">Font index. 0-based</param>
        /// <returns>Font definition</returns>
        TFlxFont GetFont(int fontIndex);

        /// <summary>
        /// Fonts in the document.
        /// </summary>
        int FontCount { get; }

        /// <summary>
        /// Adds a new Font.
        /// </summary>
        /// <param name="aFont"></param>
        /// <returns></returns>
        int AddFont(TFlxFont aFont);
    }


    /// <summary>
    /// Interface for passing palette and theme information.
    /// XlsFile implements IFlexCelPalette, so you can pass any XlsFile object whenever you need to use this interface.
    /// </summary>
    public interface IFlexCelPalette
    {
        ///<inheritdoc cref = "FlexCel.Core.ExcelFile.NearestColorTheme(System.Drawing.Color,out Double)" />
        TThemeColor NearestColorTheme(Color value, out double tint);

        ///<inheritdoc cref = "FlexCel.Core.ExcelFile.NearestColorIndex(System.Drawing.Color)" />
        int NearestColorIndex(Color value);

        ///<inheritdoc cref = "FlexCel.Core.ExcelFile.GetColorTheme(FlexCel.Core.TThemeColor)" />
        TDrawingColor GetColorTheme(TThemeColor themeColor);

        ///<inheritdoc cref = "FlexCel.Core.ExcelFile.GetColorPalette(System.Int32)" />
        Color GetColorPalette(int index);

        ///<inheritdoc cref = "FlexCel.Core.ExcelFile.PaletteContainsColor(FlexCel.Core.TExcelColor)" />
        bool PaletteContainsColor(TExcelColor value);

#if (FRAMEWORK30)
        ///<inheritdoc cref = "FlexCel.Core.ExcelFile.GetTheme()" />
        TTheme GetTheme();
#endif
    }

    /// <summary>
    /// Interface for row heights and columns widths. XlsFile implements this interface, so you can pass an XlsFile object anytime you need to pass this interface.
    /// </summary>
    public interface IRowColSize
    {
        ///<inheritdoc cref = "FlexCel.Core.ExcelFile.DefaultColWidth" />
        int DefaultColWidth { get; set; }

        ///<inheritdoc cref = "FlexCel.Core.ExcelFile.DefaultRowHeight" />
        int DefaultRowHeight { get; set; }

        ///<inheritdoc cref = "FlexCel.Core.ExcelFile.IsEmptyRow(int)" />
        bool IsEmptyRow(int row);

        ///<inheritdoc cref = "FlexCel.Core.ExcelFile.GetRowHeight(int, bool)" />
        int GetRowHeight(int row, bool HiddenIsZero);

        ///<inheritdoc cref = "FlexCel.Core.ExcelFile.GetColWidth(int, bool)" />
        int GetColWidth(int col, bool HiddenIsZero);

        ///<inheritdoc cref = "FlexCel.Core.ExcelFile.ShowFormulaText" />
        bool ShowFormulaText { get; set; }

        ///<inheritdoc cref = "FlexCel.Core.ExcelFile.GetDefaultFont" />
        TFlxFont GetDefaultFont { get; }

        ///<inheritdoc cref = "FlexCel.Core.ExcelFile.WidthCorrection" />
        real WidthCorrection { get; set; }

        ///<inheritdoc cref = "FlexCel.Core.ExcelFile.HeightCorrection" />
        real HeightCorrection { get; set; }
    }


    /// <summary>
    /// Use this interface to read or write Embedded drawing objects inside other object.
    /// </summary>
    public interface IEmbeddedObjects
    {
        /// <summary>
        /// The number of objects that are embedded inside this object.
        /// </summary>
        int ObjectCount { get; }

        /// <summary>
        /// Returns information on an object and all of its children. 
        /// </summary>
        /// <param name="objectIndex">Index of the object (1-based)</param>
        /// <param name="GetShapeOptions">When true, shape options will be retrieved. As this can be a slow operation,
        /// only specify true when you really need those options.</param>
        /// <returns></returns>
        TShapeProperties GetObjectProperties(int objectIndex, bool GetShapeOptions);

        /// <summary>
        /// Changes the text inside an object of this object.
        /// </summary>
        /// <param name="objectIndex">Index of the object, between 1 and <see cref="ObjectCount"/></param>
        /// <param name="objectPath">Index to the child object you want to change the text.
        /// If it is a simple object, you can use String.Empty here, if not you need to get the ObjectPath from <see cref="GetObjectProperties"/><br></br>
        /// If it is "absolute"(it starts with "\\"), then the path includes the objectIndex, and the objectIndex is
        /// not used. An object path of "\\1\\2\\3" is exactly the same as using objectIndex = 1 and objectPath = "2\\3"</param>
        /// <param name="text">Text you want to use. Use null to delete text from an AutoShape.</param>
        void SetObjectText(int objectIndex, string objectPath, TRichString text);

        /// <summary>
        /// Deletes the graphic object at objectIndex. Use it with care, there are some graphics objects you
        /// <b>don't</b> want to remove (like comment boxes when you don't delete the associated comment.)
        /// </summary>
        /// <param name="objectIndex">Index of the object (1 based).</param>
        void DeleteObject(int objectIndex);
    }
}
