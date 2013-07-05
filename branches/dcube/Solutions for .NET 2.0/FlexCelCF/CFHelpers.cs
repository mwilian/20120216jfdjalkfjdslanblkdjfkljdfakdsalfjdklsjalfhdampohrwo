using System;
//Some helper things to compile it on CF.
namespace FlexCel.Core
{
#if (!FRAMEWORK20)
    public sealed class ClassInterfaceAttribute : Attribute
    {
        public ClassInterfaceAttribute(ClassInterfaceType t){}
    }

    public enum ClassInterfaceType
    {
        None
    }

    public sealed class SerializableAttribute : Attribute
    {
    }

    public sealed class SerializationInfo
    {
    }
#endif
    public struct StreamingContext 
    {
    }

    public sealed class BrowsableAttribute : Attribute
    {
        public BrowsableAttribute(bool t){}
    }

    public sealed class CategoryAttribute : Attribute
    {
        public CategoryAttribute(string t){}
    }

	public sealed class DescriptionAttribute : Attribute
	{
		public DescriptionAttribute(string t){}
	}
}

namespace System
{
	public sealed class ThreadStaticAttribute : Attribute
	{
		public ThreadStaticAttribute(){}
	}
}

#if (!FRAMEWORK20)
namespace System.Runtime.CompilerServices
{
	public sealed class IsVolatile
	{
		private IsVolatile()
		{
		}
	}
}
#endif

#if (FRAMEWORK20)
namespace System.Runtime.Serialization
{
}

namespace System.Drawing
{
    public enum HatchStyle
    {
        Percent50,
        Percent75,
        Percent25,
        DarkHorizontal,
        DarkVertical,
        DarkUpwardDiagonal,
        DarkDownwardDiagonal,
        SmallCheckerBoard,
        Percent70,
        LightHorizontal, //  thin horz lines
        LightVertical, //  thin vert lines
        LightUpwardDiagonal,
        LightDownwardDiagonal,
        SmallGrid,
        Percent60,
        Percent10,
        Percent05
    }

/*
    public class HatchBrush: Brush 
    {
        internal HatchStyle HatchStyle;
        internal Color ForegroundColor;
        internal Color BackgroundColor;

        internal HatchBrush(HatchStyle aStyle, Color aFgColor, Color aBgColor)
        {
            HatchStyle = aStyle;
            ForegroundColor = aFgColor;
            BackgroundColor = aBgColor;
        }
    }
*/
}
#endif
