#if(MONOTOUCH)
using MonoTouch.UIKit;
using Color = MonoTouch.UIKit.UIColor;

namespace FlexCel.Core
{
	internal class SystemColors
	{
		public static Color WindowText = ColorUtil.FromArgb(0xff, 0x00, 0x00, 0x00);
        public static Color WindowFrame = ColorUtil.FromArgb(0xff, 0xff, 0xff, 0xff);
        public static Color DefaultForeground = ColorUtil.FromArgb(0xff, 0x00, 0x00, 0x00);
        public static Color DefaultBackground = ColorUtil.FromArgb(0xff, 0xff, 0xff, 0xff);

		
		public static Color ActiveBorder  = ColorUtil.FromArgb(0xff, 0xb4, 0xb4, 0xb4);
		public static Color ActiveCaption  = ColorUtil.FromArgb(0xff, 0x99, 0xb4, 0xd1);
		public static Color ActiveCaptionText  = ColorUtil.FromArgb(0xff, 0x00, 0x00, 0x00);
		public static Color AppWorkspace  = ColorUtil.FromArgb(0xff, 0xab, 0xab, 0xab);
		public static Color Control  = ColorUtil.FromArgb(0xff, 0xf0, 0xf0, 0xf0);
		public static Color ControlDark  = ColorUtil.FromArgb(0xff, 0xa0, 0xa0, 0xa0);
		public static Color ControlDarkDark  = ColorUtil.FromArgb(0xff, 0x69, 0x69, 0x69);
		public static Color ControlLight  = ColorUtil.FromArgb(0xff, 0xe3, 0xe3, 0xe3);
		public static Color ControlLightLight  = ColorUtil.FromArgb(0xff, 0xff, 0xff, 0xff);
		public static Color ControlText  = ColorUtil.FromArgb(0xff, 0x00, 0x00, 0x00);
		public static Color Desktop  = ColorUtil.FromArgb(0xff, 0x00, 0x00, 0x00);
		public static Color GrayText  = ColorUtil.FromArgb(0xff, 0x6d, 0x6d, 0x6d);
		public static Color Highlight  = ColorUtil.FromArgb(0xff, 0x33, 0x99, 0xff);
		public static Color HighlightText  = ColorUtil.FromArgb(0xff, 0xff, 0xff, 0xff);
		public static Color HotTrack  = ColorUtil.FromArgb(0xff, 0x00, 0x66, 0xcc);
		public static Color InactiveBorder  = ColorUtil.FromArgb(0xff, 0xf4, 0xf7, 0xfc);
		public static Color InactiveCaption  = ColorUtil.FromArgb(0xff, 0xbf, 0xcd, 0xdb);
		public static Color InactiveCaptionText  = ColorUtil.FromArgb(0xff, 0x43, 0x4e, 0x54);
		public static Color Info  = ColorUtil.FromArgb(0xff, 0xff, 0xff, 0xe1);
		public static Color InfoText  = ColorUtil.FromArgb(0xff, 0x00, 0x00, 0x00);
		public static Color Menu  = ColorUtil.FromArgb(0xff, 0xf0, 0xf0, 0xf0);
		public static Color MenuText  = ColorUtil.FromArgb(0xff, 0x00, 0x00, 0x00);
		public static Color ScrollBar  = ColorUtil.FromArgb(0xff, 0xc8, 0xc8, 0xc8);
		public static Color Window  = ColorUtil.FromArgb(0xff, 0xff, 0xff, 0xff);
		//public static Color WindowFrame  = ColorUtil.FromArgb(0xff, 0x64, 0x64, 0x64);
		//public static Color WindowText  = ColorUtil.FromArgb(0xff, 0x00, 0x00, 0x00);
	
		public static Color ButtonFace  = ColorUtil.FromArgb(0xff, 0xf0, 0xf0, 0xf0);
		public static Color ButtonHighlight  = ColorUtil.FromArgb(0xff, 0xff, 0xff, 0xff);
		public static Color ButtonShadow  = ColorUtil.FromArgb(0xff, 0xa0, 0xa0, 0xa0);
		public static Color GradientActiveCaption  = ColorUtil.FromArgb(0xff, 0xb9, 0xd1, 0xea);
		public static Color GradientInactiveCaption  = ColorUtil.FromArgb(0xff, 0xd7, 0xe4, 0xf2);
		public static Color MenuBar  = ColorUtil.FromArgb(0xff, 0xf0, 0xf0, 0xf0);
		public static Color MenuHighlight  = ColorUtil.FromArgb(0xff, 0x33, 0x99, 0xff);

	}
}
#endif
