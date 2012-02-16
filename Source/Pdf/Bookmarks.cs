using System;
using System.Collections;
using System.Collections.Generic;

using FlexCel.Core;

#if (WPF)
using System.Windows.Media;
#else
using System.Drawing;
#endif


namespace FlexCel.Pdf
{
	/// <summary>
	/// This enum indicates the text style for a bookmark entry.
	/// You can combine the entries by or'ing them together.
	/// </summary>
	[Flags]
	public enum TBookmarkStyle
	{
		/// <summary>
		/// Normal text.
		/// </summary>
		None = 0,

		/// <summary>
		/// Italic text.
		/// </summary>
		Italic = 1,

		/// <summary>
		/// Bold text.
		/// </summary>
		Bold = 2
	}

	/// <summary>
	/// An entry on the Bookmark list for a PDF file.
	/// </summary>
	public class TBookmark: ICloneable
	{
		#region Private variables
		private string FTitle;
		private TPdfDestination FDestination;
		private bool FChildrenCollapsed;
		private Color FTextColor;
		private TBookmarkStyle FTextStyle;
		internal TBookmarkList FChildren;
		#endregion

		#region Constructors
		/// <summary>
		/// Creates a new TBookmark instance.
		/// </summary>
		/// <param name="aTitle">Title of the bookmark item.</param>
		/// <param name="aDestination">Page where the bookmark points to.</param>
		/// <param name="aChildrenCollapsed">If true all children from this bookmark will be collapsed.</param>
		public TBookmark(string aTitle, TPdfDestination aDestination, bool aChildrenCollapsed)
		{
			FTitle = aTitle;
			FDestination = aDestination;
			FChildrenCollapsed = aChildrenCollapsed;
			FTextColor = ColorUtil.FromArgb(0,0,0);
			FTextStyle = TBookmarkStyle.None;
			FChildren = new TBookmarkList();
		}

		/// <summary>
		/// Creates a new TBookmark instance.
		/// </summary>
		/// <param name="aTitle">Title of the bookmark item.</param>
		/// <param name="aDestination">Page where the bookmark points to.</param>
		/// <param name="aChildrenCollapsed">If true all children from this bookmark will be collapsed.</param>
		/// <param name="aTextColor">Text color for the bookmark entry.</param>
		/// <param name="aTextStyle">Text style for the bookmark entry.</param>
		public TBookmark(string aTitle, TPdfDestination aDestination, bool aChildrenCollapsed, Color aTextColor, TBookmarkStyle aTextStyle): this(aTitle, aDestination, aChildrenCollapsed)
		{
			FTextColor = aTextColor;
			FTextStyle = aTextStyle;
		}
		#endregion

		#region Public Properties
		/// <summary>
		/// Title of the bookmark item.
		/// </summary>
		public string Title {get {return FTitle;} set{FTitle = value;}}

		/// <summary>
		/// Page where the bookmark points to.
		/// </summary>
		public TPdfDestination Destination {get {return FDestination;} set{FDestination = value;}}

		/// <summary>
		/// If true, all children of this bookmark will be collapsed.
		/// </summary>
		public bool ChildrenCollapsed {get {return FChildrenCollapsed;} set{FChildrenCollapsed = value;}}

		/// <summary>
		/// Text color for the bookmark entry.
		/// </summary>
		public Color TextColor {get {return FTextColor;} set{FTextColor = value;}}

		/// <summary>
		/// Text style for the bookmark entry.
		/// </summary>
		public TBookmarkStyle TextStyle {get {return FTextStyle;} set{FTextStyle = value;}}
		#endregion

		#region Public methods
		/// <summary>
		/// Adds a new child of this bookmark on the outline.
		/// </summary>
		/// <param name="child">Child bookmark to add.</param>
		public void AddChild(TBookmark child)
		{
			FChildren.Add(child);
		}

		/// <summary>
		/// Returns the number of children of this bookmark.
		/// </summary>
        public int ChildCount
		{
			get

			{
				return FChildren.Count;
			}
		}

        /// <summary>
        /// Returns one child of the current bookmark.
        /// </summary>
        /// <param name="i">Number of child to get.</param>
        /// <returns>i-th Child of the bookmark.</returns>
		public TBookmark Child(int i)
		{
			return FChildren[i];
		}

		/// <summary>
		/// Returns a list of all open children of this bookmark. Mostly for internal use.
		/// </summary>
		/// <returns></returns>
        public int AllOpenCount()
		{
			if (ChildrenCollapsed) return 1;
            return 1 + FChildren.AllOpenCount();
		}

		#endregion

		#region ICloneable Members

		/// <summary>
		/// Returns a deep copy of this object.
		/// </summary>
		/// <returns></returns>
        public object Clone()
		{
			TBookmark Result = (TBookmark)MemberwiseClone();
			if (FDestination != null) Result.FDestination = (TPdfDestination)FDestination.Clone();
			Result.FChildren = (TBookmarkList)FChildren.Clone();
			return Result;
		}

		#endregion
	}

	/// <summary>
	/// Zoom options for a PDF destination.
	/// </summary>
	public enum TZoomOptions
	{
		/// <summary>
		/// None, leave the zoom unchanged.
		/// </summary>
		None,

		/// <summary>
		/// Display the page with its contents magnified just enough
		/// to fit the entire page within the window both horizontally and vertically. If
	    /// the required horizontal and vertical magnification factors are different, use
	    /// the smaller of the two, centering the page within the window in the other
	    /// dimension.
		/// </summary>
		Fit,

		/// <summary>
		/// Display the page with the vertical coordinate top positioned
		/// at the top edge of the window and the contents of the page magnified
	    /// just enough to fit the entire width of the page within the window.
		/// </summary>
		FitH,

		/// <summary>
		/// Display the page with the horizontal coordinate left positioned
		/// at the left edge of the window and the contents of the page magnified
	    /// just enough to fit the entire height of the page within the window.
		/// </summary>
		FitV
	}

	/// <summary>
	/// Represents a destination inside a PDF document.
	/// </summary>
	public class TPdfDestination: ICloneable
	{
		private int FPageNumber;
		private TZoomOptions FZoomOptions;

		/// <summary>
		/// Creates a new TPdfDestination instance.
		/// </summary>
		/// <param name="aPageNumber">Page where the destination points to. (1 based)</param>
		public TPdfDestination(int aPageNumber)
		{
			PageNumber = aPageNumber;
		}

		/// <summary>
		/// Creates a new TPdfDestination instance.
		/// </summary>
		/// <param name="aPageNumber">Page where the destination points to. (1 based)</param>
		/// <param name="aZoomOptions">Zoom options for this destination.</param>
		public TPdfDestination(int aPageNumber, TZoomOptions aZoomOptions) : this(aPageNumber)
		{
			FZoomOptions = aZoomOptions;
		}


		#region ICloneable Members

        /// <summary>
        /// Returns a deep copy of this object.
        /// </summary>
        /// <returns></returns>
		public object Clone()
		{
			return MemberwiseClone();
		}

		/// <summary>
		/// Page where the destination will point to. (1 based)
		/// </summary>
		public int PageNumber {get {return FPageNumber;} set{FPageNumber = value;}}

		/// <summary>
		/// Zoom options for this destination.
		/// </summary>
		public TZoomOptions ZoomOptions {get {return FZoomOptions;} set{FZoomOptions = value;}}

		#endregion
	}

	/// <summary>
	/// A list of bookmarks.
	/// </summary>
	public class TBookmarkList: IEnumerable, ICloneable
#if (FRAMEWORK20 && !DELPHIWIN32)
, IEnumerable<TBookmark>	
#endif
	{
#if (FRAMEWORK20)
        private List<TBookmark> FList;
#else
        private ArrayList FList;
#endif
		/// <summary>
		/// Creates a new instance of TBookmarkList.
		/// </summary>
		public TBookmarkList()
		{
#if (FRAMEWORK20)
            FList = new List<TBookmark>();
#else
			FList = new ArrayList();
#endif
        }

		/// <summary>
		/// Adds a new bookmark to the list.
		/// </summary>
		/// <param name="bookmark"></param>
		public TBookmark Add(TBookmark bookmark)
		{
			FList.Add(bookmark);
			return bookmark;
		}

		/// <summary>
		/// Removes all items on the list.
		/// </summary>
		public void Clear()
		{
			FList.Clear();
		}

		/// <summary>
		/// Returns item at position index on the list.
		/// </summary>
		public TBookmark this[int index]
		{
			get
			{
				return (TBookmark) FList[index];
			}
			set
			{
				FList[index] = value;
			}
		}

		/// <summary>
		/// Number of items on the list.
		/// </summary>
		public int Count
		{
			get
			{
				return FList.Count;
			}
		}

		/// <summary>
		/// Removes item at position index on the list.
		/// </summary>
		/// <param name="index">Position where to remove the bookmark.</param>
		public void RemoveAt (int index)
		{
			FList.RemoveAt(index);
		}

		/// <summary>
		/// Returns the count of all open bookmarks in all levels.
		/// </summary>
		/// <returns></returns>
		public int AllOpenCount()
		{
			int Result = 0;
			for (int i = FList.Count - 1; i >=0; i--)
			{
				Result += this[i].AllOpenCount();
			}
			return Result;
		}


		#region IEnumerable Members

        /// <summary>
        /// An enumerator for this collection.
        /// </summary>
        /// <returns>An enumerator for this collection.</returns>
		public IEnumerator GetEnumerator()
		{
			return FList.GetEnumerator();
		}

		#endregion

		#region ICloneable Members

        /// <summary>
        /// Returns a deep copy of this object.
        /// </summary>
        /// <returns></returns>
		public object Clone()
		{
			TBookmarkList Result = new TBookmarkList();
			for (int i = 0; i < Count; i++)
			{
				Result.Add((TBookmark) this[i].Clone());
			}

			return Result;
		}

		#endregion

        #region IEnumerable<TBookmark> Members
#if (FRAMEWORK20 && !DELPHIWIN32)
        IEnumerator<TBookmark> IEnumerable<TBookmark>.GetEnumerator()
        {
            return FList.GetEnumerator();
        }
#endif
        #endregion

    }
}
