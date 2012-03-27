using System;
using System.Collections.Generic;
using System.Globalization;

namespace FlexCel.Core
{
    #region ShapeGeom
    class TShapeGeom
    {
        internal string Name;
        internal TShapeGuideList AvList;
        internal TShapeGuideList GdList;
        internal TShapeAdjustHandleList AhList;
        internal TShapeConnectionList ConnList;
        internal TShapeTextRect TextRect;
        internal TShapePathList PathList;

        public TShapeGeom(string aName)
        {
            Name = aName;
            AvList = new TShapeGuideList();
            GdList = new TShapeGuideList();
            AhList = new TShapeAdjustHandleList();
            ConnList = new TShapeConnectionList();
            PathList = new TShapePathList();
        }

        internal bool FindGuide(string name, out TShapeGuide ResultGuide)
        {
#if (FRAMEWORK30)
            ResultGuide = GdList.Find((x) => x != null && x.Name == name);
            if (ResultGuide != null) return true;
            //Not even sure Av is valid in paths, but in any case they are searched later.
            ResultGuide = AvList.Find((x) => x != null && x.Name == name);
            if (ResultGuide != null) return true;

            if (TShapePresetGuides.FindGuide(name, out ResultGuide)) return true; //After used defined. A var can be defined in both places (see teardrop and r2)
#endif
            ResultGuide = null;
            return false;
        }

        internal TShapeGeom Clone()
        {
            return Clone(AvList);
        }

        internal TShapeGeom Clone(TShapeGuideList NewAv)
        {
            TShapeGeom Result = new TShapeGeom(Name);
            foreach (TShapeGuide guide in AvList)
            {
                TShapeGuide FinalGuide = guide;
                if (AvList != NewAv)
                {
                    TShapeGuide NewGuide = FindGuide(NewAv, guide.Name);
                    if (NewGuide != null) FinalGuide = NewGuide;
                }
                Result.AvList.Add(FinalGuide.Clone(Result));
            }

            foreach (TShapeGuide guide in GdList)
            {
                Result.GdList.Add(guide.Clone(Result));
            }

            foreach (TShapeAdjustHandle ah in AhList)
            {
                Result.AhList.Add(ah.Clone(Result));
            }

            foreach (TShapeConnection conn in ConnList)
            {
                Result.ConnList.Add(conn.Clone(Result));
            }

            if (TextRect != null) Result.TextRect = TextRect.Clone(Result);

            foreach (TShapePath path in PathList)
            {
                Result.PathList.Add(path.Clone(Result));
            }

            return Result;
        }

        private TShapeGuide FindGuide(TShapeGuideList GuideList, string GuideName)
        {
            if (GuideName == null) return null;

            foreach (TShapeGuide guide in GuideList)
            {
                if (guide.Name == GuideName) return guide;
            }
            return null;
        }

        public override bool Equals(object obj)
        {
            TShapeGeom o2 = obj as TShapeGeom;
            if (o2 == null) return false;
            return
                object.Equals(AvList, o2.AvList) &&
                object.Equals(GdList, o2.GdList) &&
                object.Equals(AhList, o2.AhList) &&
                object.Equals(ConnList, o2.ConnList) &&
                object.Equals(TextRect, o2.TextRect) &&
                object.Equals(PathList, o2.PathList);
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(Name, PathList);
        }

    }
    #endregion

    #region Path
    enum TPathFillMode
    {
        Norm,
        None,
        Lighten,
        LightenLess,
        Darken,
        DarkenLess
    }

    class TShapePathList : List<TShapePath>
    {
        public override bool Equals(object obj)
        {
            TShapePathList o2 = obj as TShapePathList;
            if (o2 == null) return false;

            if (o2.Count != Count) return false;
            for (int i = 0; i < Count; i++)
            {
                if (!Object.Equals(this[i], o2[i])) return false;
            }
            return true;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }

    class TShapePath
    {
        internal bool ExtrusionOk;
        internal TPathFillMode PathFill;
        internal bool PathStroke;
        internal int Width;
        internal int Height;

        internal List<TShapeAction> Actions;

        public TShapePath()
        {
            Actions = new List<TShapeAction>();
        }

        #region Equals
        public override bool Equals(object obj)
        {
            TShapePath o2 = obj as TShapePath;
            if (o2 == null) return false;

            if (!
                ExtrusionOk == o2.ExtrusionOk &&
                PathFill == o2.PathFill &&
                PathStroke == o2.PathStroke &&
                Width == o2.Width &&
                Height == o2.Height
                ) return false;

            if (Actions.Count != o2.Actions.Count) return false;
            for (int i = 0; i < Actions.Count; i++)
            {
                if (!Object.Equals(Actions[i], o2.Actions[i])) return false;
            }
            return true;

        }

        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(ExtrusionOk, PathFill, PathStroke, Width, Height, Actions);
        }

        public TShapePath Clone(TShapeGeom TargetShape)
        {
            TShapePath Result = new TShapePath();
            Result.ExtrusionOk = ExtrusionOk;
            Result.PathFill = PathFill;
            Result.PathStroke = PathStroke;
            Result.Width = Width;
            Result.Height = Height;
            if (Actions != null)
            {
                foreach (TShapeAction action in Actions)
                {
                    Result.Actions.Add(action.Clone(TargetShape));
                }
            }
            return Result;
        }
        #endregion
    }


    #region Shape Actions
    enum TShapeActionType
    {
        Close,
        MoveTo,
        LineTo,
        ArcTo,
        CubicBezierTo,
        QuadBezierTo
    }

    abstract class TShapeAction
    {
        internal TShapeActionType ActionType;

        public abstract TShapeAction Clone(TShapeGeom TargetShape);
    }

    class TShapeActionClose : TShapeAction
    {
        internal TShapeActionClose()
        {
            ActionType = TShapeActionType.Close;
        }

        #region Equals
        public override bool Equals(object obj)
        {
            return obj is TShapeActionClose;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public override TShapeAction Clone(TShapeGeom TargetShape)
        {
            return new TShapeActionClose();
        }
        #endregion
    }

    class TShapeActionMoveTo : TShapeAction
    {
        internal TShapePoint Target;

        internal TShapeActionMoveTo(TShapePoint aTarget)
        {
            ActionType = TShapeActionType.MoveTo;
            Target = aTarget;
        }


        #region Equals
        public override bool Equals(object obj)
        {
            TShapeActionMoveTo o2 = obj as TShapeActionMoveTo;
            if (o2 == null) return false;
            return o2.Target == Target;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(Target);
        }

        public override TShapeAction Clone(TShapeGeom TargetShape)
        {
            return new TShapeActionMoveTo(Target.Clone(TargetShape)); 
        }
        #endregion

    }

    class TShapeActionLineTo : TShapeAction
    {
        internal TShapePoint Target;

        internal TShapeActionLineTo(TShapePoint aTarget)
        {
            ActionType = TShapeActionType.LineTo;
            Target = aTarget;
        }

        #region Equals
        public override bool Equals(object obj)
        {
            TShapeActionLineTo o2 = obj as TShapeActionLineTo;
            if (o2 == null) return false;
            return o2.Target == Target;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(Target);
        }

        public override TShapeAction Clone(TShapeGeom TargetShape)
        {
            return new TShapeActionLineTo(Target.Clone(TargetShape));
        }
        #endregion

    }

    class TShapeActionArcTo : TShapeAction
    {
        internal TShapeGuide HeightRadius;
        internal TShapeGuide WidthRadius;
        internal TShapeGuide StartAngle;
        internal TShapeGuide SwingAngle;

        internal TShapeActionArcTo()
        {
            ActionType = TShapeActionType.ArcTo;
        }

        #region Equals
        public override bool Equals(object obj)
        {
            TShapeActionArcTo o2 = obj as TShapeActionArcTo;
            if (o2 == null) return false;
            if (!object.Equals(HeightRadius, o2.HeightRadius)) return false;
            if (!object.Equals(WidthRadius, o2.WidthRadius)) return false;
            if (!object.Equals(StartAngle, o2.StartAngle)) return false;
            if (!object.Equals(SwingAngle, o2.SwingAngle)) return false;

            return true;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(HeightRadius, WidthRadius, StartAngle, SwingAngle);
        }

        public override TShapeAction Clone(TShapeGeom TargetShape)
        {
            TShapeActionArcTo Result = new TShapeActionArcTo();
            if (HeightRadius != null) Result.HeightRadius = HeightRadius.Clone(TargetShape);
            if (WidthRadius != null) Result.WidthRadius = WidthRadius.Clone(TargetShape);
            if (StartAngle != null) Result.StartAngle = StartAngle.Clone(TargetShape);
            if (SwingAngle != null) Result.SwingAngle = SwingAngle.Clone(TargetShape);
            return Result;
        }
        #endregion

    }

    abstract class TShapeActionBezierTo : TShapeAction
    {
        internal TShapePoint[] Target;

        internal TShapeActionBezierTo(TShapePoint[] aTarget)
        {
            Target = aTarget;
        }

        #region Equals
        public override bool Equals(object obj)
        {
            TShapeActionBezierTo o2 = obj as TShapeActionBezierTo;
            if (o2 == null || ActionType != o2.ActionType) return false;
            if (o2.Target == null) return Target == null;
            if (Target == null) return false;

            if (Target.Length != o2.Target.Length) return false;
            for (int i = 0; i < Target.Length; i++)
            {
                if (Target[i] != o2.Target[i]) return false;
            }
            return true;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(Target);
        }

        internal TShapeActionBezierTo CloneTo(TShapeActionBezierTo Bez, TShapeGeom TargetShape)
        {
            if (Target != null)
            {
                Bez.Target = new TShapePoint[Target.Length];
                for (int i = 0; i < Target.Length; i++)
                {
                    Bez.Target[i] = Target[i].Clone(TargetShape);
                }
            }

            return Bez;
        }
        #endregion

    }

    class TShapeActionCubicBezierTo : TShapeActionBezierTo
    {
        internal TShapeActionCubicBezierTo(TShapePoint[] aTarget): base(aTarget)
        {
            ActionType = TShapeActionType.CubicBezierTo;
        }

        public override TShapeAction Clone(TShapeGeom TargetShape)
        {
            return CloneTo(new TShapeActionCubicBezierTo(null), TargetShape);
        }
    }

    class TShapeActionQuadBezierTo : TShapeActionBezierTo
    {
        internal TShapeActionQuadBezierTo(TShapePoint[] aTarget): base(aTarget)
        {
            ActionType = TShapeActionType.QuadBezierTo;
        }

        public override TShapeAction Clone(TShapeGeom TargetShape)
        {
            return CloneTo(new TShapeActionQuadBezierTo(null), TargetShape);
        }

    }
    #endregion

    struct TShapePoint
    {
        internal TShapeGuide x;
        internal TShapeGuide y;

        internal TShapePoint(TShapeGuide ax, TShapeGuide ay)
        {
            x = ax;
            y = ay;
        }

        internal TShapePoint Clone(TShapeGeom TargetShape)
        {
            TShapePoint Result = new TShapePoint();
            if (x != null) Result.x = x.Clone(TargetShape);
            if (y != null) Result.y = y.Clone(TargetShape);

            return Result;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is TShapePoint)) return false;
            return ((TShapePoint) obj) == this;
        }

        public static bool operator ==(TShapePoint o1, TShapePoint o2)
        {
            return Object.Equals(o1.x, o2.x) && Object.Equals(o1.y, o2.y);
        }

        public static bool operator !=(TShapePoint o1, TShapePoint o2)
        {
            return !(o1 == o2);
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(x, y);
        }

    }
    #endregion

    #region Guides
    class TShapeGuideList : List<TShapeGuide>
    {
        public override bool Equals(object obj)
        {
            TShapeGuideList o2 = obj as TShapeGuideList;
            if (o2 == null) return false;

            if (o2.Count != Count) return false;
            for (int i = 0; i < Count; i++)
            {
                if (!Object.Equals(this[i], o2[i])) return false; 
            }
            return true;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }

    class TShapeGuide
    {
        internal string Name;
        internal TShapeFormula Fmla;
        internal double CachedValue; //Shapes like leftRightCircularArrow are so complex that without this caching it could stay 1/2 a minute calculating the points.
        internal Guid CachedID;

        internal TShapeGuide(string aName, TShapeFormula aFmla)
        {
            Name = aName;
            Fmla = aFmla;
        }

        internal double Value(int level, TDrawingRelativeRect bounds, Guid aCachedID) 
        { 
            if (CachedID == aCachedID) return CachedValue;
            CachedValue = Fmla.GetValue(level, bounds, aCachedID);
            CachedID = aCachedID;
            return CachedValue;
        }

        internal double AbsValueInPoints(TDrawingRelativeRect bounds, Guid aCachedID, double PathWidthOrHeight, double BoundsWidthOrHeight)
        {
            double m = Value(0, bounds, aCachedID);
            if (PathWidthOrHeight != 0) m = BoundsWidthOrHeight * m / PathWidthOrHeight;
            return m / TDrawingCoordinate.PointsToEmu;
        }

        internal double XInPoints(TDrawingRelativeRect bounds, Guid aCachedID, double PathWidth)
        {
            double m = Value(0, bounds, aCachedID);
            if (PathWidth != 0) m = bounds.Width * m / PathWidth;
            return (bounds.Left + m) / TDrawingCoordinate.PointsToEmu;
        }

        internal double YInPoints(TDrawingRelativeRect bounds, Guid aCachedID, double PathHeight)
        {
            double m = Value(0, bounds, aCachedID);
            if (PathHeight != 0) m = bounds.Height * m / PathHeight;
            return (bounds.Top + m) / TDrawingCoordinate.PointsToEmu;
        }

        internal string NameOrValue()
        {
            if (Name != null) return Name;
            return Convert.ToString(Fmla.GetValue(0, new TDrawingRelativeRect(0, 0, 0, 0), Guid.NewGuid()), CultureInfo.InvariantCulture);
        }

        public override string ToString()
        {
            return "["+ Name + ": (" + Fmla.ToString() + ")] ";
        }

        #region Equals Boilerplate
        public override bool Equals(object obj)
        {
            TShapeGuide o2 = obj as TShapeGuide;
            if (o2 == null) return false;

            if (!Object.Equals(Name, o2.Name)) return false;
            if (!Object.Equals(Fmla, o2.Fmla)) return false;

            return true;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public TShapeGuide Clone(TShapeGeom TargetShape)
        {
            TShapeGuide Result;
            if (Name != null && TargetShape.FindGuide(Name, out Result)) return Result;
            return new TShapeGuide(Name, Fmla == null ? null : Fmla.Clone(TargetShape));
        }
        #endregion
    }
    #endregion

    #region Formulas
    abstract class TShapeFormula
    {
        protected abstract double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID);

        public double GetValue(int level, TDrawingRelativeRect bounds, Guid CacheId)
        {
            if (level > 1024) return 0; //avoid infinite recursion.
            return InternalVal(level + 1, bounds, CacheId);
        }

        protected double Angle60000(double p)
        {
            return p * 180 / Math.PI * 60000;
        }

        protected double AngleRad(double p)
        {
            return p / 180 * Math.PI / 60000;
        }

        internal string XlsxString()
        {
            return XlsxName() + " " + XlsxArgs();
        }

        internal abstract string XlsxName();
        internal abstract string XlsxArgs();

        public abstract TShapeFormula Clone(TShapeGeom TargetShape);

    }

    abstract class T1ArgShapeFormula : TShapeFormula
    {
        internal TShapeGuide x;

        public override string ToString()
        {
            return this.GetType().Name + "{" + x.ToString() + "}";
        }

        internal override string XlsxArgs()
        {
            return x.NameOrValue();
        }

        #region Equals Boilerplate
        public override bool Equals(object obj)
        {
            T1ArgShapeFormula o2 = obj as T1ArgShapeFormula;
            if (o2 == null) return false;

            if (!Object.Equals(x, o2.x)) return false;
            return true;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(x);
        }

        public T1ArgShapeFormula CloneFmla(T1ArgShapeFormula NewFmla, TShapeGeom TargetShape)
        {
            NewFmla.x = x.Clone(TargetShape);
            return NewFmla;
        }

        #endregion

    }

    abstract class T2ArgShapeFormula : T1ArgShapeFormula
    {
        internal TShapeGuide y;

        public override string ToString()
        {
            return this.GetType().Name + "{" + x.ToString() + ", " + y.ToString() + "}";
        }

        internal override string XlsxArgs()
        {
            return base.XlsxArgs() + " " + y.NameOrValue();
        }

        #region Equals Boilerplate
        public override bool Equals(object obj)
        {
            T2ArgShapeFormula o2 = obj as T2ArgShapeFormula;
            if (o2 == null) return false;

            if (!Object.Equals(x, o2.x)) return false;
            if (!Object.Equals(y, o2.y)) return false;
            return true;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(x, y);
        }

        public T2ArgShapeFormula CloneFmla(T2ArgShapeFormula NewFmla, TShapeGeom TargetShape)
        {
            NewFmla.x = x.Clone(TargetShape);
            NewFmla.y = y.Clone(TargetShape);

            return NewFmla;
        }


        #endregion
    }

    abstract class T3ArgShapeFormula : T2ArgShapeFormula
    {
        internal TShapeGuide z;

        public override string ToString()
        {
            return this.GetType().Name + "{" + x.ToString() + ", " + y.ToString() + ", " + z.ToString() + "}";
        }

        internal override string XlsxArgs()
        {
            return base.XlsxArgs() + " " + z.NameOrValue();
        }

        #region Equals Boilerplate
        public override bool Equals(object obj)
        {
            T3ArgShapeFormula o2 = obj as T3ArgShapeFormula;
            if (o2 == null) return false;

            if (!Object.Equals(x, o2.x)) return false;
            if (!Object.Equals(y, o2.y)) return false;
            if (!Object.Equals(z, o2.z)) return false;
            return true;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(x, y, z);
        }

        public T3ArgShapeFormula CloneFmla(T3ArgShapeFormula NewFmla, TShapeGeom TargetShape)
        {
            NewFmla.x = x.Clone(TargetShape);
            NewFmla.y = y.Clone(TargetShape);
            NewFmla.z = z.Clone(TargetShape);

            return NewFmla;
        }
        #endregion

    }

    class TShapeMulDiv : T3ArgShapeFormula
    {
        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            return x.Value(level, bounds, CacheID) * y.Value(level, bounds, CacheID) / z.Value(level, bounds, CacheID);
        }

        internal override string XlsxName()
        {
            return "*/";
        }

        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return CloneFmla(new TShapeMulDiv(), TargetShape);
        }
    }

    class TShapeAddSub : T3ArgShapeFormula
    {
        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            return x.Value(level, bounds, CacheID) + y.Value(level, bounds, CacheID) - z.Value(level, bounds, CacheID);
        }

        internal override string XlsxName()
        {
            return "+-";
        }

        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return CloneFmla(new TShapeAddSub(), TargetShape);
        }

    }

    class TShapeAddDiv : T3ArgShapeFormula
    {
        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            return (x.Value(level, bounds, CacheID) + y.Value(level, bounds, CacheID)) / z.Value(level, bounds, CacheID);
        }
       
        internal override string XlsxName()
        {
            return "+/";
        }

        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return CloneFmla(new TShapeAddDiv(), TargetShape);
        }

    }

    class TShapeIfElse : T3ArgShapeFormula
    {
        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            if (x.Value(level, bounds, CacheID) > 0)  return y.Value(level, bounds, CacheID); else return z.Value(level, bounds, CacheID);
        }
    
        internal override string XlsxName()
        {
            return "?:";
        }

        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return CloneFmla(new TShapeIfElse(), TargetShape);
        }

    }

    class TShapeAbs : T1ArgShapeFormula
    {
        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            return Math.Abs(x.Value(level, bounds, CacheID));
        }

        internal override string XlsxName()
        {
            return "abs";
        }

        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return CloneFmla(new TShapeAbs(), TargetShape);
        }

    }

    class TShapeArcTan : T2ArgShapeFormula
    {
        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            return Angle60000(Math.Atan2(y.Value(level, bounds, CacheID), x.Value(level, bounds, CacheID)));
        }
    
        internal override string XlsxName()
        {
            return "at2";
        }

        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return CloneFmla(new TShapeArcTan(), TargetShape);
        }

    }

    class TShapeCosArcTan : T3ArgShapeFormula
    {
        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            return x.Value(level, bounds, CacheID) * Math.Cos(Math.Atan2(z.Value(level, bounds, CacheID), y.Value(level, bounds, CacheID)));
        }

        internal override string XlsxName()
        {
            return "cat2";
        }

        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return CloneFmla(new TShapeCosArcTan(), TargetShape);
        }

    }

    class TShapeCos : T2ArgShapeFormula
    {
        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            return x.Value(level, bounds, CacheID) * Math.Cos(AngleRad(y.Value(level, bounds, CacheID)));
        }

        internal override string XlsxName()
        {
            return "cos";
        }

        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return CloneFmla(new TShapeCos(), TargetShape);
        }


    }

    class TShapeMax : T2ArgShapeFormula
    {
        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            return Math.Max(x.Value(level, bounds, CacheID), y.Value(level, bounds, CacheID));
        }

        internal override string XlsxName()
        {
            return "max";
        }

        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return CloneFmla(new TShapeMax(), TargetShape);
        }


    }

    class TShapeMin : T2ArgShapeFormula
    {
        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            return Math.Min(x.Value(level, bounds, CacheID), y.Value(level, bounds, CacheID));
        }

        internal override string XlsxName()
        {
            return "min";
        }

        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return CloneFmla(new TShapeMin(), TargetShape);
        }

    }

    class TShapeMod : T3ArgShapeFormula
    {
        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            return Math.Sqrt(Math.Pow(x.Value(level, bounds, CacheID), 2) + Math.Pow(y.Value(level, bounds, CacheID), 2) + Math.Pow(z.Value(level, bounds, CacheID), 2));
        }

        internal override string XlsxName()
        {
            return "mod";
        }

        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return CloneFmla(new TShapeMod(), TargetShape);
        }


    }

    class TShapePin : T3ArgShapeFormula
    {
        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            double xv = x.Value(level, bounds, CacheID);
            double yv = y.Value(level, bounds, CacheID);
            double zv = z.Value(level, bounds, CacheID);
            if (yv < xv) return xv;
            if (yv > zv) return zv;
            return yv;
        }

        internal override string XlsxName()
        {
            return "pin";
        }

        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return CloneFmla(new TShapePin(), TargetShape);
        }

    }

    class TShapeSinArcTan: T3ArgShapeFormula
    {
        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            return x.Value(level, bounds, CacheID) * Math.Sin(Math.Atan2(z.Value(level, bounds, CacheID), y.Value(level, bounds, CacheID)));
        }

        internal override string XlsxName()
        {
            return "sat2";
        }

        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return CloneFmla(new TShapeSinArcTan(), TargetShape);
        }

    }

    class TShapeSin : T2ArgShapeFormula
    {
        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            return x.Value(level, bounds, CacheID) * Math.Sin(AngleRad(y.Value(level, bounds, CacheID)));
        }

        internal override string XlsxName()
        {
            return "sin";
        }
        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return CloneFmla(new TShapeSin(), TargetShape);
        }

    }

    class TShapeSqrt : T1ArgShapeFormula
    {
        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            return Math.Sqrt(x.Value(level, bounds, CacheID));
        }

        internal override string XlsxName()
        {
            return "sqrt";
        }

        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return CloneFmla(new TShapeSqrt(), TargetShape);
        }

    }

    class TShapeTan : T2ArgShapeFormula
    {
        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            return x.Value(level, bounds, CacheID) * Math.Tan(AngleRad(y.Value(level, bounds, CacheID)));
        }

        internal override string XlsxName()
        {
            return "tan";
        }

        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return CloneFmla(new TShapeTan(), TargetShape);
        }

    }

    class TShapeVal : TShapeFormula
    {
        double Val;

        internal TShapeVal(double aVal)
        {
            Val = aVal;
        }

        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            return Val;
        }

        public double ConstantVal
        {
            get { return Val; }
        }   

        public override string ToString()
        {
            return Val.ToString();
        }

        internal override string XlsxName()
        {
            return "val";
        }

        internal override string XlsxArgs()
        {
            return Convert.ToString(Val, CultureInfo.InvariantCulture);
        }

        #region Equals Boilerplate
        public override bool Equals(object obj)
        {
            TShapeVal o2 = obj as TShapeVal;
            if (o2 == null) return false;

            if (!Object.Equals(Val, o2.Val)) return false;
            return true;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(Val);
        }

        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return new TShapeVal(Val);
        }
        #endregion

    }

    class TShapeUndefFormula : TShapeFormula
    {
        string Fmla;

        internal TShapeUndefFormula(string aFmla)
        {
            Fmla = aFmla;
        }

        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            return 0;
        }

        internal override string XlsxName()
        {
            return "";
        }

        internal override string XlsxArgs()
        {
            return Fmla;
        }

        #region Equals Boilerplate
        public override bool Equals(object obj)
        {
            TShapeUndefFormula o2 = obj as TShapeUndefFormula;
            if (o2 == null) return false;

            if (!Object.Equals(Fmla, o2.Fmla)) return false;
            return true;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(Fmla);
        }

        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return new TShapeUndefFormula(Fmla);
        }

        #endregion

    }

    #endregion

    #region Rect
    class TShapeTextRect
    {
        internal TShapeGuide Left;
        internal TShapeGuide Top;
        internal TShapeGuide Right;
        internal TShapeGuide Bottom;

        #region Equals Boilerplate
        public override bool Equals(object obj)
        {
            TShapeTextRect o2 = obj as TShapeTextRect;
            if (o2 == null) return false;

            if (!Object.Equals(Left, o2.Left)) return false;
            if (!Object.Equals(Top, o2.Top)) return false;
            if (!Object.Equals(Right, o2.Right)) return false;
            if (!Object.Equals(Bottom, o2.Bottom)) return false;
            return true;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(Left, Top, Right, Bottom);
        }

        public TShapeTextRect Clone(TShapeGeom TargetShape)
        {
            TShapeTextRect Result = new TShapeTextRect();
            if (Left != null) Result.Left = Left.Clone(TargetShape);
            if (Top != null) Result.Top = Top.Clone(TargetShape);
            if (Right != null) Result.Right = Right.Clone(TargetShape);
            if (Bottom != null) Result.Bottom = Bottom.Clone(TargetShape);

            return Result;
        }

        #endregion

    }
    #endregion


    #region Adjust Handles
    class TShapeAdjustHandleList: List<TShapeAdjustHandle>
    {
        public override bool Equals(object obj)
        {
            TShapeAdjustHandleList o2 = obj as TShapeAdjustHandleList;
            if (o2 == null) return false;

            if (o2.Count != Count) return false;
            for (int i = 0; i < Count; i++)
            {
                if (!Object.Equals(this[i], o2[i])) return false;
            }
            return true;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

    }

    abstract class TShapeAdjustHandle
    {
        internal TShapePoint Location;

        internal abstract TShapeAdjustHandle Clone(TShapeGeom TargetShape);
    }

    class TShapeAdjustHandlePolar: TShapeAdjustHandle
    {
        internal TShapeGuide GdRefR;
        internal TShapeGuide GdRefAng;

        internal TShapeGuide MinR;
        internal TShapeGuide MaxR;
        internal TShapeGuide MinAng;
        internal TShapeGuide MaxAng;

        #region Equals
        public override bool Equals(object obj)
        {
            TShapeAdjustHandlePolar o2 = obj as TShapeAdjustHandlePolar;
            if (o2 == null) return false;
            if (!Object.Equals(Location, o2.Location)) return false;
            if (!Object.Equals(GdRefR, o2.GdRefR)) return false;
            if (!Object.Equals(GdRefAng, o2.GdRefAng)) return false;
            if (!Object.Equals(MinR, o2.MinR)) return false;
            if (!Object.Equals(MaxR, o2.MaxR)) return false;
            if (!Object.Equals(MinAng, o2.MinAng)) return false;
            if (!Object.Equals(MaxAng, o2.MaxAng)) return false;
            return true;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(Location, GdRefR, GdRefAng, MinR, MaxR, MinAng, MaxAng);
        }

        internal override TShapeAdjustHandle Clone(TShapeGeom TargetShape)
        {
            TShapeAdjustHandlePolar Result = new TShapeAdjustHandlePolar();
            Result.Location = Location.Clone(TargetShape);
            if (GdRefR != null) Result.GdRefR = GdRefR.Clone(TargetShape);
            if (GdRefAng != null) Result.GdRefAng = GdRefAng.Clone(TargetShape);
            if (MinR != null) Result.MinR = MinR.Clone(TargetShape);
            if (MaxR != null) Result.MaxR = MaxR.Clone(TargetShape);
            if (MinAng != null) Result.MinAng = MinAng.Clone(TargetShape);
            if (MaxAng != null) Result.MaxAng = MaxAng.Clone(TargetShape);

            return Result;
        }
        #endregion
    }
   
    class TShapeAdjustHandleXY: TShapeAdjustHandle
    {
        internal TShapeGuide GdRefX;
        internal TShapeGuide GdRefY;

        internal TShapeGuide MinX;
        internal TShapeGuide MaxX;
        internal TShapeGuide MinY;
        internal TShapeGuide MaxY;

        #region Equals
        public override bool Equals(object obj)
        {
            TShapeAdjustHandleXY o2 = obj as TShapeAdjustHandleXY;
            if (o2 == null) return false;
            if (!Object.Equals(Location, o2.Location)) return false;
            if (!Object.Equals(GdRefX, o2.GdRefX)) return false;
            if (!Object.Equals(GdRefY, o2.GdRefY)) return false;
            if (!Object.Equals(MinX, o2.MinX)) return false;
            if (!Object.Equals(MaxX, o2.MaxX)) return false;
            if (!Object.Equals(MinY, o2.MinY)) return false;
            if (!Object.Equals(MaxY, o2.MaxY)) return false;
            return true;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(Location, GdRefX, GdRefY, MinX, MaxX, MinY, MaxY);
        }

        internal override TShapeAdjustHandle Clone(TShapeGeom TargetShape)
        {
            TShapeAdjustHandleXY Result = new TShapeAdjustHandleXY();
            Result.Location = Location.Clone(TargetShape);
            if (GdRefX != null) Result.GdRefX = GdRefX.Clone(TargetShape);
            if (GdRefY != null) Result.GdRefY = GdRefY.Clone(TargetShape);
            if (MinX != null) Result.MinX = MinX.Clone(TargetShape);
            if (MaxX != null) Result.MaxX = MaxX.Clone(TargetShape);
            if (MinY != null) Result.MinY = MinY.Clone(TargetShape);
            if (MaxY != null) Result.MaxY = MaxY.Clone(TargetShape);

            return Result;
        }
        #endregion

    }

    #endregion

    #region Connections
    class TShapeConnectionList : List<TShapeConnection>
    {
        public override bool Equals(object obj)
        {
            TShapeConnectionList o2 = obj as TShapeConnectionList;
            if (o2 == null) return false;

            if (o2.Count != Count) return false;
            for (int i = 0; i < Count; i++)
            {
                if (!Object.Equals(this[i], o2[i])) return false;
            }
            return true;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

    }

    class TShapeConnection
    {
        internal TShapePoint Position;
        internal TShapeGuide Angle;

        #region Equals
        public override bool Equals(object obj)
        {
            TShapeConnection o2 = obj as TShapeConnection;
            if (o2 == null) return false;
            if (!Object.Equals(Position, o2.Position)) return false;
            if (!Object.Equals(Angle, o2.Angle)) return false;
            return true;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(Position, Angle);
        }

        internal TShapeConnection Clone(TShapeGeom TargetShape)
        {
            TShapeConnection Result = new TShapeConnection();
            Result.Position = Position.Clone(TargetShape);
            if (Angle != null) Result.Angle = Angle.Clone(TargetShape);

            return Result;
        }
        #endregion

    }
    #endregion

    #region Preset guides
    /// <summary>
    /// Internal use.
    /// </summary>
    public enum TShapePresetGuideType
    {
        /// <summary>
        /// m/n of a Circle ('mcdn') - 
        /// The units here are in 60,000ths of a degree. 
        /// </summary>
        Pr_cd,

        /// <summary>
        /// Shape Bottom Edge ('b') - Constant value of "h" 
        /// This is the bottom edge of the shape and since the top edge of the shape is considered the 0 point, the bottom edge is thus the shape height. 
        /// </summary>
        Pr_b,

        /// <summary>
        /// Shape Height ('h') 
        /// This is the variable height of the shape defined in the shape properties. This value is received from the shape transform listed within the spPr element. 
        /// </summary>
        Pr_h,

        /// <summary>
        /// Horizontal Center ('hc') - Calculated value of "*/ w 1.0 2.0" 
        /// This is the horizontal center of the shape which is just the width divided by 2. 
        /// </summary>
        Pr_hc,

        /// <summary>
        /// 1/n of Shape Height ('hdn') - Calculated value of "*/ h 1.0 n.0" 
        /// This is 1/n the shape height. 
        /// </summary>
        Pr_hd,

        /// <summary>
        /// Shape Left Edge ('l') - Constant value of "0" 
        /// This is the left edge of the shape and the left edge of the shape is considered the horizontal 0 point. 
        /// </summary>
        Pr_l,

        /// <summary>
        /// Longest Side of Shape ('ls') - Calculated value of "max w h" 
        /// This is the longest side of the shape. This value is either the width or the height depending on which is greater. 
        /// </summary>
        Pr_ls,

        /// <summary>
        /// Shape Right Edge ('r') - Constant value of "w" 
        /// This is the right edge of the shape and since the left edge of the shape is considered the 0 point, the right edge is thus the shape width.
        /// </summary>
        Pr_r,
 
        /// <summary>
        /// Shortest Side of Shape ('ss') - Calculated value of "min w h" 
        /// This is the shortest side of the shape. This value is either the width or the height depending on which is smaller. 
        /// </summary>
        Pr_ss,

        /// <summary>
        /// 1/n Shortest Side of Shape ('ssdn') - Calculated value of "*/ ss 1.0 n.0" 
        /// </summary>
        Pr_ssd,

        /// <summary>
        /// Shape Top Edge ('t') - Constant value of "0" 
        /// This is the top edge of the shape and the top edge of the shape is considered the vertical 0 point. 
        /// </summary>
        Pr_t,

        /// <summary>
        /// Vertical Center of Shape ('vc') - Calculated value of "*/ h 1.0 2.0" 
        /// This is the vertical center of the shape which is just the height divided by 2. 
        /// </summary>
        Pr_vc,

        /// <summary>
        /// Shape Width ('w') 
        /// This is the variable width of the shape defined in the shape properties. This value is received from the shape transform listed within the spPr element. 
        /// </summary>
        Pr_w,
 
        /// <summary>
        /// 1/n of Shape Width ('wdn') - Calculated value of "*/ w 1.0 n.0" 
        /// This is 1/n the shape width. 
        /// </summary>
        Pr_wd,
    }

    class TShapePresetFormula: TShapeFormula
    {
        internal TShapePresetGuideType Preset;
        double Mult;
        double Divv;

        internal TShapePresetFormula(TShapePresetGuideType aPreset, double aMult, double aDivv)
        {
            Preset = aPreset;
            Mult = aMult;
            Divv = aDivv;
        }

        protected override double InternalVal(int level, TDrawingRelativeRect bounds, Guid CacheID)
        {
            const double Circle = 21600000;
            double w = bounds.Right - bounds.Left;
            double h = bounds.Bottom - bounds.Top;
            switch (Preset)
            {
                case TShapePresetGuideType.Pr_cd:
                    return Circle * Mult / Divv;

                case TShapePresetGuideType.Pr_b:
                    return h * Mult / Divv;

                case TShapePresetGuideType.Pr_h:
                    return h * Mult / Divv;
                
                case TShapePresetGuideType.Pr_hc:
                    return w / 2 * Mult / Divv;
                
                case TShapePresetGuideType.Pr_hd:
                    return h * Mult / Divv;

                case TShapePresetGuideType.Pr_l:
                    return 0 * Mult / Divv;
                
                case TShapePresetGuideType.Pr_ls:
                    return Math.Max(w, h) * Mult / Divv;
                
                case TShapePresetGuideType.Pr_r:
                    return w * Mult / Divv;
                
                case TShapePresetGuideType.Pr_ss:
                    return Math.Min(w, h) * Mult / Divv;

                case TShapePresetGuideType.Pr_t:
                    return 0 * Mult / Divv;
                
                case TShapePresetGuideType.Pr_vc:
                    return h / 2 * Mult / Divv;

                case TShapePresetGuideType.Pr_w:
                    return w * Mult / Divv;

                case TShapePresetGuideType.Pr_wd:
                    return w * Mult / Divv;

                default:
                    return 0;
            }
        }

        internal override string XlsxName()
        {
            FlxMessages.ThrowException(FlxErr.ErrInternal); //Preset guides shouldn't be saved.
            return null;
        }

        internal override string XlsxArgs()
        {
            FlxMessages.ThrowException(FlxErr.ErrInternal); //Preset guides shouldn't be saved.
            return null;
        }

        #region Equals Boilerplate
        public override bool Equals(object obj)
        {
            TShapePresetFormula o2 = obj as TShapePresetFormula;
            if (o2 == null) return false;

            if (!Object.Equals(Preset, o2.Preset)) return false;
            if (!Object.Equals(Mult, o2.Mult)) return false;
            if (!Object.Equals(Divv, o2.Divv)) return false;
            return true;
        }

        public override int GetHashCode()
        {
            return HashCoder.GetHashObj(Preset, Mult, Divv);
        }

        public override TShapeFormula Clone(TShapeGeom TargetShape)
        {
            return new TShapePresetFormula(Preset, Mult, Divv);
        }
        #endregion

    }
    
    static class TShapePresetGuides
    {
        static Dictionary<string, TShapeGuide> Presets = GetPresets();

        private static Dictionary<string, TShapeGuide> GetPresets()
        {
            Dictionary<string, TShapeGuide> Result = new Dictionary<string, TShapeGuide>();

            foreach (TShapePresetGuideType gt in TCompactFramework.EnumGetValues(typeof(TShapePresetGuideType)))
            {
                string PresetName = GetPresetName(gt);
                Result.Add(PresetName, new TShapeGuide(PresetName, new TShapePresetFormula(gt, 1, 1))); 
            }

            return Result;
        }

        private static string GetPresetName(TShapePresetGuideType gt)
        {
            string Result = TCompactFramework.EnumGetName(typeof(TShapePresetGuideType), gt);
            return Result.Substring(3);
        }

        internal static bool FindGuide(string name, out TShapeGuide value)
        {
            double Divv;
            double Mult;
            string BaseName;
            CalcPresetMult(name, out BaseName, out Mult, out Divv);

            if (!Presets.TryGetValue(BaseName, out value)) return false;
            if (Divv != 1 || Mult != 1)
            {
                value = new TShapeGuide(name, new TShapePresetFormula((value.Fmla as TShapePresetFormula).Preset, Mult, Divv));
            }
            return true;
        }

        private static void CalcPresetMult(string name, out string BaseName, out double Mult, out double Divv)
        {
            Mult = 1;
            Divv = 1;
            BaseName = String.Empty;

            int start = 0;
            while (start < name.Length)
            {
                if (!Char.IsDigit(name[start])) break;
                start++;
            }
            if (start >= name.Length) return; //simple number

            int finish = name.Length - 1;
            while (finish >= 0)
            {
                if (!Char.IsDigit(name[finish])) break;
                finish--;
            }
            if (finish < start) return;

            if (start > 0) Mult = Convert.ToInt32(name.Substring(0, start), CultureInfo.InvariantCulture);
            if (finish < name.Length - 1) Divv = Convert.ToInt32(name.Substring(finish + 1), CultureInfo.InvariantCulture);
            BaseName = name.Substring(start, 1 + finish - start);

        }

    }
    
    #endregion
}
