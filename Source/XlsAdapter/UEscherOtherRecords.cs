using System;
using FlexCel.Core;

namespace FlexCel.XlsAdapter
{
	/// <summary>
	/// A record Group
	/// </summary>
	internal class TEscherRecordGroups: TEscherDataRecord
	{
		public TEscherRecordGroups(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
			: base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
		{
		}
	}

	/// <summary>
	/// Base record for all Rules.
	/// </summary>
	internal abstract class TRuleRecord: TEscherDataRecord
	{
		protected TRuleRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
			: base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
		{
		}

		internal abstract bool DeleteRef(TEscherSpRecord aShape);
		internal abstract void FixPointers();
		internal abstract void ArrangeCopyRange(int RowOfs, int ColOfs, TSheetInfo SheetInfo);

	}

	/// <summary>
	/// Connector Rule.
	/// </summary>
	internal class TEscherConnectorRuleRecord: TRuleRecord
	{
		private TEscherSpRecord[] Shapes;
		internal TEscherConnectorRuleRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
			: base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
		{
			Init();
		}

        private void Init()
        {
            Shapes=new TEscherSpRecord[3];
        }

        protected override TEscherRecord DoCopyTo(int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
        {
            TEscherConnectorRuleRecord Result= (TEscherConnectorRuleRecord) base.DoCopyTo(RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo);
            Result.Init();
            for (int c=0;c<3;c++)
                if (Shapes[c] !=null)
                {
                    Result.Shapes[c] = (TEscherSpRecord)Shapes[c].CopiedTo(SheetInfo.CopiedGen);
                    if (Result.Shapes[c]!=null)  Result.SetSpIds(c, Result.Shapes[c].ShapeId); else Result.SetSpIds(c,0);
                }
            Result.RuleId= DwgCache.Solver.IncMaxRuleId();
            return Result;
        }
		
		
		internal long RuleId {get {return BitOps.GetCardinal(Data,0);} set {BitOps.SetCardinal(Data,0,value);}}
		internal long SpIds(int c){return BitOps.GetCardinal(Data, 4+c*4);}
		internal void SetSpIds(int c, long value){ BitOps.SetCardinal(Data, 4+c*4, value);}
		internal long CpA {get {return BitOps.GetCardinal(Data, 4+3*4);}}
		internal long CpB {get {return BitOps.GetCardinal(Data, 4+3*4+4);}}



		internal override bool DeleteRef(TEscherSpRecord aShape)
		{
			for (int c=0; c<3 ;c++)
				if (Shapes[c]== aShape)
				{
					Shapes[c]=null;
					SetSpIds(c,0);
				}
			return Shapes[2]==null;
		}

        internal override void FixPointers()
        {
            if (DwgCache != null) DwgCache.Solver.CheckMax(RuleId);
            int Index = -1;
            for (int c = 0; c < 3; c++)
            {
                if (DwgCache.Shape.Find(SpIds(c), ref Index))
                {
                    Shapes[c] = DwgCache.Shape[Index];
                }
                else Shapes[c] = null;
            }
        }

		internal override void ArrangeCopyRange(int RowOfs, int ColOfs, TSheetInfo SheetInfo)
		{
            if ((Shapes[2] != null) && (Shapes[2].CopiedTo(SheetInfo.CopiedGen) != null))
				DwgCache.Solver.ContainedRecords.Add(TEscherConnectorRuleRecord.Clone(this, RowOfs, ColOfs, DwgCache, DwgGroupCache, SheetInfo));
		}
	}

/// <summary>
/// Align rule record. Not implemented yet.
/// </summary>
	internal class TEscherAlignRuleRecord: TRuleRecord
	{
		public TEscherAlignRuleRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
			: base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
		{
		}
		
		internal long RuleId {get {return BitOps.GetCardinal(Data,0);} set {BitOps.SetCardinal(Data,0,value);}}
		internal long Align {get {return BitOps.GetCardinal(Data, 4);}}
		internal long nProxies {get {return BitOps.GetCardinal(Data, 4+4);}}


		protected override TEscherRecord DoCopyTo(int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
		{
			XlsMessages.ThrowException(XlsErr.ErrNotImplemented, "Align Rule");
			return null;
		}
		
		internal override bool DeleteRef(TEscherSpRecord aShape)
		{
			XlsMessages.ThrowException(XlsErr.ErrNotImplemented, "Align Rule");
			return false;
		}

		internal override void FixPointers()
		{
			XlsMessages.ThrowException(XlsErr.ErrNotImplemented, "Align Rule");
		}

		internal override void ArrangeCopyRange(int RowOfs, int ColOfs, TSheetInfo SheetInfo)
		{
			XlsMessages.ThrowException(XlsErr.ErrNotImplemented, "Align Rule");
		}
	}

	/// <summary>
	/// Arc Rule.
	/// </summary>
	internal class TEscherArcRuleRecord: TRuleRecord
	{
		private TEscherSpRecord Shape;
		public TEscherArcRuleRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
			: base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
		{
		}
		
		internal long RuleId {get {return BitOps.GetCardinal(Data,0);} set {BitOps.SetCardinal(Data,0,value);}}
		internal long SpId {get {return BitOps.GetCardinal(Data, 4);} set {BitOps.SetCardinal(Data,4,value);}}

		protected override TEscherRecord DoCopyTo(int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
		{
			TEscherArcRuleRecord R= (TEscherArcRuleRecord) base.DoCopyTo(RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo);
				if (Shape !=null)
				{
					R.Shape= (TEscherSpRecord) Shape.CopiedTo(SheetInfo.CopiedGen);
					if (R.Shape!=null)  R.SpId= R.Shape.ShapeId; else R.SpId=0;
				}
			R.RuleId= DwgCache.Solver.IncMaxRuleId();
			return R;
		}
		
		internal override bool DeleteRef(TEscherSpRecord aShape)
		{
				if (Shape== aShape)
				{
					Shape=null;
					SpId=0;
				}
			return Shape==null;
		}

        internal override void FixPointers()
        {
            if (DwgCache != null) DwgCache.Solver.CheckMax(RuleId);
            int Index = -1;
            if (DwgCache.Shape.Find(SpId, ref Index))
            {
                Shape = DwgCache.Shape[Index];
            }
            else Shape = null;
        }

		internal override void ArrangeCopyRange(int RowOfs, int ColOfs, TSheetInfo SheetInfo)
		{
            if ((Shape != null) && (Shape.CopiedTo(SheetInfo.CopiedGen) != null))
				DwgCache.Solver.ContainedRecords.Add(TEscherArcRuleRecord.Clone(this, RowOfs, ColOfs, DwgCache, DwgGroupCache, SheetInfo));
		}
	}

	/// <summary>
	/// CallOut Rule.
	/// </summary>
	internal class TEscherCallOutRuleRecord: TRuleRecord
	{
		private TEscherSpRecord Shape;
		public TEscherCallOutRuleRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
			: base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
		{
		}
		
		internal long RuleId {get {return BitOps.GetCardinal(Data,0);} set {BitOps.SetCardinal(Data,0,value);}}
		internal long SpId {get {return BitOps.GetCardinal(Data, 4);} set {BitOps.SetCardinal(Data,4,value);}}

		protected override TEscherRecord DoCopyTo(int RowOfs, int ColOfs, TEscherDwgCache NewDwgCache, TEscherDwgGroupCache NewDwgGroupCache, TSheetInfo SheetInfo)
		{
			TEscherCallOutRuleRecord R= (TEscherCallOutRuleRecord) base.DoCopyTo(RowOfs, ColOfs, NewDwgCache, NewDwgGroupCache, SheetInfo);
			if (Shape !=null)
			{
                R.Shape = (TEscherSpRecord)Shape.CopiedTo(SheetInfo.CopiedGen);
				if (R.Shape!=null)  R.SpId= R.Shape.ShapeId; else R.SpId=0;
			}
			R.RuleId= DwgCache.Solver.IncMaxRuleId();
			return R;
		}
		
		internal override bool DeleteRef(TEscherSpRecord aShape)
		{
			if (Shape== aShape)
			{
				Shape=null;
				SpId=0;
			}
			return Shape==null;
		}

        internal override void FixPointers()
        {
            if (DwgCache != null) DwgCache.Solver.CheckMax(RuleId);
            int Index = -1;
            if (DwgCache.Shape.Find(SpId, ref Index))
            {
                Shape = DwgCache.Shape[Index];
            }
            else Shape = null;
        }

		internal override void ArrangeCopyRange(int RowOfs, int ColOfs, TSheetInfo SheetInfo)
		{
            if ((Shape != null) && (Shape.CopiedTo(SheetInfo.CopiedGen) != null))
				DwgCache.Solver.ContainedRecords.Add(TEscherCallOutRuleRecord.Clone(this, RowOfs, ColOfs, DwgCache, DwgGroupCache, SheetInfo));
		}
	}

	internal class TEscherClientTextBoxRecord: TEscherClientDataRecord
	{
		internal TEscherClientTextBoxRecord(TEscherRecordHeader aEscherHeader, TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
			: base(aEscherHeader, aDwgGroupCache, aDwgCache, aParent)
		{
		}

		/// <summary>
		/// Create from data.
		/// </summary>
		/// <param name="aDwgGroupCache"></param>
		/// <param name="aDwgCache"></param>
		/// <param name="aParent"></param>
		internal TEscherClientTextBoxRecord(TEscherDwgGroupCache aDwgGroupCache, TEscherDwgCache aDwgCache, TEscherContainerRecord aParent)
			: base(new TEscherRecordHeader(0,(int)Msofbt.ClientTextbox,0), aDwgGroupCache, aDwgCache, aParent)
		{
			LoadedDataSize=0;
		}

		internal TRichString GetValue()
        {
            return ((TTXO)ClientData).GetText();
        }
        
        internal void SetValue(TRichString value) 
        {
            ((TTXO)ClientData).SetText(value);
        }

        internal bool LockText
        {
            get
            {
                return (((TTXO)ClientData).OptionFlags & 0x0200) != 0;
            }
            set
            {
                if (value) ((TTXO)ClientData).OptionFlags |= 0x0200; else ((TTXO)ClientData).OptionFlags &= ~0x0200; 
            }
        }

        internal THFlxAlignment HAlign
        {
            get
            {
                switch ((((TTXO)ClientData).OptionFlags >> 1) & 0x07)
                {
                    case 2: return THFlxAlignment.center;
                    case 3: return THFlxAlignment.right;
                    case 4: return THFlxAlignment.justify;
                    case 7: return THFlxAlignment.distributed;
                    default:
                        return THFlxAlignment.left;
                }
            }
            set
            {
                ((TTXO)ClientData).OptionFlags &= ~0x0E; //Clear it

                switch (value)
                {
                    case THFlxAlignment.center:
                        ((TTXO)ClientData).OptionFlags |= (0x02 << 1);
                        break;

                    case THFlxAlignment.right:
                        ((TTXO)ClientData).OptionFlags |= (0x03 << 1);
                        break;
                    
                    case THFlxAlignment.justify:
                        ((TTXO)ClientData).OptionFlags |= (0x04 << 1);
                        break;

                    case THFlxAlignment.distributed:
                        ((TTXO)ClientData).OptionFlags |= (0x07 << 1);
                        break;
                    
                    default:
                        ((TTXO)ClientData).OptionFlags |= (0x01 << 1);
                        break;
                }
            }

        }

        internal TVFlxAlignment VAlign
        {
            get
            {
                switch ((((TTXO)ClientData).OptionFlags >> 4) & 0x07)
                {
                    case 2: return TVFlxAlignment.center;
                    case 3: return TVFlxAlignment.bottom;
                    case 4: return TVFlxAlignment.justify;
                    case 7: return TVFlxAlignment.distributed;
                    default:
                        return TVFlxAlignment.top;
                }
            }
            set
            {
                ((TTXO)ClientData).OptionFlags &= ~0x70; //Clear it

                switch (value)
                {
                    case TVFlxAlignment.center:
                        ((TTXO)ClientData).OptionFlags |= (0x02 << 4);
                        break;

                    case TVFlxAlignment.bottom:
                        ((TTXO)ClientData).OptionFlags |= (0x03 << 4);
                        break;

                    case TVFlxAlignment.justify:
                        ((TTXO)ClientData).OptionFlags |= (0x04 << 4);
                        break;

                    case TVFlxAlignment.distributed:
                        ((TTXO)ClientData).OptionFlags |= (0x07 << 4);
                        break;

                    default:
                        ((TTXO)ClientData).OptionFlags |= (0x01 << 4);
                        break;
                }
            }
        }








        internal TTextRotation TextRotation
        {
            get
            {
                switch (((TTXO)ClientData).Rotation)
                {
                    case 1: return TTextRotation.Vertical;
                    case 2: return TTextRotation.Rotated90Degrees;
                    case 3: return TTextRotation.RotatedMinus90Degrees;
                }

                return TTextRotation.Normal;
            }
            set
            {
                switch (value)
                {
                    case TTextRotation.Normal:
                        ((TTXO)ClientData).Rotation = 0;
                        break;

                    case TTextRotation.Rotated90Degrees:
                        ((TTXO)ClientData).Rotation = 2;
                        break;

                    case TTextRotation.RotatedMinus90Degrees:
                        ((TTXO)ClientData).Rotation = 3;
                        break;

                    case TTextRotation.Vertical:
                        ((TTXO)ClientData).Rotation = 1;
                        break;
                }
            }

        }

		internal override bool WaitingClientData(ref TClientType ClientType)
		{
			bool Result= base.WaitingClientData(ref ClientType);
			ClientType= TClientType.TTXO;
			return Result;
		}
	}


}
