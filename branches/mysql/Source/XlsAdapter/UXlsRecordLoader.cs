using System;
using FlexCel.Core;
using System.Diagnostics;

namespace FlexCel.XlsAdapter
{
    /// <summary>
    /// Loads a Record from a biff8 stream and creates the object in memory to hold it.
    /// </summary>
    internal sealed class TXlsRecordLoader : TBinRecordLoader
    {
        internal TOle2File DataStream;
        internal TBorderList BorderList;
        internal TPatternList PatternList;

        internal TXlsRecordLoader(TOle2File aDataStream, TBiff8XFMap aXFMap, TSST aSST, IFlexCelFontList aFontList, TBorderList aBorderList,
            TPatternList aPatternList, TEncryptionData aEncryption, TXlsBiffVersion aXlsBiffVersion, TNameRecordList aNames, TVirtualReader VirtualReader)
            :base(aSST, aFontList, aEncryption, aXlsBiffVersion, aXFMap, aNames, VirtualReader)
        {
            DataStream = aDataStream;
            BorderList = aBorderList;
            PatternList = aPatternList;
        }

        internal override void ReadHeader()
        {
            DataStream.Read(RecordHeader.Data, RecordHeader.Length);
        }

        internal override bool Eof
        {
            get
            {
                return DataStream.NextEof(3);
            }
        }

		/// <summary>
		/// We have this method as a loop, since a recursive call to LoadRecord can cause stack overflow for sst or other deeply nested continues.
		/// </summary>
		internal void LoadContinues(TBaseRecord Master)
		{
			TBaseRecord LastRecord = Master;
			
			do
			{
				int Id = RecordHeader.Id;
				//Debug.Assert(Id == (int) xlr.CONTINUE); It might be ContinueFrt, or others...

				byte[] Data= new byte[RecordHeader.Size];
				DataStream.Read(Data, Data.Length);

				if (Encryption.Engine!=null && Id!=(int)xlr.BOF && Id!=(int) xlr.INTERFACEHDR) 
					Data=Encryption.Engine.Decode(Data, DataStream.Position-Data.Length,0,Data.Length, Data.Length);  //Note that we do not care about BoundSheet, as it's data won't be used.

				TContinueRecord R = new TContinueRecord(Id, Data);
				LastRecord.AddContinue(R);
				LastRecord = R;

				if (!Eof)
				{
					DataStream.Read(RecordHeader.Data, RecordHeader.Length);
				}
				else
				{
					//Array.Clear(RecordHeader.Data,0,RecordHeader.Length);
					RecordHeader.Id=(int) xlr.EOF;  //Return EOFs, so in case of bad formed files we don't get on an infinite loop.
					RecordHeader.Size=0;
				}

			}  while (RecordHeader.Id == (int) xlr.CONTINUE ||
                RecordHeader.Id == (int) xlr.CONTINUEFRT ||
                RecordHeader.Id == (int) xlr.CONTINUEFRT11 ||
                RecordHeader.Id == (int) xlr.CONTINUEFRT12 ||
                RecordHeader.Id == (int) xlr.CONTINUEBIGNAME ||
                RecordHeader.Id == (int)xlr.CONTINUECRTMLFRT);

		}

        internal override TBaseRecord LoadUnsupportdedRecord() //This method is unrolled from the main one because it should never be called, so it doesn't make sense to slow down the main LoadRecord.
        {
            int Id = RecordHeader.Id;
            byte[] Data= new byte[RecordHeader.Size];
            DataStream.Read(Data, Data.Length);
            TBaseRecord R=null;

            if (Encryption.Engine!=null && Id!=(int)xlr.BOF && Id!=(int) xlr.INTERFACEHDR) 
                Data=Encryption.Engine.Decode(Data, DataStream.Position-Data.Length,0,Data.Length, Data.Length);  //Note that we do not care about BoundSheet, as it's data won't be used.

            switch (Id)
            {
                case (int)xlr.BOF: R = new TBOFRecord(Id, Data); break;
                case (int)xlr.EOF: R = new TEOFRecord(Id, Data); break;
                default: R = new TxBaseRecord(Id, Data); break;
			} //case

            //Peek at the next record...
            if (!Eof)
            {
                DataStream.Read(RecordHeader.Data, RecordHeader.Length);
            }
            else
            {
                //Array.Clear(RecordHeader.Data,0,RecordHeader.Length);
                RecordHeader.Id = (int)xlr.EOF;  //Return EOFs, so in case of bad formed files we don't get on an infinite loop.
                RecordHeader.Size = 0;
            }

            return R;
        }

        internal override TBaseRecord LoadRecord(out int rRow, bool InGlobalSection)
        {
            int Id = RecordHeader.Id;
            byte[] Data= new byte[RecordHeader.Size];
            DataStream.Read(Data, Data.Length);
            TBaseRecord R=null;

            if (Encryption.Engine!=null && Id!=(int)xlr.BOF && Id!=(int) xlr.INTERFACEHDR) 
                Data=Encryption.Engine.Decode(Data, DataStream.Position-Data.Length,0,Data.Length, Data.Length);  //Note that we do not care about BoundSheet, as it's data won't be used.

            if (Data.Length > 2) rRow = BitConverter.ToUInt16(Data, 0); else rRow = 0;

			switch (Id)  
			{
				case (int)xlr.BOF         : R = new TBOFRecord(Id, Data);break;
				case (int)xlr.EOF         : R = new TEOFRecord(Id, Data);break;
				
				case (int)xlr.TEMPLATE	  : R = new TTemplateRecord(Id, Data);break;
										  
				case (int)xlr.FORMULA	  : R = TFormulaRecord.CreateFromBiff8(Names, Id, Data, XFMap);break;
				case (int)xlr.FORMULABiff4  //you will never see this on a biff8 file. Just for lazy 3rd party solutions.
										  : R = TFormulaRecord.CreateFromBiff4(Names, Data, XFMap);break;
				
				case (int)xlr.SHRFMLA     : R = new TBiff8ShrFmlaRecord(Id, Data);break;

				case (int)xlr.OBJ         : R = new TObjRecord(Id, Data, Names);break;
				case (int)xlr.MSODRAWING  : R = new TDrawingRecord(Id, Data);break;
				case (int)xlr.MSODRAWINGGROUP
				: R = new TDrawingGroupRecord(Id, Data);break;

				case (int)xlr.HEADERIMG  : if (InGlobalSection)
											   R = new THeaderImageGroupRecord(Id, Data);
										   else
											   R = new THeaderImageRecord(Id, Data);

					break;

				case (int)xlr.TXO         : R = new TTXORecord(Id, Data);break;
				case (int)xlr.NOTE        : R = new TNoteRecord(Id, Data);break;
					//case (int)xlr.RECALCID:   //So the workbook gets recalculated. Not really useful because xl97 won't use it. Also the message "file has been saved with an older version" will be shown on xls2000 and up. The solution is to set the individual formulas to recalc on open.
				case (int)xlr.EXTSST:     // We will have to generate this again
				case (int)xlr.DBCELL:     //To find rows in blocks... we need to calculate it again
				case (int)xlr.INDEX:      //Same as DBCELL
				case (int)xlr.MSODRAWINGSELECTION:  // Object selection. We do not need to select any drawing
                    case(int)xlr.ENTEXU2
				: R = null;break;

				case (int)xlr.DIMENSIONS  //Used range of a sheet
				: R = new TDimensionsRecord(Id, Data);break;
				case (int)xlr.SST         : R = new TSSTRecord(Id, Data);break;
				case (int)xlr.BOUNDSHEET  : R = new TBoundSheetRecord(Id, Data);break;
				case (int)xlr.CODENAME    : R = new TCodeNameRecord(Id, Data);break;

				case (int)xlr.OBPROJ      : R = new TObProjRecord(Id, Data);break;

				case (int)xlr.SHEETEXT    : R = new TSheetExtRecord(Id, Data);break;

				case (int)xlr.ARRAY       :
				case (int)xlr.ARRAY2      : R = TArrayRecord.CreateFromBiff8(Names, Id, Data);break;
				case (int)xlr.BLANK       : R = new TBlankRecord(Id, Data, XFMap);break;
				case (int)xlr.BOOLERR     : R = new TBoolErrRecord(Id, Data, XFMap);break;
				case (int)xlr.NUMBER      : R = new TNumberRecord(Id, Data, XFMap);break;
				case (int)xlr.MULBLANK    : R = new TMulBlankRecord(Id, Data, XFMap);break;
				case (int)xlr.MULRK       : R = new TMulRKRecord(Id, Data, XFMap);break;
				case (int)xlr.RK          : R = new TRKRecord(Id, Data, XFMap);break;
				case (int)xlr.STRING      : R = new TStringRecord(Id, Data);break;  //String record saves the result of a formula

				case (int)xlr.INTERFACEHDR: R = new TInterfaceHdrRecord(Id, Data); break;
				case (int)xlr.INTERFACEEND: R = new TInterfaceEndRecord(Id, Data); break;

                case (int)xlr.XF          : R = new TXFRecord(Id, Data, BorderList, PatternList, XFMap); TXFCRCRecord.UpdateCRC(Data, ref XFCRC); XFCount++; break;
                case (int)xlr.XFEXT       : R = new TXFExtRecord(Data); break;
                case (int)xlr.XFCRC       : R = new TXFCRCRecord(Data); break;
                case (int)xlr.THEME       : R = new TThemeRecord(Data); break;
                case (int)xlr.DXF         : R = new TDXFRecord(Id, Data); break;
                case (int)xlr.TABLESTYLE  : R = new TTableStyleRecord(Id, Data); break;
                case (int)xlr.TABLESTYLES : R = new TTableStylesRecord(Id, Data); break;
                case (int)xlr.TABLESTYLEELEMENT  : R = new TTableStyleElementRecord(Id, Data); break;

				case (int)xlr.FONT        : R = new TFontRecord(Id, Data);break;
				case (int)xlr.xFORMAT     : R = new TFormatRecord(Id, Data);break;
				case (int)xlr.PALETTE     : R = new TPaletteRecord(Id, Data);break;

                case (int)xlr.CLRTCLIENT  : R = new TClrtClientRecord(Id, Data); break;
                case (int)xlr.FRTINFO     : R = new TFrtInfoRecord(Id, Data);break;

                case (int)xlr.STYLE       : R = new TStyleRecord(Id, Data, XFMap); break;
                case (int)xlr.STYLEEX     : R = new TStyleExRecord(Id, Data); break;

				case (int)xlr.LABELSST    : R = new TLabelSSTRecord(Id, Data, SST, FontList, XFMap);break;
				case (int)xlr.LABEL       : R = new TLabelRecord(Id, Data, XFMap);break;
				case (int)xlr.RSTRING     : R = new TRStringRecord(Id, Data, XFMap);break;
				case (int)xlr.ROW         : R = new TRowRecord(Id, Data, XFMap);break;
				case (int)xlr.NAME        : R = TNameRecord.CreateFromBiff8(Names, Id, Data);break;
                case (int)xlr.NAMECMT     : R = new TNameCmtRecord(Id, Data); break;
				case (int)xlr.TABLE       : R = TTableRecord.CreateFromBiff8(Id, Data);break;

				case (int)xlr.CELLMERGING : R = new TCellMergingRecord(Id, Data);break;
				case (int)xlr.CONDFMT     : R = new TCondFmtRecord(Id, Data);break;
				case (int)xlr.CF          : R = TCFRecord.LoadFromBiff8(Names, Id, Data);break;
				case (int)xlr.DVAL        : R = new TDValRecord(Id, Data);break;
				case (int)xlr.DV          : R = new TDVRecord(Id, Data);break;
				case (int)xlr.HLINK       : R = THLinkRecord.CreateFromBiff8(Id, Data);break;
				case (int)xlr.SCREENTIP   : R = TScreenTipRecord.CreateFromBiff8(Id, Data);break;
				case (int)xlr.CONTINUE    : R = new TContinueRecord(Id, Data);break;

				case (int)xlr.FOOTER      : R = new TPageFooterRecord(Id, Data);break;
				case (int)xlr.HEADER      : R = new TPageHeaderRecord(Id, Data);break;
				case (int)xlr.HEADERFOOTER: R = new THeaderFooterExtRecord(Id, Data);break;

				case (int)xlr.PRINTGRIDLINES : R = new TPrintGridLinesRecord(Id, Data);break;

                case (int)xlr.LEFTMARGIN   : R = new TLeftMarginRecord(Id, Data); break;
                case (int)xlr.RIGHTMARGIN  : R = new TRightMarginRecord(Id, Data); break;
                case (int)xlr.TOPMARGIN    : R = new TTopMarginRecord(Id, Data); break;
				case (int)xlr.BOTTOMMARGIN : R = new TBottomMarginRecord(Id, Data);break;

				case (int)xlr.SETUP        : R = new TSetupRecord(Id, Data);break;
				case (int)xlr.PLS          : R = new TPlsRecord(Id, Data);break;
				case (int)xlr.PRINTHEADERS : R = new TPrintHeadersRecord(Id, Data);break;
				case (int)xlr.VCENTER      : R = new TVCenterRecord(Id, Data);break; 
				case (int)xlr.HCENTER      : R = new THCenterRecord(Id, Data);break;
				case (int)xlr.WSBOOL       : R = new TWsBoolRecord(Id, Data);break;

                case (int)xlr.AutoFilterINFO: R = new TAutoFilterInfoRecord(Id, Data); break;
                case (int)xlr.AutoFilter    : R = new TAutoFilterRecord(Id, Data); break;
                case (int)xlr.AutoFilter12  : R = new TAutoFilter12Record(Id, Data); break;

				case (int)xlr.XCT:        // Cached values of a external workbook... 
				case (int)xlr.CRN         // Cached values also
				                          : R = null;break;

                case (int)xlr.DSF         : R = new TDSFRecord(); break;

				case (int)xlr.SUPBOOK     : R = new TSupBookRecord(Id, Data);break;
				case (int)xlr.EXTERNSHEET : R = new TExternSheetRecord(Id, Data);break;
				case (int)xlr.EXTERNNAME  :
				case (int)xlr.EXTERNNAME2
				                          : R = new TExternNameRecord(Id, Data);break;

				case (int)xlr.COUNTRY     : R = new TCountryRecord(Id, Data);break;
				case (int)xlr.CODEPAGE    : R = new TCodePageRecord(Id, Data);break;
                case (int)xlr.XL9FILE     : R = new TExcel9FileRecord(Id, Data); break;
                case (int)xlr.OBNOMACROS  : R = new TObNoMacrosRecord(Id, Data); break;

                case (int)xlr.OLESIZE     : R = new TOleObjectSizeRecord(Id, Data); break;

				case (int)xlr.WINDOW1     : R = new TWindow1Record(Id, Data);break;
				case (int)xlr.WINDOW2     : R = new TWindow2Record(Id, Data);break;

				case (int)xlr.PANE        : R = new TPaneRecord(Id, Data);break;
				case (int)xlr.SELECTION   : R = new TBiff8SelectionRecord(Id, Data);break;

				case (int)xlr.SCL         : R = new TSCLRecord(Id, Data);break;

				case (int)xlr.GUTS        : R = new TGutsRecord(Id, Data); break;

				case (int)xlr.SXVIEW      : R = new TSxViewRecord(Id, Data); break;     

                case (int)xlr.SXIDSTM:
                case (int)xlr.SXVS:
                //case (int)  xlr.DCONREF: //used by SXVS   This have their own now.
                //case (int)  xlr.DCONNAME: //used by SXVS
                //case (int)  xlr.DCONBIN: //used by SXVS
                case (int)  xlr.SXTBL: //used by SXVS
                case (int)  xlr.SXEXTPARAMQRY: //used by SXVS
                
                case (int)    xlr.SXTBPG: //used by SXTBL
                case (int)    xlr.SXTBRGITEM: //used by SXTBL
                case (int)    xlr.SXSTRING: //used by SXTBL
                
                case (int) xlr.SXADDL:
                case (int) xlr.SXADDL12:
                                          R = new TPivotCacheRecord(Id, Data);break;

                    //PivotCore
                case (int)xlr.SXVD:
                case (int)xlr.SXVI:
                case (int)xlr.SXVDEX:
                case (int)xlr.SXIVD:
                case (int)xlr.SXPI:
                case (int)xlr.SXDI:
                case (int)xlr.SXLI:
                case (int)xlr.SXEX:
                case (int)xlr.SXSELECT:
                case (int)xlr.SXFORMAT:
                case (int)xlr.SXDXF:
                case (int)xlr.SXRULE:
                case (int)xlr.SXFILT:
                case (int)xlr.SXITM:

                    //PivotFRT
                case (int)xlr.QSISXTAG:

                case (int)xlr.DBQUERYEXT:
                case (int)xlr.EXTSTRING:
                case (int)xlr.OLEDBCONN:
                case (int)xlr.TXTQRY:

                case (int)xlr.SXVIEWEX:
                case (int)xlr.SXTH:
                case (int)xlr.SXPIEX:
                case (int)xlr.SXVDTEX:
                case (int)xlr.SXVIEWEX9:
                    R = new TPivotSheetRecord(Id, Data); break;

                case (int)xlr.SXVIEWLINK: R = new TSxViewLinkRecord(Id, Data); break;
                case (int)xlr.PIVOTCHARTBITS: R = new TPivotChartBitsRecord(Id, Data); break;
                case (int)xlr.ChartSbaseref: R = new TChartSBaseRefRecord(Id, Data); break;

                case (int)xlr.INTL: R = new TInternationalRecord(Id, Data); break;

				case (int)xlr.HORIZONTALPAGEBREAKS: R = new TBiff8HPageBreakRecord(Id, Data);break;
				case (int)xlr.VERTICALPAGEBREAKS  : R = new TBiff8VPageBreakRecord(Id, Data);break;

				case (int)xlr.COLINFO     : R = new TColInfoRecord(Id, Data, XFMap);break;
				case (int)xlr.DEFCOLWIDTH : R = new TDefColWidthRecord(Id, Data);break;
				case (int)xlr.STANDARDWIDTH : R = new TStandardWidthRecord(Id, Data);break;
				case (int)xlr.DEFAULTROWHEIGHT: R = new TDefaultRowHeightRecord(Id, Data);break;

                case (int)xlr.DOCROUTE: R = new TDocRouteRecord(Id, Data); break;
                case (int)xlr.RECIPNAME: R = new TRecipNameRecord(Id, Data); break;
                case (int)xlr.USERBVIEW: R = new TUserBViewRecord(Id, Data); break;

                case (int)xlr.USERSVIEWBEGIN: R = new TUserSViewBeginRecord(Id, Data); break;
                case (int)xlr.USERSVIEWEND: R = new TUserSViewEndRecord(Id, Data); break;

                case (int)xlr.UNITS: R = new TUnitsRecord(Id, Data); break;
                case (int)xlr.CRTMLFRT: R = new TCrtMlFrtRecord(Id, Data); break;
					
				case (int)xlr.FILEPASS    : TFilePassRecord Fr = new TFilePassRecord(Id, Data, false); 
					Encryption.Engine=Fr.CreateEncryptionEngine(Encryption.ReadPassword); R=null;

                    if (!Encryption.Engine.CheckHash(Encryption.ReadPassword))
                    {
                        if (Encryption.OnPassword != null)
                        {
                            OnPasswordEventArgs ea = new OnPasswordEventArgs(Encryption.Xls);
                            Encryption.OnPassword(ea);
                            Encryption.ReadPassword = ea.Password;
							Encryption.Engine=Fr.CreateEncryptionEngine(Encryption.ReadPassword);
						}
                    }

                    if (!Encryption.Engine.CheckHash(Encryption.ReadPassword))
                        XlsMessages.ThrowException(XlsErr.ErrInvalidPassword);
                    
                    break;

				case (int)xlr.PROTECT       : R = new TProtectRecord(Id, Data);break;
				case (int)xlr.WINDOWPROTECT : R = new TWindowProtectRecord(Id, Data);break;
				case (int)xlr.OBJPROTECT    : R = new TObjProtectRecord(Id, Data);break;
				case (int)xlr.SCENPROTECT   : R = new TScenProtectRecord(Id, Data);break;
				case (int)xlr.PASSWORD      : R = new TPasswordRecord(Id, Data);break;
				case (int)xlr.FEATHDR       : R = TFeatHdrRecord.Create(Id, Data); break;

				case (int)xlr.WRITEPROT     : R = new TWriteProtRecord(Id, Data);break;
				case (int)xlr.WRITEACCESS   : R = new TWriteAccessRecord(Id, Data);break;
				case (int)xlr.FILESHARING   : R = new TFileSharingRecord(Id, Data);break;

                case (int)xlr.PROT4REV      : R = new TProt4RevRecord(Id, Data); break;
                case (int)xlr.PROT4REVPASS  : R = new TProt4RevPassRecord(Id, Data); break;
 
                case (int)xlr.HIDEOBJ       : R = new THideObjRecord(Id, Data); break;
                
                case (int)xlr.x1904         : R = new T1904Record(Id, Data); break;
                case (int)xlr.BACKUP        : R = new TBackupRecord(Id, Data); break;
                case (int)xlr.REFRESHALL    : R = new TRefreshAllRecord(Id, Data); break;
				case (int)xlr.PRECISION     : R = new TPrecisionRecord(Id, Data);break;
				case (int)xlr.BOOKBOOL      : R = new TBookBoolRecord(Id, Data);break;
                case (int)xlr.MTRSETTINGS   : R = new TMTRSettingsRecord(Id, Data); break;
                case (int)xlr.FORCEFULLCALCULATION: R = new TForceFullCalculationRecord(Id, Data); break;

                case (int)xlr.CALCMODE      : R = new TCalcModeRecord(Id, Data); break;
                case (int)xlr.CALCCOUNT     : R = new TCalcCountRecord(Id, Data); break;


                case (int)xlr.USESELFS      : R = new TUsesELFsRecord(Id, Data); break;
                case (int)xlr.RECALCID      : R = new TRecalcIdRecord(Id, Data); break;
                case (int)xlr.WEBPUB        : R = new TWebPubRecord(Id, Data); break;
                case (int)xlr.WOPT          : R = new TWOptRecord(Id, Data); break;
                case (int)xlr.BOOKEXT       : R = new TBookExtRecord(Id, Data); break;

                case (int)xlr.CRERR         : R = new TCRErrRecord(Id, Data); break; //workbook is marked for recovery. We will ignore this one, but we need to load it here (as a TxBaseRecord) because it can have continues

                case (int)xlr.LEL           : R = new TLelRecord(Id, Data); break;

                case (int)xlr.FNGROUPCOUNT:
                case (int)xlr.FNGROUPNAME:
                case (int)xlr.FNGRP12:
                                              R = new TFnGroupRecord(Id, Data);break;

                case (int)xlr.MDB:
                case (int)xlr.MDTInfo:
                case (int)xlr.MDXKPI:
                case (int)xlr.MDXProp:
                case (int)xlr.MDXSet:
                case (int)xlr.MDXStr:
                case (int)xlr.MDXTuple:
                                              R = new TMetaDataRecord(Id, Data); break;

                case (int)xlr.RTD           : R = new TRTDRecord(Id, Data); break;
                case (int)xlr.DCONN         : R = new TDConnRecord(Id, Data); break;

                case (int)xlr.TABID         : R = new TTabIdRecord(Id, Data); break;

                case (int)xlr.COMPRESSPICTURES : R = new TCompressPicturesRecord(Id, Data); break;
                case (int)xlr.COMPAT12      : R = new TCompat12Record(Id, Data); break;
                case (int)xlr.GUIDTYPELIB   : R = new TGUIDTypeLibRecord(Id, Data); break;

				case (int)xlr.REFMODE       : R = new TRefModeRecord(Id, Data);break;
                case (int)xlr.ITERATION     : R = new TIterationRecord(Id, Data); break;
                case (int)xlr.DELTA         : R = new TDeltaRecord(Id, Data); break;
                    
                case (int)xlr.SAVERECALC    : R = new TSaveRecalcRecord(Id, Data); break;
                case (int)xlr.GRIDSET       : R = new TGridSetRecord(Id, Data); break;

                case (int)xlr.SYNC          : R = new TSyncRecord(Id, Data); break;
                case (int)xlr.LPR           : R = new TLprRecord(Id, Data); break;
                case (int)xlr.PLV           : R = new TPlvRecord(Id, Data); break;
                
                case (int)xlr.BITMAP        : R = new TBgPicRecord(Id, Data); break;
                
                case (int)xlr.BIGNAME:
                case (int)xlr.CONTINUEBIGNAME //we don't really care about this one, just add it to the list.
                : R = new TBigNameRecord(Id, Data); break;

                case (int)xlr.SCENMAN       : R = new TScenManRecord(Id, Data); break;
                case (int)xlr.SCENARIO      : R = new TScenarioRecord(Id, Data); break;

                case (int)xlr.SORT          : R = new TSortRecord(Id, Data);break;
                case (int)xlr.SORTDATA      : R = new TSortDataRecord(Id, Data);break;

                case (int)xlr.DROPDOWNOBJIDS: R = new TDropDownObjIdsRecord(Id, Data); break;

                case (int)xlr.RRSORT        : R = new TRRSortRecord(Id, Data); break;
                case (int)xlr.LRNG          : R = new TLRngRecord(Id, Data);break;

                case (int)xlr.PHONETIC      : R = new TPhoneticRecord(Id, Data); break;
                
                case (int)xlr.FILTERMODE    : R = new TFilterModeRecord(Id, Data); break;

                case (int)xlr.PRINTSIZE     : R = new TPrintSizeRecord(Id, Data);break;
                case (int)xlr.DCON          : R = new TDConRecord(Id, Data);break;
                case (int)xlr.DCONNAME      : R = new TDConNameRecord(Id, Data);break;
                case (int)xlr.DCONBIN       : R = new TDConBinRecord(Id, Data);break;
                case (int)xlr.DCONREF       : R = new TDConRefRecord(Id, Data);break;

                case (int)xlr.QSI           : R = new TQSIRecord(Id, Data);break;

                case (int)xlr.FEAT          : R = new TFeatRecord(Id, Data);break;
                case (int)xlr.FEAT11:        
                case (int)xlr.FEAT12        
                : R = new TFeat1112Record(Id, Data);break;

                case (int)xlr.FEATHDR11     : R = new TFeatHdr11Record(Id, Data);break;
                case (int)xlr.LIST12        : R = new TList12Record(Id, Data);break;


				#region Charts
                case (int)xlr.ChartSiindex     : R = new TChartSIIndexRecord(Id, Data); break;
                case (int)xlr.ChartFbi2:
				case (int)xlr.ChartFbi         : R = null; break; //R = new TChartFBIRecord(Id, Data); break; //ChartFBI records are very dangerous. Each one of them must point to a different font, and other data cannot point to the same font record either. The problem is that we do not know all the reocords that might exist and carry a font, so we  cannot reliably enable this. As it is now, if you uncomment this, fbi records will work for known records, but some (like the "thousands" label) will not and crash Excel.
                                                                  //FBI records are ignored in Excle 2007 anyway. 
				case (int)xlr.BEGIN            : R = new TBeginRecord(Id, Data);break;
				case (int)xlr.END              : R = new TEndRecord(Id, Data);break;

				case (int)xlr.ChartAI          : R = TChartAIRecord.CreateFromBiff8(Names, Id, Data);break;
				case (int)xlr.ChartChart       : R = new TChartChartRecord(Id, Data);break;
				case (int)xlr.ChartFrame       : R = new TChartFrameRecord(Id, Data);break;
				case (int)xlr.ChartPlotgrowth  : R = new TChartPlotGrowthRecord(Id, Data);break;
				case (int)xlr.ChartSeries      : R = new TChartSeriesRecord(Id, Data);break;
				case (int)xlr.ChartDefaulttext : R = new TChartDefaultTextRecord(Id, Data);break;
				case (int)xlr.ChartText        : R = new TChartTextRecord(Id, Data);break;
				case (int)xlr.ChartSeriestext  : R = new TChartSeriesTextRecord(Id, Data);break;
				case (int)xlr.ChartPos         : R = new TChartPosRecord(Id, Data);break;
				case (int)xlr.ChartAxis        : R = new TChartAxisRecord(Id, Data);break;
				case (int)xlr.ChartAxcext      : R = new TChartAxcExtRecord(Id, Data);break;
				case (int)xlr.ChartValuerange  : R = new TChartValueRangeRecord(Id, Data);break;
				case (int)xlr.ChartAxisparent  : R = new TChartAxisParentRecord(Id, Data);break;
				case (int)xlr.ChartChartformat : R = new TChartChartFormatRecord(Id, Data);break;
				case (int)xlr.ChartLegend      : R = new TChartLegendRecord(Id, Data);break;
				case (int)xlr.ChartLegendxn    : R = new TChartLegendXnRecord(Id, Data);break;
				case (int)xlr.ChartDataformat  : R = new TChartDataFormatRecord(Id, Data);break;
				case (int)xlr.ChartShtprops    : R = new TChartShtPropsRecord(Id, Data); break;

				case (int)xlr.ChartObjectLink  : R = new TChartObjectLinkRecord(Id, Data); break;
				case (int)xlr.ChartAlruns      : R = new TChartALRunsRecord(Id, Data); break;

                case (int)xlr.ChartDataLabExtContent: R = new TChartDataLabExtContentsRecord(Id, Data); break;

				case (int)xlr.ChartPlotarea    : R = new TChartPlotAreaRecord(Id, Data);break;

				case (int)xlr.ChartAreaformat  : R = new TChartAreaFormatRecord(Id, Data);break;
				case (int)xlr.ChartLineformat  : R = new TChartLineFormatRecord(Id, Data);break;
				case (int)xlr.ChartPieformat   : R = new TChartPieFormatRecord(Id, Data);break;
				case (int)xlr.ChartMarkerformat: R = new TChartMarkerFormatRecord(Id, Data);break;
				case (int)xlr.ChartSerfmt      : R = new TChartSerFmtRecord(Id, Data);break;
				case (int)xlr.ChartGelframe    : R = new TChartGelFrameRecord(Id, Data);break;

				case (int)xlr.ChartFontx       : R = new TChartFontXRecord(Id, Data);break;
				case (int)xlr.ChartIfmt        : R = new TChartIFmtRecord(Id, Data);break;
				case (int)xlr.ChartTick        : R = new TChartTickRecord(Id, Data);break;
				case (int)xlr.ChartCatserrange : R = new TChartCatSerRangeRecord(Id, Data);break;

				case (int)xlr.ChartArea        : R = new TChartAreaRecord(Id, Data);break;
				case (int)xlr.ChartBar         : R = new TChartBarRecord(Id, Data);break;
				case (int)xlr.ChartLine        : R = new TChartLineRecord(Id, Data);break;
				case (int)xlr.ChartPie         : R = new TChartPieRecord(Id, Data);break;
				case (int)xlr.ChartRadar       : R = new TChartRadarRecord(Id, Data);break;
				case (int)xlr.ChartScatter     : R = new TChartScatterRecord(Id, Data);break;
				case (int)xlr.ChartSurface     : R = new TChartSurfaceRecord(Id, Data);break;

				case (int)xlr.ChartDropbar     : R = new TChartDropBarRecord(Id, Data);break;
				case (int)xlr.ChartChartline   : R = new TChartChartLineRecord(Id, Data);break;

				case (int)xlr.ChartAttachedlabel: R = new TChartAttachedLabelRecord(Id, Data);break;


				case (int)xlr.ChartAxislineformat  : R = new TChartAxisLineFormatRecord(Id, Data);break;
					#endregion
 
				default                     : if (Id >0x1000) 
												  R= new TxChartBaseRecord(Id, Data);
											  else
												  R= new TxBaseRecord(Id, Data);
					break;
			} //case

            //Peek at the next record...
            if (!Eof)
            {
                DataStream.Read(RecordHeader.Data, RecordHeader.Length);
                int Id2 = RecordHeader.Id;

                switch (Id2)
                {
                    case (int)xlr.CONTINUE:
                    case (int)xlr.CONTINUEFRT:
                    case (int)xlr.CONTINUEFRT11:
                    case (int)xlr.CONTINUEFRT12:
                    case (int)xlr.CONTINUEBIGNAME:
                    case (int)xlr.CONTINUECRTMLFRT:
						LoadContinues(R);
                        break;

                    case (int)xlr.TABLE:
                        TFormulaRecord Rf = R as TFormulaRecord;
                        if (Rf != null)                                                
                            Rf.TableRecord=(TTableRecord)LoadRecord(InGlobalSection);
                        else XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                        break;

					case (int)xlr.ARRAY:
					case (int)xlr.ARRAY2:
                        TFormulaRecord Rfa = R as TFormulaRecord;
                        if (Rfa != null)                                                
                            Rfa.ArrayRecord=(TArrayRecord)LoadRecord(InGlobalSection);
                        else XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                        break;
				
                    case (int)xlr.SCREENTIP:
                        THLinkRecord Rs = R as THLinkRecord;
                        if (Rs != null)                                                
                            Rs.Hint=(TScreenTipRecord)LoadRecord(InGlobalSection);
                        else XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                        break;
					
                    case (int)xlr.STRING:
                        if (!(R is TFormulaRecord) &  !(R is TBiff8ShrFmlaRecord) & !(R is TArrayRecord) & !(R is TTableRecord)) XlsMessages.ThrowException(XlsErr.ErrExcelInvalid);
                        break;
                }
            }
            else
            {
                //Array.Clear(RecordHeader.Data,0,RecordHeader.Length);
                RecordHeader.Id=(int) xlr.EOF;  //Return EOFs, so in case of bad formed files we don't get on an infinite loop.
                RecordHeader.Size=0;
            }
            return R;
        }
    }
}
