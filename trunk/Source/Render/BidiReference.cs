using System;
using FlexCel.Core;
//THIS FILE IS CONVERTED FROM THE REFERENCE IMPLEMENTATION FROM THE UNICODE
//BIDIRECTIONAL ALGORITHM IN JAVA. 
//http://www.unicode.org/reports/tr9/BidiReferenceJava/BidiReference.java.txt


/*
* (C) Copyright IBM Corp. 1999, All Rights Reserved
*
* version 1.1
*/

namespace FlexCel.Render
{

	/// <summary> Reference implementation of the Unicode 3.0 Bidi algorithm. 
	/// <p>
	/// This implementation is not optimized for performance.  It is intended
	/// as a reference implementation that closely follows the specification
	/// of the Bidirectional Algorithm in The Unicode Standard version 3.0.
	/// </p>
	/// <p>
	/// <b>Input:</b>
	/// There are two levels of input to the algorithm, since clients may prefer
	/// to supply some information from out-of-band sources rather than relying on
	/// the default behavior.
	/// </p>
	/// <ol>
	/// <li>unicode type array</li>
	/// <li>unicode type array, with externally supplied base line direction</li>
	/// </ol>
	/// <p><b>Output:</b>
	/// Output is separated into several stages as well, to better enable clients
	/// to evaluate various aspects of implementation conformance.
	/// </p>
	/// <ol>
	/// <li>levels array over entire paragraph</li>
	/// <li>reordering array over entire paragraph</li>
	/// <li>levels array over line</li>
	/// <li>reordering array over line</li>
	/// </ol>
	/// Note that for conformance, algorithms are only required to generate correct
	/// reordering and character directionality (odd or even levels) over a line.
	/// Generating identical level arrays over a line is not required.  Bidi
	/// explicit format codes (LRE, RLE, LRO, RLO, PDF) and BN can be assigned
	/// arbitrary levels and positions as long as the other text matches.
	/// <p>
	/// As the algorithm is defined to operate on a single paragraph at a time,
	/// this implementation is written to handle single paragraphs.  Thus
	/// rule P1 is presumed by this implementation-- the data provided to the
	/// implementation is assumed to be a single paragraph, and either contains no
	/// 'B' codes, or a single 'B' code at the end of the input.  'B' is allowed
	/// as input to illustrate how the algorithm assigns it a level.
	/// </p>
	/// <p>
	/// Also note that rules L3 and L4 depend on the rendering engine that uses 
	/// the result of the bidi algorithm.  This implementation assumes that the 
	/// rendering engine expects combining marks in visual order (e.g. to the 
	/// left of their base character in RTL runs) and that it adjust the glyphs 
	/// used to render mirrored characters that are in RTL runs so that they 
	/// render appropriately.
	/// </p>
	/// </summary>
	/// <author>  Doug Felt
	/// </author>
	internal sealed class BidiReference
	{
		/// <summary> Return the base level of the paragraph.</summary>
		public sbyte BaseLevel
		{
			get
			{
				return paragraphEmbeddingLevel;
			}
		
		}
		private static sbyte[] allTypes;
		private sbyte[] initialTypes;
		private sbyte[] embeddings; // generated from processing format codes
		private sbyte paragraphEmbeddingLevel = - 1; // undefined
	
		private int textLength; // for convenience
		private sbyte[] resultTypes; // for paragraph, not lines
		private sbyte[] resultLevels; // for paragraph, not lines
	
		// The bidi types
	
		/// <summary>Left-to-right</summary>
		public const sbyte L = 0;
	
		/// <summary>Left-to-Right Embedding </summary>
		public const sbyte LRE = 1;
	
		/// <summary>Left-to-Right Override </summary>
		public const sbyte LRO = 2;
	
		/// <summary>Right-to-Left </summary>
		public const sbyte R = 3;
	
		/// <summary>Right-to-Left Arabic </summary>
		public const sbyte AL = 4;
	
		/// <summary>Right-to-Left Embedding </summary>
		public const sbyte RLE = 5;
	
		/// <summary>Right-to-Left Override </summary>
		public const sbyte RLO = 6;
	
		/// <summary>Pop Directional Format </summary>
		public const sbyte PDF = 7;
	
		/// <summary>European Number </summary>
		public const sbyte EN = 8;
	
		/// <summary>European Number Separator </summary>
		public const sbyte ES = 9;
	
		/// <summary>European Number Terminator </summary>
		public const sbyte ET = 10;
	
		/// <summary>Arabic Number </summary>
		public const sbyte AN = 11;
	
		/// <summary>Common Number Separator </summary>
		public const sbyte CS = 12;
	
		/// <summary>Non-Spacing Mark </summary>
		public const sbyte NSM = 13;
	
		/// <summary>Boundary Neutral </summary>
		public const sbyte BN = 14;
	
		/// <summary>Paragraph Separator </summary>
		public const sbyte B = 15;
	
		/// <summary>Segment Separator </summary>
		public const sbyte S = 16;
	
		/// <summary>Whitespace </summary>
		public const sbyte WS = 17;
	
		/// <summary>Other Neutrals </summary>
		public const sbyte ON = 18;
	
		/// <summary>Minimum bidi type value. </summary>
		public const sbyte TYPE_MIN = 0;
	
		/// <summary>Maximum bidi type value. </summary>
		public const sbyte TYPE_MAX = 18;
	
		/// <summary>Shorthand names of bidi type values, for error reporting. </summary>
		public static readonly System.String[] typenames = new System.String[]{"L", "LRE", "LRO", "R", "AL", "RLE", "RLO", "PDF", "EN", "ES", "ET", "AN", "CS", "NSM", "BN", "B", "S", "WS", "ON"};
	
		//
		// Input
		//
	
		/// <summary> Initialize using an array of direction types.  Types range from TYPE_MIN to TYPE_MAX inclusive 
		/// and represent the direction codes of the characters in the text.
		/// 
		/// </summary>
		/// <param name="Text">the Text we want to order.
		/// </param>
		public BidiReference(string Text)
		{
			if (allTypes==null) allTypes = CreateAllTypes();
			initialTypes = new sbyte[Text.Length];
			for (int i=0; i< Text.Length; i++)
				initialTypes[i] = allTypes[(int)Text[i]];

			runAlgorithm();
		}
	
		public static sbyte Direction(char c) 
	    {
			if (allTypes==null) allTypes = CreateAllTypes();
			return allTypes[(int)c];
		}

		/// <summary> The algorithm.
		/// Does not include line-based processing (Rules L1, L2).
		/// These are applied later in the line-based phase of the algorithm.
		/// </summary>
		private void  runAlgorithm()
		{
			// Ensure trace hook does not change while running algorithm.
			// Trace hook is a shared class resource.
			//lock (typeof(BidiReference))
		{
			textLength = initialTypes.Length;
			
			// Initialize output types.
			// Result types initialized to input types.
			resultTypes = new sbyte[initialTypes.Length];
			initialTypes.CopyTo(resultTypes, 0);
			
			// 1) determining the paragraph level
			// Rule P1 is the requirement for entering this algorithm.
			// Rules P2, P3. 
			// If no externally supplied paragraph embedding level, use default.
			if (paragraphEmbeddingLevel == - 1)
			{
				determineParagraphEmbeddingLevel();
			}
			
			// Initialize result levels to paragraph embedding level.
			resultLevels = new sbyte[textLength];
			setLevels(0, textLength, paragraphEmbeddingLevel);
			
			// 2) Explicit levels and directions
			// Rules X1-X8.
			determineExplicitEmbeddingLevels();
			
			// Rule X9.
			textLength = removeExplicitCodes();
			
			// Rule X10.
			// Run remainder of algorithm one level run at a time
			sbyte prevLevel = paragraphEmbeddingLevel;
			int start = 0;
			while (start < textLength)
			{
				sbyte level = resultLevels[start];
				sbyte prevType = typeForLevel(System.Math.Max((byte) prevLevel, (byte) level));
				
				int limit = start + 1;
				while (limit < textLength && resultLevels[limit] == level)
				{
					++limit;
				}
				
				sbyte succLevel = limit < textLength?resultLevels[limit]:paragraphEmbeddingLevel;
				sbyte succType = typeForLevel(System.Math.Max((byte) succLevel, (byte) level));
				
				// 3) resolving weak types
				// Rules W1-W7.
				resolveWeakTypes(start, limit, level, prevType, succType);
				
				// 4) resolving neutral types
				// Rules N1-N3.
				resolveNeutralTypes(start, limit, level, prevType, succType);
				
				// 5) resolving implicit embedding levels
				// Rules I1, I2.
				resolveImplicitLevels(start, limit, level, prevType, succType);
				
				prevLevel = level;
				start = limit;
			}
		}
		
			// Reinsert explicit codes and assign appropriate levels to 'hide' them.
			// This is for convenience, so the resulting level array maps 1-1 
			// with the initial array.
			// See the implementation suggestions section of TR#9 for guidelines on 
			// how to implement the algorithm without removing and reinserting the codes.
			textLength = reinsertExplicitCodes(textLength);
		}
	
		/// <summary> 1) determining the paragraph level.
		/// <p>
		/// Rules P2, P3.
		/// </p>
		/// At the end of this function, the member variable paragraphEmbeddingLevel is set to either 0 or 1.
		/// </summary>
		private void  determineParagraphEmbeddingLevel()
		{
			sbyte strongType = - 1; // unknown
		
			// Rule P2.
			for (int i = 0; i < textLength; ++i)
			{
				sbyte t = resultTypes[i];
				if (t == L || t == AL || t == R)
				{
					strongType = t;
					break;
				}
			}
		
			// Rule P3.
			if (strongType == - 1)
			{
				// none found
				// default embedding level when no strong types found is 0.
				paragraphEmbeddingLevel = 0;
			}
			else if (strongType == L)
			{
				paragraphEmbeddingLevel = 0;
			}
			else
			{
				// AL, R
				paragraphEmbeddingLevel = 1;
			}
		}
	
		/// <summary> Process embedding format codes.
		/// <p>
		/// Calls processEmbeddings to generate an embedding array from the explicit format codes.  The
		/// embedding overrides in the array are then applied to the result types, and the result levels are
		/// initialized.</p>
		/// </summary>
		private void  determineExplicitEmbeddingLevels()
		{
			embeddings = processEmbeddings(resultTypes, paragraphEmbeddingLevel);
		
			for (int i = 0; i < textLength; ++i)
			{
				sbyte level = embeddings[i];
				if ((level & 0x80) != 0)
				{
					level &= 0x7f;
					resultTypes[i] = typeForLevel(level);
				}
				resultLevels[i] = level;
			}
		}
	
		/// <summary> Rules X9.
		/// Remove explicit codes so that they may be ignored during the remainder 
		/// of the main portion of the algorithm.  The length of the resulting text 
		/// is returned.
		/// </summary>
		/// <returns> the length of the data excluding explicit codes and BN.
		/// </returns>
		private int removeExplicitCodes()
		{
			int w = 0;
			for (int i = 0; i < textLength; ++i)
			{
				sbyte t = initialTypes[i];
				if (!(t == LRE || t == RLE || t == LRO || t == RLO || t == PDF || t == BN))
				{
					embeddings[w] = embeddings[i];
					resultTypes[w] = resultTypes[i];
					resultLevels[w] = resultLevels[i];
					w++;
				}
			}
			return w; // new textLength while explicit levels are removed
		}
	
		/// <summary> Reinsert levels information for explicit codes.
		/// This is for ease of relating the level information 
		/// to the original input data.  Note that the levels
		/// assigned to these codes are arbitrary, they're
		/// chosen so as to avoid breaking level runs.
		/// </summary>
		/// <param name="aTextLength">the length of the data after compression
		/// </param>
		/// <returns> the length of the data (original length of 
		/// types array supplied to constructor)
		/// </returns>
		private int reinsertExplicitCodes(int aTextLength)
		{
			//int r = textLength;
			for (int i = initialTypes.Length; --i >= 0; )
			{
				sbyte t = initialTypes[i];
				if (t == LRE || t == RLE || t == LRO || t == RLO || t == PDF || t == BN)
				{
					embeddings[i] = 0;
					resultTypes[i] = t;
					resultLevels[i] = - 1;
				}
				else
				{
					--aTextLength;
					embeddings[i] = embeddings[aTextLength];
					resultTypes[i] = resultTypes[aTextLength];
					resultLevels[i] = resultLevels[aTextLength];
				}
			}
		
			// now propagate forward the levels information (could have 
			// propagated backward, the main thing is not to introduce a level
			// break where one doesn't already exist).
		
			if (resultLevels[0] == - 1)
			{
				resultLevels[0] = paragraphEmbeddingLevel;
			}
			for (int i = 1; i < initialTypes.Length; ++i)
			{
				if (resultLevels[i] == - 1)
				{
					resultLevels[i] = resultLevels[i - 1];
				}
			}
		
			// Embedding information is for informational purposes only
			// so need not be adjusted.
		
			return initialTypes.Length;
		}
	
		/// <summary> 2) determining explicit levels
		/// Rules X1 - X8
		/// 
		/// The interaction of these rules makes handling them a bit complex.
		/// This examines resultTypes but does not modify it.  It returns embedding and
		/// override information in the result array.  The low 7 bits are the level, the high
		/// bit is set if the level is an override, and clear if it is an embedding.
		/// </summary>
		private static sbyte[] processEmbeddings(sbyte[] resultTypes, sbyte paragraphEmbeddingLevel)
		{
			int EXPLICIT_LEVEL_LIMIT = 62;
		
			int textLength = resultTypes.Length;
			sbyte[] embeddings = new sbyte[textLength];
		
			// This stack will store the embedding levels and override status in a single byte
			// as described above.
			sbyte[] embeddingValueStack = new sbyte[EXPLICIT_LEVEL_LIMIT];
			int stackCounter = 0;
		
			// An LRE or LRO at level 60 is invalid, since the new level 62 is invalid.  But
			// an RLE at level 60 is valid, since the new level 61 is valid.  The current wording
			// of the rules requires that the RLE remain valid even if a previous LRE is invalid.
			// This keeps track of ignored LRE or LRO codes at level 60, so that the matching PDFs
			// will not try to pop the stack.
			int overflowAlmostCounter = 0;
		
			// This keeps track of ignored pushes at level 61 or higher, so that matching PDFs will
			// not try to pop the stack.
			int overflowCounter = 0;
		
			// Rule X1.
		
			// Keep the level separate from the value (level | override status flag) for ease of access.
			sbyte currentEmbeddingLevel = paragraphEmbeddingLevel;
			sbyte currentEmbeddingValue = paragraphEmbeddingLevel;
		
			// Loop through types, handling all remaining rules
			for (int i = 0; i < textLength; ++i)
			{
			
				embeddings[i] = currentEmbeddingValue;
			
				sbyte t = resultTypes[i];
			
				// Rules X2, X3, X4, X5
				switch (t)
				{
				
					case RLE: 
					case LRE: 
					case RLO: 
					case LRO: 
						if (overflowCounter == 0)
						{
							sbyte newLevel;
							if (t == RLE || t == RLO)
							{
								newLevel = (sbyte) ((currentEmbeddingLevel + 1) | 1); // least greater odd
							}
							else
							{
								// t == LRE || t == LRO
								newLevel = (sbyte) ((currentEmbeddingLevel + 2) & ~ 1); // least greater even
							}
						
							// If the new level is valid, push old embedding level and override status
							// No check for valid stack counter, since the level check suffices.
							if (newLevel < EXPLICIT_LEVEL_LIMIT)
							{
								embeddingValueStack[stackCounter] = currentEmbeddingValue;
								stackCounter++;
							
								currentEmbeddingLevel = newLevel;
								if (t == LRO || t == RLO)
								{
									// override
									unchecked
									{
										currentEmbeddingValue = (sbyte) ((byte)newLevel | 0x80);
									}
								}
								else
								{
									currentEmbeddingValue = newLevel;
								}
							
								// Adjust level of format mark (for expositional purposes only, this gets
								// removed later).
								embeddings[i] = currentEmbeddingValue;
								break;
							}
						
							// Otherwise new level is invalid, but a valid level can still be achieved if this
							// level is 60 and we encounter an RLE or RLO further on.  So record that we
							// 'almost' overflowed.
							if (currentEmbeddingLevel == 60)
							{
								overflowAlmostCounter++;
								break;
							}
						}
					
						// Otherwise old or new level is invalid.
						overflowCounter++;
						break;
				
				
					case PDF: 
					
						if (overflowCounter > 0)
						{
							--overflowCounter;
						}
						else if (overflowAlmostCounter > 0 && currentEmbeddingLevel != 61)
						{
							--overflowAlmostCounter;
						}
						else if (stackCounter > 0)
						{
							--stackCounter;
							currentEmbeddingValue = embeddingValueStack[stackCounter];
							currentEmbeddingLevel = (sbyte) (currentEmbeddingValue & 0x7f);
						}
						break;
				
				
					case B: 
						stackCounter = 0;
						overflowCounter = 0;
						overflowAlmostCounter = 0;
						currentEmbeddingLevel = paragraphEmbeddingLevel;
						currentEmbeddingValue = paragraphEmbeddingLevel;
					
						embeddings[i] = paragraphEmbeddingLevel;
						break;
				
				
					default: 
						break;
				
				}
			}
		
			return embeddings;
		}
	
	
		/// <summary> 3) resolving weak types
		/// Rules W1-W7.
		/// 
		/// Note that some weak types (EN, AN) remain after this processing is complete.
		/// </summary>
		private void  resolveWeakTypes(int start, int limit, sbyte level, sbyte sor, sbyte eor)
		{
		
			// on entry, only these types remain
            sbyte[] remainingTypes = new sbyte[]{L, R, AL, EN, ES, ET, AN, CS, B, S, WS, ON, NSM};
			assertOnly(start, limit, remainingTypes);
		
			// Rule W1.
			// Changes all NSMs.
			sbyte preceedingCharacterType = sor;
			for (int i = start; i < limit; ++i)
			{
				sbyte t = resultTypes[i];
				if (t == NSM)
				{
					resultTypes[i] = preceedingCharacterType;
				}
				else
				{
					preceedingCharacterType = t;
				}
			}
		
			// Rule W2.
			// EN does not change at the start of the run, because sor != AL.
			for (int i = start; i < limit; ++i)
			{
				if (resultTypes[i] == EN)
				{
					for (int j = i - 1; j >= start; --j)
					{
						sbyte t = resultTypes[j];
						if (t == L || t == R || t == AL)
						{
							if (t == AL)
							{
								resultTypes[i] = AN;
							}
							break;
						}
					}
				}
			}
		
			// Rule W3.
			for (int i = start; i < limit; ++i)
			{
				if (resultTypes[i] == AL)
				{
					resultTypes[i] = R;
				}
			}
		
			// Rule W4.
			// Since there must be values on both sides for this rule to have an
			// effect, the scan skips the first and last value.
			//
			// Although the scan proceeds left to right, and changes the type values
			// in a way that would appear to affect the computations later in the scan,
			// there is actually no problem.  A change in the current value can only 
			// affect the value to its immediate right, and only affect it if it is
			// ES or CS.  But the current value can only change if the value to its
			// right is not ES or CS.  Thus either the current value will not change,
			// or its change will have no effect on the remainder of the analysis.
		
			for (int i = start + 1; i < limit - 1; ++i)
			{
				if (resultTypes[i] == ES || resultTypes[i] == CS)
				{
					sbyte prevSepType = resultTypes[i - 1];
					sbyte succSepType = resultTypes[i + 1];
					if (prevSepType == EN && succSepType == EN)
					{
						resultTypes[i] = EN;
					}
					else if (resultTypes[i] == CS && prevSepType == AN && succSepType == AN)
					{
						resultTypes[i] = AN;
					}
				}
			}
		
			// Rule W5.
			for (int i = start; i < limit; ++i)
			{
				if (resultTypes[i] == ET)
				{
					// locate end of sequence
					int runstart = i;
                    sbyte[] rt = new sbyte[]{ET};
					int runlimit = findRunLimit(runstart, limit, rt);
				
					// check values at ends of sequence
					sbyte t = runstart == start?sor:resultTypes[runstart - 1];
				
					if (t != EN)
					{
						t = runlimit == limit?eor:resultTypes[runlimit];
					}
				
					if (t == EN)
					{
						setTypes(runstart, runlimit, EN);
					}
				
					// continue at end of sequence
					i = runlimit;
				}
			}
		
			// Rule W6.
			for (int i = start; i < limit; ++i)
			{
				sbyte t = resultTypes[i];
				if (t == ES || t == ET || t == CS)
				{
					resultTypes[i] = ON;
				}
			}
		
			// Rule W7.
			for (int i = start; i < limit; ++i)
			{
				if (resultTypes[i] == EN)
				{
					// set default if we reach start of run
					sbyte prevStrongType = sor;
					for (int j = i - 1; j >= start; --j)
					{
						sbyte t = resultTypes[j];
						if (t == L || t == R)
						{
							// AL's have been removed
							prevStrongType = t;
							break;
						}
					}
					if (prevStrongType == L)
					{
						resultTypes[i] = L;
					}
				}
			}
		}
	
		/// <summary> 6) resolving neutral types
		/// Rules N1-N2.
		/// </summary>
		private void  resolveNeutralTypes(int start, int limit, sbyte level, sbyte sor, sbyte eor)
		{
		
			// on entry, only these types can be in resultTypes
            sbyte[] ValidResultTypes = new sbyte[]{L, R, EN, AN, B, S, WS, ON};
			assertOnly(start, limit, ValidResultTypes);
		
			for (int i = start; i < limit; ++i)
			{
				sbyte t = resultTypes[i];
				if (t == WS || t == ON || t == B || t == S)
				{
					// find bounds of run of neutrals
					int runstart = i;
                    sbyte[] neutrals = new sbyte[]{B, S, WS, ON};
					int runlimit = findRunLimit(runstart, limit, neutrals);
				
					// determine effective types at ends of run
					sbyte leadingType;
					sbyte trailingType;
				
					if (runstart == start)
					{
						leadingType = sor;
					}
					else
					{
						leadingType = resultTypes[runstart - 1];
						if (leadingType == L || leadingType == R)
						{
							// found the strong type
						}
						else if (leadingType == AN)
						{
							leadingType = R;
						}
						else if (leadingType == EN)
						{
							// Since EN's with previous strong L types have been changed
							// to L in W7, the leadingType must be R.
							leadingType = R;
						}
					}
				
					if (runlimit == limit)
					{
						trailingType = eor;
					}
					else
					{
						trailingType = resultTypes[runlimit];
						if (trailingType == L || trailingType == R)
						{
							// found the strong type
						}
						else if (trailingType == AN)
						{
							trailingType = R;
						}
						else if (trailingType == EN)
						{
							trailingType = R;
						}
					}
				
					sbyte resolvedType;
					if (leadingType == trailingType)
					{
						// Rule N1.
						resolvedType = leadingType;
					}
					else
					{
						// Rule N2.
						// Notice the embedding level of the run is used, not
						// the paragraph embedding level.
						resolvedType = typeForLevel(level);
					}
				
					setTypes(runstart, runlimit, resolvedType);
				
					// skip over run of (former) neutrals
					i = runlimit;
				}
			}
		}
	
		/// <summary> 7) resolving implicit embedding levels
		/// Rules I1, I2.
		/// </summary>
		private void  resolveImplicitLevels(int start, int limit, sbyte level, sbyte sor, sbyte eor)
		{
		
			// on entry, only these types can be in resultTypes
            sbyte[] ValidResultTypes = new sbyte[]{L, R, EN, AN};
			assertOnly(start, limit, ValidResultTypes);
		
			if ((level & 1) == 0)
			{
				// even level
				for (int i = start; i < limit; ++i)
				{
					sbyte t = resultTypes[i];
					// Rule I1.
					if (t == L)
					{
						// no change
					}
					else if (t == R)
					{
						resultLevels[i] = (sbyte) (resultLevels[i] + 1);
					}
					else
					{
						// t == AN || t == EN
						resultLevels[i] = (sbyte) (resultLevels[i] + 2);
					}
				}
			}
			else
			{
				// odd level
				for (int i = start; i < limit; ++i)
				{
					sbyte t = resultTypes[i];
					// Rule I2.
					if (t == R)
					{
						// no change
					}
					else
					{
						// t == L || t == AN || t == EN
						resultLevels[i] = (sbyte) (resultLevels[i] + 1);
					}
				}
			}
		}
	
		//
		// Output
		//
	
		/// <summary> Return levels array breaking lines at offsets in linebreaks. <br></br>
		/// Rule L1.
		/// <p>
		/// The returned levels array contains the resolved level for each
		/// bidi code passed to the constructor.</p>
		/// <p>
		/// The linebreaks array must include at least one value.
		/// The values must be in strictly increasing order (no duplicates)
		/// between 1 and the length of the text, inclusive.  The last value
		/// must be the length of the text.
		/// </p>
		/// </summary>
		/// <param name="linebreaks">the offsets at which to break the paragraph
		/// </param>
		/// <returns> the resolved levels of the text
		/// </returns>
		public sbyte[] getLevels(int[] linebreaks)
		{
		
			// Note that since the previous processing has removed all 
			// P, S, and WS values from resultTypes, the values referred to
			// in these rules are the initial types, before any processing
			// has been applied (including processing of overrides).
			//
			// This example implementation has reinserted explicit format codes
			// and BN, in order that the levels array correspond to the 
			// initial text.  Their final placement is not normative.
			// These codes are treated like WS in this implementation,
			// so they don't interrupt sequences of WS.  
		
			validateLineBreaks(linebreaks, textLength);
		
			sbyte[] generated_var = new sbyte[resultLevels.Length];
			resultLevels.CopyTo(generated_var, 0);
			sbyte[] result = generated_var; // will be returned to caller
		
			// don't worry about linebreaks since if there is a break within
			// a series of WS values preceeding S, the linebreak itself
			// causes the reset.
			for (int i = 0; i < result.Length; ++i)
			{
				sbyte t = initialTypes[i];
				if (t == B || t == S)
				{
					// Rule L1, clauses one and two.
					result[i] = paragraphEmbeddingLevel;
				
					// Rule L1, clause three.
					for (int j = i - 1; j >= 0; --j)
					{
						if (isWhitespace(initialTypes[j]))
						{
							// including format codes
							result[j] = paragraphEmbeddingLevel;
						}
						else
						{
							break;
						}
					}
				}
			}
		
			// Rule L1, clause four.
			int start = 0;
			for (int i = 0; i < linebreaks.Length; ++i)
			{
				int limit = linebreaks[i];
				for (int j = limit - 1; j >= start; --j)
				{
					if (isWhitespace(initialTypes[j]))
					{
						// including format codes
						result[j] = paragraphEmbeddingLevel;
					}
					else
					{
						break;
					}
				}
			
				start = limit;
			}
		
			return result;
		}
	
		/// <summary> Return reordering array breaking lines at offsets in linebreaks.
		/// <p>
		/// The reordering array maps from a visual index to a logical index.
		/// Lines are concatenated from left to right.  So for example, the
		/// fifth character from the left on the third line is 
		/// <pre> getReordering(linebreaks)[linebreaks[1] + 4]</pre>
		/// (linebreaks[1] is the position after the last character of the 
		/// second line, which is also the index of the first character on the 
		/// third line, and adding four gets the fifth character from the left).</p>
		/// <p>
		/// The linebreaks array must include at least one value.
		/// The values must be in strictly increasing order (no duplicates)
		/// between 1 and the length of the text, inclusive.  The last value
		/// must be the length of the text.
		/// </p>
		/// </summary>
		/// <param name="linebreaks">the offsets at which to break the paragraph.
		/// </param>
		public int[] getReordering(int[] linebreaks)
		{
			validateLineBreaks(linebreaks, textLength);
		
			sbyte[] levels = getLevels(linebreaks);
		
			return computeMultilineReordering(levels, linebreaks);
		}
	
		/// <summary> Return multiline reordering array for a given level array.
		/// Reordering does not occur across a line break.
		/// </summary>
		private static int[] computeMultilineReordering(sbyte[] levels, int[] linebreaks)
		{
			int[] result = new int[levels.Length];
		
			int start = 0;
			for (int i = 0; i < linebreaks.Length; ++i)
			{
				int limit = linebreaks[i];
			
				sbyte[] templevels = new sbyte[limit - start];
				Array.Copy(levels, start, templevels, 0, templevels.Length);
			
				int[] temporder = computeReordering(templevels);
				for (int j = 0; j < temporder.Length; ++j)
				{
					result[start + j] = temporder[j] + start;
				}
			
				start = limit;
			}
		
			return result;
		}
	
		/// <summary> Return reordering array for a given level array.  This reorders a single line.
		/// The reordering is a visual to logical map.  For example,
		/// the leftmost char is string.charAt(order[0]).
		/// Rule L2.
		/// </summary>
		private static int[] computeReordering(sbyte[] levels)
		{
			int lineLength = levels.Length;
		
			int[] result = new int[lineLength];
		
			// initialize order
			for (int i = 0; i < lineLength; ++i)
			{
				result[i] = i;
			}
		
			// locate highest level found on line.
			// Note the rules say text, but no reordering across line bounds is performed,
			// so this is sufficient.
			sbyte highestLevel = 0;
			sbyte lowestOddLevel = 63;
			for (int i = 0; i < lineLength; ++i)
			{
				sbyte slevel = levels[i];
				if (slevel > highestLevel)
				{
					highestLevel = slevel;
				}
				if (((slevel & 1) != 0) && slevel < lowestOddLevel)
				{
					lowestOddLevel = slevel;
				}
			}
		
			for (int level = highestLevel; level >= lowestOddLevel; --level)
			{
				for (int i = 0; i < lineLength; ++i)
				{
					if (levels[i] >= level)
					{
						// find range of text at or above this level
						int start = i;
						int limit = i + 1;
						while (limit < lineLength && levels[limit] >= level)
						{
							++limit;
						}
					
						// reverse run
						for (int j = start, k = limit - 1; j < k; ++j, --k)
						{
							int temp = result[j];
							result[j] = result[k];
							result[k] = temp;
						}
					
						// skip to end of level run
						i = limit;
					}
				}
			}
		
			return result;
		}
	
		// --- internal utilities -------------------------------------------------
	
		/// <summary> Return true if the type is considered a whitespace type for the line break rules.</summary>
		private static bool isWhitespace(sbyte biditype)
		{
			switch (biditype)
			{
			
				case LRE: 
				case RLE: 
				case LRO: 
				case RLO: 
				case PDF: 
				case BN: 
				case WS: 
					return true;
			
				default: 
					return false;
			
			}
		}
	
		/// <summary> Return the strong type (L or R) corresponding to the level.</summary>
		private static sbyte typeForLevel(int level)
		{
			return ((level & 0x1) == 0)?L:R;
		}
	
		/// <summary> Return the limit of the run starting at index that includes only resultTypes in validSet.
		/// This checks the value at index, and will return index if that value is not in validSet.
		/// </summary>
		private int findRunLimit(int index, int limit, sbyte[] validSet)
		{
			--index;
			//UPGRADE_NOTE: Label 'loop' was moved. 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="jlca1014"'
			while (++index < limit)
			{
				sbyte t = resultTypes[index];
				for (int i = 0; i < validSet.Length; ++i)
				{
					if (t == validSet[i])
					{
						//UPGRADE_NOTE: Labeled continue statement was changed to a goto statement. 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="jlca1015"'
						goto loop;
					}
				}
				// didn't find a match in validSet
				return index;
				//UPGRADE_NOTE: Label 'loop' was moved. 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="jlca1014"'

			loop: ;
			}
			return limit;
		}
	
		/// <summary> Set resultTypes from start up to (but not including) limit to newType.</summary>
		private void  setTypes(int start, int limit, sbyte newType)
		{
			for (int i = start; i < limit; ++i)
			{
				resultTypes[i] = newType;
			}
		}
	
		/// <summary> Set resultLevels from start up to (but not including) limit to newLevel.</summary>
		private void  setLevels(int start, int limit, sbyte newLevel)
		{
			for (int i = start; i < limit; ++i)
			{
				resultLevels[i] = newLevel;
			}
		}
	
		// --- algorithm internal validation --------------------------------------
	
		/// <summary> Algorithm validation.
		/// Assert that all values in resultTypes are in the provided set.
		/// </summary>
		private void  assertOnly(int start, int limit, sbyte[] codes)
		{
			//UPGRADE_NOTE: Label 'loop1' was moved. 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="jlca1014"'
			for (int i = start; i < limit; ++i)
			{
				sbyte t = resultTypes[i];
				for (int j = 0; j < codes.Length; ++j)
				{
					if (t == codes[j])
					{
						//UPGRADE_NOTE: Labeled continue statement was changed to a goto statement. 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="jlca1015"'
						goto loop1;
					}
				}
			
				throw new FlexCelException("invalid bidi code " + typenames[t] + " present in assertOnly at position " + i);
				//UPGRADE_NOTE: Label 'loop1' was moved. 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="jlca1014"'

			loop1: ;
			}
		}
		
		/// <summary> Throw exception if line breaks array is invalid.</summary>
		private static void  validateLineBreaks(int[] linebreaks, int textLength)
		{
			int prev = 0;
			for (int i = 0; i < linebreaks.Length; ++i)
			{
				int next = linebreaks[i];
				if (next <= prev)
				{
					throw new System.ArgumentException("bad linebreak: " + next + " at index: " + i);
				}
				prev = next;
			}
			if (prev != textLength)
			{
				throw new System.ArgumentException("last linebreak must be at " + textLength);
			}
		}	

		#region Added utilities
		/// <summary>
		/// Both the starting range and the value.
		/// </summary>
		private static readonly int[] UnicodeTypes = 
		{
			0	,	 BN	,	9	,	 S	,	10	,	 B	,	11	,	 S	,	12	,	 WS	,	13	,	 B	,
			14	,	 BN	,	28	,	 B	,	31	,	 S	,	32	,	 WS	,	33	,	 ON	,	35	,	 ET	,
			38	,	 ON	,	43	,	 ET	,	44	,	 CS	,	45	,	 ET	,	46	,	 CS	,	47	,	 ES	,
			48	,	 EN	,	58	,	 CS	,	59	,	 ON	,	65	,	 L	,	91	,	 ON	,	97	,	 L	,
			123	,	 ON	,	127	,	 BN	,	133	,	 B	,	134	,	 BN	,	160	,	 CS	,
			161	,	 ON	,	162	,	 ET	,	166	,	 ON	,	170	,	 L	,	171	,	 ON	,
			176	,	 ET	,	178	,	 EN	,	180	,	 ON	,	181	,	 L	,	182	,	 ON	,
			185	,	 EN	,	186	,	 L	,	187	,	 ON	,	192	,	 L	,	215	,	 ON	,
			216	,	 L	,	247	,	 ON	,	248	,	 L	,	697	,	 ON	,	699	,	 L	,
			706	,	 ON	,	720	,	 L	,	722	,	 ON	,	736	,	 L	,	741	,	 ON	,
			750	,	 L	,	751	,	 ON	,	768	,	 NSM	,	856	,	 L	,	861	,	 NSM	,
			880	,	 L	,	884	,	 ON	,	886	,	 L	,	894	,	 ON	,	895	,	 L	,
			900	,	 ON	,	902	,	 L	,	903	,	 ON	,	904	,	 L	,	1014	,	 ON	,
			1015	,	 L	,	1155	,	 NSM	,	1159	,	 L	,	1160	,	 NSM	,
			1162	,	 L	,	1418	,	 ON	,	1419	,	 L	,	1425	,	 NSM	,
			1442	,	 L	,	1443	,	 NSM	,	1466	,	 L	,	1467	,	 NSM	,
			1470	,	 R	,	1471	,	 NSM	,	1472	,	 R	,	1473	,	 NSM	,
			1475	,	 R	,	1476	,	 NSM	,	1477	,	 L	,	1488	,	 R	,
			1515	,	 L	,	1520	,	 R	,	1525	,	 L	,	1536	,	 AL	,
			1540	,	 L	,	1548	,	 CS	,	1549	,	 AL	,	1550	,	 ON	,
			1552	,	 NSM	,	1558	,	 L	,	1563	,	 AL	,	1564	,	 L	,
			1567	,	 AL	,	1568	,	 L	,	1569	,	 AL	,	1595	,	 L	,
			1600	,	 AL	,	1611	,	 NSM	,	1625	,	 L	,	1632	,	 AN	,
			1642	,	 ET	,	1643	,	 AN	,	1645	,	 AL	,	1648	,	 NSM	,
			1649	,	 AL	,	1750	,	 NSM	,	1757	,	 AL	,	1758	,	 NSM	,
			1765	,	 AL	,	1767	,	 NSM	,	1769	,	 ON	,	1770	,	 NSM	,
			1774	,	 AL	,	1776	,	 EN	,	1786	,	 AL	,	1806	,	 L	,
			1807	,	 BN	,	1808	,	 AL	,	1809	,	 NSM	,	1810	,	 AL	,
			1840	,	 NSM	,	1867	,	 L	,	1869	,	 AL	,	1872	,	 L	,
			1920	,	 AL	,	1958	,	 NSM	,	1969	,	 AL	,	1970	,	 L	,
			2305	,	 NSM	,	2307	,	 L	,	2364	,	 NSM	,	2365	,	 L	,
			2369	,	 NSM	,	2377	,	 L	,	2381	,	 NSM	,	2382	,	 L	,
			2385	,	 NSM	,	2389	,	 L	,	2402	,	 NSM	,	2404	,	 L	,
			2433	,	 NSM	,	2434	,	 L	,	2492	,	 NSM	,	2493	,	 L	,
			2497	,	 NSM	,	2501	,	 L	,	2509	,	 NSM	,	2510	,	 L	,
			2530	,	 NSM	,	2532	,	 L	,	2546	,	 ET	,	2548	,	 L	,
			2561	,	 NSM	,	2563	,	 L	,	2620	,	 NSM	,	2621	,	 L	,
			2625	,	 NSM	,	2627	,	 L	,	2631	,	 NSM	,	2633	,	 L	,
			2635	,	 NSM	,	2638	,	 L	,	2672	,	 NSM	,	2674	,	 L	,
			2689	,	 NSM	,	2691	,	 L	,	2748	,	 NSM	,	2749	,	 L	,
			2753	,	 NSM	,	2758	,	 L	,	2759	,	 NSM	,	2761	,	 L	,
			2765	,	 NSM	,	2766	,	 L	,	2786	,	 NSM	,	2788	,	 L	,
			2801	,	 ET	,	2802	,	 L	,	2817	,	 NSM	,	2818	,	 L	,
			2876	,	 NSM	,	2877	,	 L	,	2879	,	 NSM	,	2880	,	 L	,
			2881	,	 NSM	,	2884	,	 L	,	2893	,	 NSM	,	2894	,	 L	,
			2902	,	 NSM	,	2903	,	 L	,	2946	,	 NSM	,	2947	,	 L	,
			3008	,	 NSM	,	3009	,	 L	,	3021	,	 NSM	,	3022	,	 L	,
			3059	,	 ON	,	3065	,	 ET	,	3066	,	 ON	,	3067	,	 L	,
			3134	,	 NSM	,	3137	,	 L	,	3142	,	 NSM	,	3145	,	 L	,
			3146	,	 NSM	,	3150	,	 L	,	3157	,	 NSM	,	3159	,	 L	,
			3260	,	 NSM	,	3261	,	 L	,	3276	,	 NSM	,	3278	,	 L	,
			3393	,	 NSM	,	3396	,	 L	,	3405	,	 NSM	,	3406	,	 L	,
			3530	,	 NSM	,	3531	,	 L	,	3538	,	 NSM	,	3541	,	 L	,
			3542	,	 NSM	,	3543	,	 L	,	3633	,	 NSM	,	3634	,	 L	,
			3636	,	 NSM	,	3643	,	 L	,	3647	,	 ET	,	3648	,	 L	,
			3655	,	 NSM	,	3663	,	 L	,	3761	,	 NSM	,	3762	,	 L	,
			3764	,	 NSM	,	3770	,	 L	,	3771	,	 NSM	,	3773	,	 L	,
			3784	,	 NSM	,	3790	,	 L	,	3864	,	 NSM	,	3866	,	 L	,
			3893	,	 NSM	,	3894	,	 L	,	3895	,	 NSM	,	3896	,	 L	,
			3897	,	 NSM	,	3898	,	 ON	,	3902	,	 L	,	3953	,	 NSM	,
			3967	,	 L	,	3968	,	 NSM	,	3973	,	 L	,	3974	,	 NSM	,
			3976	,	 L	,	3984	,	 NSM	,	3992	,	 L	,	3993	,	 NSM	,
			4029	,	 L	,	4038	,	 NSM	,	4039	,	 L	,	4141	,	 NSM	,
			4145	,	 L	,	4146	,	 NSM	,	4147	,	 L	,	4150	,	 NSM	,
			4152	,	 L	,	4153	,	 NSM	,	4154	,	 L	,	4184	,	 NSM	,
			4186	,	 L	,	5760	,	 WS	,	5761	,	 L	,	5787	,	 ON	,
			5789	,	 L	,	5906	,	 NSM	,	5909	,	 L	,	5938	,	 NSM	,
			5941	,	 L	,	5970	,	 NSM	,	5972	,	 L	,	6002	,	 NSM	,
			6004	,	 L	,	6071	,	 NSM	,	6078	,	 L	,	6086	,	 NSM	,
			6087	,	 L	,	6089	,	 NSM	,	6100	,	 L	,	6107	,	 ET	,
			6108	,	 L	,	6109	,	 NSM	,	6110	,	 L	,	6128	,	 ON	,
			6138	,	 L	,	6144	,	 ON	,	6155	,	 NSM	,	6158	,	 WS	,
			6159	,	 L	,	6313	,	 NSM	,	6314	,	 L	,	6432	,	 NSM	,
			6435	,	 L	,	6439	,	 NSM	,	6444	,	 L	,	6450	,	 NSM	,
			6451	,	 L	,	6457	,	 NSM	,	6460	,	 L	,	6464	,	 ON	,
			6465	,	 L	,	6468	,	 ON	,	6470	,	 L	,	6624	,	 ON	,
			6656	,	 L	,	8125	,	 ON	,	8126	,	 L	,	8127	,	 ON	,
			8130	,	 L	,	8141	,	 ON	,	8144	,	 L	,	8157	,	 ON	,
			8160	,	 L	,	8173	,	 ON	,	8176	,	 L	,	8189	,	 ON	,
			8191	,	 L	,	8192	,	 WS	,	8203	,	 BN	,	8206	,	 L	,
			8207	,	 R	,	8208	,	 ON	,	8232	,	 WS	,	8233	,	 B	,
			8234	,	 LRE	,	8235	,	 RLE	,	8236	,	 PDF	,	8237	,	 LRO	,
			8238	,	 RLO	,	8239	,	 WS	,	8240	,	 ET	,	8245	,	 ON	,
			8277	,	 L	,	8279	,	 ON	,	8280	,	 L	,	8287	,	 WS	,
			8288	,	 BN	,	8292	,	 L	,	8298	,	 BN	,	8304	,	 EN	,
			8305	,	 L	,	8308	,	 EN	,	8314	,	 ET	,	8316	,	 ON	,
			8319	,	 L	,	8320	,	 EN	,	8330	,	 ET	,	8332	,	 ON	,
			8335	,	 L	,	8352	,	 ET	,	8370	,	 L	,	8400	,	 NSM	,
			8427	,	 L	,	8448	,	 ON	,	8450	,	 L	,	8451	,	 ON	,
			8455	,	 L	,	8456	,	 ON	,	8458	,	 L	,	8468	,	 ON	,
			8469	,	 L	,	8470	,	 ON	,	8473	,	 L	,	8478	,	 ON	,
			8484	,	 L	,	8485	,	 ON	,	8486	,	 L	,	8487	,	 ON	,
			8488	,	 L	,	8489	,	 ON	,	8490	,	 L	,	8494	,	 ET	,
			8495	,	 L	,	8498	,	 ON	,	8499	,	 L	,	8506	,	 ON	,
			8508	,	 L	,	8512	,	 ON	,	8517	,	 L	,	8522	,	 ON	,
			8524	,	 L	,	8531	,	 ON	,	8544	,	 L	,	8592	,	 ON	,
			8722	,	 ET	,	8724	,	 ON	,	9014	,	 L	,	9083	,	 ON	,
			9109	,	 L	,	9110	,	 ON	,	9169	,	 L	,	9216	,	 ON	,
			9255	,	 L	,	9280	,	 ON	,	9291	,	 L	,	9312	,	 EN	,
			9372	,	 L	,	9450	,	 EN	,	9451	,	 ON	,	9752	,	 L	,
			9753	,	 ON	,	9854	,	 L	,	9856	,	 ON	,	9874	,	 L	,
			9888	,	 ON	,	9890	,	 L	,	9985	,	 ON	,	9989	,	 L	,
			9990	,	 ON	,	9994	,	 L	,	9996	,	 ON	,	10024	,	 L	,
			10025	,	 ON	,	10060	,	 L	,	10061	,	 ON	,	10062	,	 L	,
			10063	,	 ON	,	10067	,	 L	,	10070	,	 ON	,	10071	,	 L	,
			10072	,	 ON	,	10079	,	 L	,	10081	,	 ON	,	10133	,	 L	,
			10136	,	 ON	,	10160	,	 L	,	10161	,	 ON	,	10175	,	 L	,
			10192	,	 ON	,	10220	,	 L	,	10224	,	 ON	,	11022	,	 L	,
			11904	,	 ON	,	11930	,	 L	,	11931	,	 ON	,	12020	,	 L	,
			12032	,	 ON	,	12246	,	 L	,	12272	,	 ON	,	12284	,	 L	,
			12288	,	 WS	,	12289	,	 ON	,	12293	,	 L	,	12296	,	 ON	,
			12321	,	 L	,	12330	,	 NSM	,	12336	,	 ON	,	12337	,	 L	,
			12342	,	 ON	,	12344	,	 L	,	12349	,	 ON	,	12352	,	 L	,
			12441	,	 NSM	,	12443	,	 ON	,	12445	,	 L	,	12448	,	 ON	,
			12449	,	 L	,	12539	,	 ON	,	12540	,	 L	,	12829	,	 ON	,
			12831	,	 L	,	12880	,	 ON	,	12896	,	 L	,	12924	,	 ON	,
			12926	,	 L	,	12977	,	 ON	,	12992	,	 L	,	13004	,	 ON	,
			13008	,	 L	,	13175	,	 ON	,	13179	,	 L	,	13278	,	 ON	,
			13280	,	 L	,	13311	,	 ON	,	13312	,	 L	,	19904	,	 ON	,
			19968	,	 L	,	42128	,	 ON	,	42183	,	 L	,	64285	,	 R	,
			64286	,	 NSM	,	64287	,	 R	,	64297	,	 ET	,	64298	,	 R	,
			64311	,	 L	,	64312	,	 R	,	64317	,	 L	,	64318	,	 R	,
			64319	,	 L	,	64320	,	 R	,	64322	,	 L	,	64323	,	 R	,
			64325	,	 L	,	64326	,	 R	,	64336	,	 AL	,	64434	,	 L	,
			64467	,	 AL	,	64830	,	 ON	,	64832	,	 L	,	64848	,	 AL	,
			64912	,	 L	,	64914	,	 AL	,	64968	,	 L	,	65008	,	 AL	,
			65021	,	 ON	,	65022	,	 L	,	65024	,	 NSM	,	65040	,	 L	,
			65056	,	 NSM	,	65060	,	 L	,	65072	,	 ON	,	65104	,	 CS	,
			65105	,	 ON	,	65106	,	 CS	,	65107	,	 L	,	65108	,	 ON	,
			65109	,	 CS	,	65110	,	 ON	,	65119	,	 ET	,	65120	,	 ON	,
			65122	,	 ET	,	65124	,	 ON	,	65127	,	 L	,	65128	,	 ON	,
			65129	,	 ET	,	65131	,	 ON	,	65132	,	 L	,	65136	,	 AL	,
			65141	,	 L	,	65142	,	 AL	,	65277	,	 L	,	65279	,	 BN	,
			65280	,	 L	,	65281	,	 ON	,	65283	,	 ET	,	65286	,	 ON	,
			65291	,	 ET	,	65292	,	 CS	,	65293	,	 ET	,	65294	,	 CS	,
			65295	,	 ES	,	65296	,	 EN	,	65306	,	 CS	,	65307	,	 ON	,
			65313	,	 L	,	65339	,	 ON	,	65345	,	 L	,	65371	,	 ON	,
			65382	,	 L	,	65504	,	 ET	,	65506	,	 ON	,	65509	,	 ET	,
			65511	,	 L	,	65512	,	 ON	,	65519	,	 L	,	65529	,	 BN	,
			65532	,	 ON	,	65534	,	 L	,   0xFFFF  ,    L
		};

		private static sbyte[] CreateAllTypes()
		{
			sbyte[] Result = new sbyte[0xFFFF];
			int itpos =0;
			for (int i=0; i<UnicodeTypes.Length-2;i+=2)
			{
				while (itpos< UnicodeTypes[i+2]) 
				{
					Result[itpos] = (sbyte)UnicodeTypes[i+1];
					itpos++;
				}
			}

			return Result;
		}
		#endregion
	}
}
