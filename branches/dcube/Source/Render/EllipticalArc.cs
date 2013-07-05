using System;
using FlexCel.Core;

// This code is adapted from EllipticArc.java, http://www.spaceroots.com/downloads.html

#region Original Copyright Notice: (c) 2003-2004, Luc Maisonobe
// Copyright (c) 2003-2004, Luc Maisonobe
// All rights reserved.
// 
// Redistribution and use in source and binary forms, with
// or without modification, are permitted provided that
// the following conditions are met:
// 
//    Redistributions of source code must retain the
//    above copyright notice, this list of conditions and
//    the following disclaimer. 
//    Redistributions in binary form must reproduce the
//    above copyright notice, this list of conditions and
//    the following disclaimer in the documentation
//    and/or other materials provided with the
//    distribution. 
//    Neither the names of spaceroots.org, spaceroots.com
//    nor the names of their contributors may be used to
//    endorse or promote products derived from this
//    software without specific prior written permission. 
// 
// THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND
// CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED
// WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
// WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A
// PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL
// THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY
// DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
// CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
// PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF
// USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
// HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER
// IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
// NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE
// USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
// POSSIBILITY OF SUCH DAMAGE.
#endregion

namespace FlexCel.Render
{
	/// <summary>
	/// Draws an Elliptical arc using Bezier curves
	/// </summary>
	internal sealed class TEllipticalArc
	{
		private TEllipticalArc()
		{
		}

		// coefficients for error estimation
		// while using cubic Bezier curves for approximation
		// 0 < b/a < 1/4
		private static readonly double[][][] coeffs3Low = new double[][][] 
		{
			new double[][]
			{
				new double[]{  3.85268,   -21.229,      -0.330434,    0.0127842  },
				new double[]{ -1.61486,     0.706564,    0.225945,    0.263682   },
				new double[]{ -0.910164,    0.388383,    0.00551445,  0.00671814 },
				new double[]{ -0.630184,    0.192402,    0.0098871,   0.0102527  }
			},
			new double[][]
			{
				new double[]{ -0.162211,    9.94329,     0.13723,     0.0124084  },
				new double[]{ -0.253135,    0.00187735,  0.0230286,   0.01264    },
				new double[]{ -0.0695069,  -0.0437594,   0.0120636,   0.0163087  },
				new double[]{ -0.0328856,  -0.00926032, -0.00173573,  0.00527385 }
		   }
		};

		// coefficients for error estimation
		// while using cubic Bezier curves for approximation
		// 1/4 <= b/a <= 1
		private static readonly double[][][] coeffs3High = new double[][][] 
		{
			new double[][]
			{
				new double[]{  0.0899116, -19.2349,     -4.11711,     0.183362   },
				new double[]{  0.138148,   -1.45804,     1.32044,     1.38474    },
				new double[]{  0.230903,   -0.450262,    0.219963,    0.414038   },
				new double[]{  0.0590565,  -0.101062,    0.0430592,   0.0204699  }
			},
			new double[][]
			{
				new double[]{  0.0164649,   9.89394,     0.0919496,   0.00760802 },
				new double[]{  0.0191603,  -0.0322058,   0.0134667,  -0.0825018  },
				new double[]{  0.0156192,  -0.017535,    0.00326508, -0.228157   },
				new double[]{ -0.0236752,   0.0405821,  -0.0173086,   0.176187   }
			}
		};

		// safety factor to convert the "best" error approximation
		// into a "max bound" error
		private static readonly double[] safety3 = new double[] 
		{
			0.001, 4.98, 0.207, 0.0067
		};

		private static void AddPoint(TPointF[] Result, int index, double x, double y)
		{
			Result[index].X = (float) x;
			Result[index].Y = (float) y;
		}

		private static double rationalFunction(double x, double[] c) 
		{
			return (x * (x * c[0] + c[1]) + c[2]) / (x + c[3]);
		}

		private static double estimateError(double a, double b, double etaA, double etaB) 
		{
			double eta  = 0.5 * (etaA + etaB);

			double x    = b / a;
			double dEta = etaB - etaA;
			double cos2 = Math.Cos(2 * eta);
			double cos4 = Math.Cos(4 * eta);
			double cos6 = Math.Cos(6 * eta);

			// select the right coeficients set according to degree and b/a
			double[][][] coeffs;
			double[] safety;

			coeffs = (x < 0.25) ? coeffs3Low : coeffs3High;
			safety = safety3;

			double c0 = rationalFunction(x, coeffs[0][0])
				+ cos2 * rationalFunction(x, coeffs[0][1])
				+ cos4 * rationalFunction(x, coeffs[0][2])
				+ cos6 * rationalFunction(x, coeffs[0][3]);

			double c1 = rationalFunction(x, coeffs[1][0])
				+ cos2 * rationalFunction(x, coeffs[1][1])
				+ cos4 * rationalFunction(x, coeffs[1][2])
				+ cos6 * rationalFunction(x, coeffs[1][3]);

			return rationalFunction(x, safety) * a * Math.Exp(c0 + c1 * dEta);
		}

		/// <summary>
		/// Returns the array of points needed to create an elliptical arc with bezier curves.
		/// </summary>
		/// <param name="cx">Center coordinate (x).</param>
		/// <param name="cy">Center coordinate (y).</param>
		/// <param name="a">Major radius.</param>
		/// <param name="b">Minor radius.</param>
        /// <param name="theta">Angle of the major axis respect to the x axis. (Radians)</param>
        /// <param name="lambda1">Starting angle of the arc. (Radians)</param>
        /// <param name="lambda2">Ending angle of the arc. (Radians)</param>
		/// <returns></returns>
		internal static TPointF[] GetPoints(double cx, double cy, double a, double b, double theta, double lambda1, double lambda2)
		{
			if (a <= 0 || b <= 0) return new TPointF[0];

			double eta1 = Math.Atan2(Math.Sin(lambda1) / b, Math.Cos(lambda1) / a);
			double eta2 = Math.Atan2(Math.Sin(lambda2) / b,	Math.Cos(lambda2) / a);
			double cosTheta   = Math.Cos(theta);
			double sinTheta   = Math.Sin(theta);
			double defaultFlatness = 0.5; // half a pixel
			double twoPi = 2 * Math.PI;

			// make sure we have eta1 <= eta2 <= eta1 + 2 PI
			eta2 -= twoPi * Math.Floor((eta2 - eta1) / twoPi);

			// the preceding correction fails if we have exactly et2 - eta1 = 2 PI
			// it reduces the interval to zero length
			if ((lambda2 - lambda1 > Math.PI) && (eta2 - eta1 < Math.PI)) 
			{
				eta2 += 2 * Math.PI;
			}

			
			// find the number of Bezier curves needed
			bool found = false;
			int n = 1;
			while ((! found) && (n < 1024)) 
			{
				double dEta = (eta2 - eta1) / n;
				if (dEta <= 0.5 * Math.PI) 
				{
					double etaB = eta1;
					found = true;
					for (int i = 0; found && (i < n); ++i) 
					{
						double etaA = etaB;
						etaB += dEta;
						found = (estimateError(a, b, etaA, etaB) <= defaultFlatness);
					}
				}
				n = n << 1;
			}
		{

			TPointF[] Result = new TPointF[1 + 3*n];
			double dEta = (eta2 - eta1) / n;
			double etaB = eta1;

			double cosEtaB  = Math.Cos(etaB);
			double sinEtaB  = Math.Sin(etaB);
			double aCosEtaB = a * cosEtaB;
			double bSinEtaB = b * sinEtaB;
			double aSinEtaB = a * sinEtaB;
			double bCosEtaB = b * cosEtaB;
			double xB       = cx + aCosEtaB * cosTheta - bSinEtaB * sinTheta;
			double yB       = cy + aCosEtaB * sinTheta + bSinEtaB * cosTheta;
			double xBDot    = -aSinEtaB * cosTheta - bCosEtaB * sinTheta;
			double yBDot    = -aSinEtaB * sinTheta + bCosEtaB * cosTheta;

			AddPoint(Result, 0, xB, yB);

			double t     = Math.Tan(0.5 * dEta);
			double alpha = Math.Sin(dEta) * (Math.Sqrt(4 + 3 * t * t) - 1) / 3;

			int rp = 1;
			for (int i = 0; i < n; i++) 
			{
				double xA    = xB;
				double yA    = yB;
				double xADot = xBDot;
				double yADot = yBDot;

				etaB    += dEta;
				cosEtaB  = Math.Cos(etaB);
				sinEtaB  = Math.Sin(etaB);
				aCosEtaB = a * cosEtaB;
				bSinEtaB = b * sinEtaB;
				aSinEtaB = a * sinEtaB;
				bCosEtaB = b * cosEtaB;
				xB       = cx + aCosEtaB * cosTheta - bSinEtaB * sinTheta;
				yB       = cy + aCosEtaB * sinTheta + bSinEtaB * cosTheta;
				xBDot    = -aSinEtaB * cosTheta - bCosEtaB * sinTheta;
				yBDot    = -aSinEtaB * sinTheta + bCosEtaB * cosTheta;

				AddPoint(Result, rp, (xA + alpha * xADot), (yA + alpha * yADot));rp++;
				AddPoint(Result, rp, (xB - alpha * xBDot), (yB - alpha * yBDot));rp++;
				AddPoint(Result, rp,  xB,                   yB);rp++;

			}
			return Result;
		}
		}
	}
}
