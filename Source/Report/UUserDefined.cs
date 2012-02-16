using System;

namespace FlexCel.Report
{
	/// <summary>
	/// A class used to define a FlexCel user function, that you can call from a report.
	/// </summary>
	/// <remarks>
	/// To Create a User function:
	/// <list type="number">
	/// <item>Create a new class derived from TFlexCelUserFunction.</item>
	/// <item>Override the method <see cref="Evaluate"/>.</item>
	/// <item>Add the new user function to the report using <see cref="FlexCelReport.SetUserFunction"/>.</item>
	/// </list>
	/// </remarks>
	/// <example>
	/// To define an user function that returns "One" for param=1; "Two" for param=2 and "Unknown" on other case:
	/// 1) Define the class:
	/// <code>
	///     public class TMyUserFunction: TFlexCelUserFunction
    ///     {
    ///         public override object Evaluate(object[] parameters)
    ///         {
    ///             if (parameters==null || parameters.Length>1)
    ///                 throw new ArgumentException("Invalid number of params for user defined function \"MyUserFunction");
    ///             int p= Convert.ToInt32(parameters[0]);
    ///     
    ///             switch (p)
    ///             {
    ///                 case 1: return "One";
    ///                 case 2: return "Two";
    ///             }
    ///             return "Unknown";
    ///         }
    ///     
    ///     }
	/// </code>
	/// 2) Add the function to the report.
	/// <code>
	/// FReport.SetUserFunction("MF", new TMyUserFunction());
	/// </code>
	/// 3) Now, you can write "&lt;#MF(1)&gt;" on a template, and it will be replaced by "One".
	/// </example>
	public abstract class TFlexCelUserFunction
	{
        /// <summary>
        /// Override this method on a derived class to implement your own defined function.
        /// </summary>
        /// <param name="parameters">An array of objects.</param>
        /// <returns>The derived class should return the value of the function on the return parameter.</returns>
        public abstract object Evaluate(object[] parameters);
	}

}

