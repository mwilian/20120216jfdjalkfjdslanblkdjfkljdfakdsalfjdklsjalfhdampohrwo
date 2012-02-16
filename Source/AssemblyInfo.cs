using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Security;
using System.Resources;


[assembly: AssemblyCompany("TVC Software")]
[assembly: AssemblyProduct("TVC Studio for .NET")]
[assembly: AssemblyConfiguration("Registered")]

[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]		

[assembly: AssemblyVersion("5.5.1.0")]
[assembly: AssemblyCopyright("(c) 2002 - 2011 TVC Software")]

[assembly: NeutralResourcesLanguage("en-US")]


#if ((!COMPACTFRAMEWORK || FRAMEWORK20) && !MONOTOUCH)
[assembly:CLSCompliant(true)]
#endif
#if (!COMPACTFRAMEWORK)
[assembly:ComVisible(false)]

//Security permissions
#if (!SILVERLIGHT)
#if (FRAMEWORK40)
[assembly: SecurityRules(SecurityRuleSet.Level2)]
[assembly: AllowPartiallyTrustedCallers()] //this has a different effect than in .net 2/3, that's why we write it 2 times. It means a different thing. http://blogs.msdn.com/shawnfa/archive/2009/11/12/differences-between-the-security-rule-sets.aspx
#else
[assembly: AllowPartiallyTrustedCallers()]
[assembly: 	SecurityPermission(SecurityAction.RequestMinimum, Execution=true)]
//[assembly: 	FileIOPermissionAttribute(SecurityAction.RequestOptional, Unrestricted=true)]  We need fulltrust for calling oledb and vj#
#endif
#endif
#endif
