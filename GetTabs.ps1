$path = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
Add-Type -Path "$path\UIAutomationClient.dll"
Add-Type -Path "$path\UIAutomationTypes.dll"

$Assem =@(
        "$path\UIAutomationClient.dll",
        "$path\UIAutomationTypes.dll"
         ) 

$code = @"
using System;
using System.Text;
using System.Diagnostics;
using System.Windows.Automation;

namespace NETShell
{
    public class Program
    {     
        public static string GetChromeMainUrl()
        {
            Process[] procsChrome = Process.GetProcessesByName("chrome");
            StringBuilder result = new StringBuilder();

            foreach (Process chrome in procsChrome)
            {
                // the chrome process must have a window
                if (chrome.MainWindowHandle == IntPtr.Zero)
                {
                    continue;
                }

                // find the automation element
                AutomationElement elm = AutomationElement.FromHandle(chrome.MainWindowHandle);
                AutomationElement elmUrlBar = elm.FindFirst(TreeScope.Descendants,
                  new PropertyCondition(AutomationElement.NameProperty, "Address and search bar"));

                // if it can be found, get the value from the URL bar
                if (elmUrlBar != null)
                {
                    AutomationPattern[] patterns = elmUrlBar.GetSupportedPatterns();
                    if (patterns.Length > 0)
                    {
                        ValuePattern val = (ValuePattern)elmUrlBar.GetCurrentPattern(patterns[0]);

                        result.AppendLine(val.Current.Value);                        
                    }
                }
            }
            return result.ToString();
        }     
    }
}
"@

try
{
   [NETShell.Program] -is [type] | Out-Null
}
catch
{
    Add-Type -ReferencedAssemblies $Assem -TypeDefinition $code -Language CSharp
}
	

$url = [NETShell.Program]::GetChromeMainUrl()
$index = $url.IndexOf("1998")
if($index -ge 0){ $ticketNumber = $url.SubString($index, 5) }else{$ticketNumber = "Error"}
Write-Host "1989-$ticketNumber"







