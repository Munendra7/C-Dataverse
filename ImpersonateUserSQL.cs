using System;
using System.Runtime.InteropServices;
using System.Security;

class Program
{
    static void Main(string[] args)
    {
        // Specify the username and password directly in the code
        string userName = "domain\\username"; // Replace with the desired Windows user
        string passwordStr = "user_password"; // Replace with the user's password

        // Convert the password to a SecureString
        SecureString password = new SecureString();
        foreach (char c in passwordStr)
        {
            password.AppendChar(c);
        }

        // Pass the username and password for impersonation
        using (ImpersonationContext context = new ImpersonationContext(userName, password))
        {
            // Your code here will run under the security context of the specified user
            Console.WriteLine($"Current user: {System.Security.Principal.WindowsIdentity.GetCurrent().Name}");
        }

        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
}

// Helper class for impersonation
public class ImpersonationContext : IDisposable
{
    private readonly System.Security.Principal.WindowsImpersonationContext _impersonationContext;

    public ImpersonationContext(string userName, SecureString password)
    {
        IntPtr token = IntPtr.Zero;

        try
        {
            // Convert SecureString to plain text password
            string passwordStr = new System.Net.NetworkCredential(string.Empty, password).Password;

            // Logon the user
            bool success = LogonUser(userName, ".", passwordStr, LogonType.Interactive, LogonProvider.Default, out token);
            if (!success)
            {
                throw new InvalidOperationException("Failed to logon user.");
            }

            // Start impersonating
            _impersonationContext = System.Security.Principal.WindowsIdentity.Impersonate(token);
        }
        finally
        {
            // Close the token handle
            if (token != IntPtr.Zero)
            {
                CloseHandle(token);
            }
        }
    }

    public void Dispose()
    {
        _impersonationContext?.Undo();
    }

    // Windows API declarations
    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool LogonUser(string lpszUsername, string lpszDomain, string lpszPassword,
        LogonType dwLogonType, LogonProvider dwLogonProvider, out IntPtr phToken);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool CloseHandle(IntPtr hObject);

    private enum LogonType : int
    {
        Interactive = 2,
        Network = 3,
        Batch = 4,
        Service = 5,
        NetworkCleartext = 8,
        NewCredentials = 9,
    }

    private enum LogonProvider : int
    {
        Default = 0,
        WinNT35 = 1,
        WinNT40 = 2,
        WinNT50 = 3,
    }
}
