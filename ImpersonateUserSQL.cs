using System;
using System.Runtime.InteropServices;
using System.Security;

public class Program
{
    static void Main(string[] args)
    {
        // User to impersonate (Windows account)
        string userName = "domain\\username"; // Replace with the desired Windows user

        Console.WriteLine("Enter password for the user:");
        SecureString password = GetSecurePassword();

        // Pass the user name and password for impersonation
        using (ImpersonationContext context = new ImpersonationContext(userName, password))
        {
            // Your code here will run under the security context of the specified user
            Console.WriteLine($"Current user: {System.Security.Principal.WindowsIdentity.GetCurrent().Name}");
        }
    }

    // Helper method to securely prompt for password
    private static SecureString GetSecurePassword()
    {
        SecureString securePassword = new SecureString();
        ConsoleKeyInfo key;

        do
        {
            key = Console.ReadKey(true);

            // Ignore any key other than Enter (when finished typing)
            if (key.Key != ConsoleKey.Enter)
            {
                // Append the character to the SecureString
                securePassword.AppendChar(key.KeyChar);
                Console.Write("*");
            }
        } while (key.Key != ConsoleKey.Enter);

        Console.WriteLine(); // Add newline after password prompt

        // Make the SecureString read-only
        securePassword.MakeReadOnly();
        return securePassword;
    }
}

// Helper class for impersonation
public class ImpersonationContext : IDisposable
{
    private readonly System.Security.Principal.WindowsImpersonationContext _impersonationContext;

    public ImpersonationContext(string userName, SecureString password)
    {
        IntPtr token = IntPtr.Zero;
        IntPtr tokenDuplicate = IntPtr.Zero;

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

            // Duplicate the token for impersonation
            success = DuplicateToken(token, SecurityImpersonationLevel.Impersonation, out tokenDuplicate);
            if (!success)
            {
                throw new InvalidOperationException("Failed to duplicate token.");
            }

            // Start impersonating
            _impersonationContext = System.Security.Principal.WindowsIdentity.Impersonate(tokenDuplicate);
        }
        finally
        {
            // Close the token handles
            if (token != IntPtr.Zero)
            {
                CloseHandle(token);
            }
            if (tokenDuplicate != IntPtr.Zero)
            {
                CloseHandle(tokenDuplicate);
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

    [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private extern static bool DuplicateToken(IntPtr ExistingTokenHandle, SecurityImpersonationLevel ImpersonationLevel, out IntPtr DuplicateTokenHandle);

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

    private enum SecurityImpersonationLevel : int
    {
        SecurityAnonymous = 0,
        SecurityIdentification = 1,
        SecurityImpersonation = 2,
        SecurityDelegation = 3,
    }
}
