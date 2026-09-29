using System;
using System.Net.Mail;
using System.Diagnostics;
using System.Threading.Tasks;
using MailKit.Net.Imap;
using MailKit.Search;
using Microsoft.Extensions.AI;
using MimeKit;
using ModelContextProtocol.Server;
using System.ComponentModel;
using ModelContextProtocol;

namespace ProtonEmailOrganizer.Tools;

public class ProtonEmailTools
{
    private static readonly string BridgeUser = 
        Environment.GetEnvironmentVariable("BRIDGE_USER") 
        ?? Environment.GetEnvironmentVariable("BRIDGE_USER", EnvironmentVariableTarget.User) 
        ?? Environment.GetEnvironmentVariable("BRIDGE_USER", EnvironmentVariableTarget.Machine) 
        ?? "bogus@proton.me";

    private static readonly string BridgePass = 
        Environment.GetEnvironmentVariable("BRIDGE_PASS") 
        ?? Environment.GetEnvironmentVariable("BRIDGE_PASS", EnvironmentVariableTarget.User) 
        ?? Environment.GetEnvironmentVariable("BRIDGE_PASS", EnvironmentVariableTarget.Machine) 
        ?? "password";
    private const string BridgeHost = "127.0.0.1";
    private const int ImapPort = 1143;
    private const int SmtpPort = 1025;

    [McpServerTool]
    [Description("Search emails in Proton Mail via the local bridge")]
    public async Task<string> SearchEmailsAsync(
        [Description ("The search text query to match against email body or subject")] string query,
        [Description("The mailbox folder to search (defaults to INBOX)")] string folder = "INBOX")

    {
        Debug.WriteLine($"Username: {BridgeUser} - Searching emails in folder: {folder} with query: {query}");

        using var client = new ImapClient();

        // Proton Bridge uses self-signed certificates locally; bypass trust checks
        client.ServerCertificateValidationCallback = (s, c, h, e) => true; 

        await client.ConnectAsync(BridgeHost, ImapPort, false);
        await client.AuthenticateAsync(BridgeUser, BridgePass);
        var mailbox = string.Equals(folder, "INBOX", StringComparison.OrdinalIgnoreCase)
            ? client.Inbox
            : await client.GetFolderAsync(folder);
        await mailbox.OpenAsync(MailKit.FolderAccess.ReadOnly);
        var results = await mailbox.SearchAsync(SearchQuery.BodyContains(query)
            .Or(SearchQuery.SubjectContains(query))
            .Or(SearchQuery.FromContains(query)));

        var messages = new List<string>();
        foreach (var uid in results)
        {
            var message = await mailbox.GetMessageAsync(uid);
            messages.Add($"Subject: {message.Subject}\nFrom: {message.From}\nDate: {message.Date:yyyy-MM-dd HH:mm:ss zzz}");
        }
        await client.DisconnectAsync(true);

        return messages.Count == 0
            ? $"No emails matching '{query}' were found in {folder}."
            : $"Found {messages.Count} email(s) in {folder}:\n\n{string.Join("\n\n", messages)}";
    }

}