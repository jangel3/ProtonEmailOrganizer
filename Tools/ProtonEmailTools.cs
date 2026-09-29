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
        ?? "stop@proton.me";

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
    public async Task SearchEmailsAsync(
        [Description ("The search text query to match against email body or subject")] string query,
        [Description("The mailbox folder to search (defaults to INBOX)")] string folder = "INBOX")

    {
        Debug.WriteLine($"Username: {BridgeUser} - Searching emails in folder: {folder} with query: {query}");
        var fileStream = File.Open("search_log.txt", FileMode.Append);
        fileStream.Write(System.Text.Encoding.UTF8.GetBytes($"Username: {BridgeUser} - Searching emails in folder: {folder} with query: {query}{Environment.NewLine}"));
        fileStream.Close();
        
        using var client = new ImapClient();

        // Proton Bridge uses self-signed certificates locally; bypass trust checks
        client.ServerCertificateValidationCallback = (s, c, h, e) => true; 

        await client.ConnectAsync(BridgeHost, ImapPort, false);
        await client.AuthenticateAsync(BridgeUser, BridgePass);
        var inbox = client.Inbox;
        await inbox.OpenAsync(MailKit.FolderAccess.ReadOnly);
        var results = await inbox.SearchAsync(SearchQuery.BodyContains(query)
            .Or(SearchQuery.SubjectContains(query))
            .Or(SearchQuery.FromContains(query)));
        foreach (var uid in results)
        {
            var message = await inbox.GetMessageAsync(uid);
            Console.WriteLine($"Subject: {message.Subject}");
        }
        await client.DisconnectAsync(true);
    }

}