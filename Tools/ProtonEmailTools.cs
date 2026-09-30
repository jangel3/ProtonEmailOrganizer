using System;
using System.Diagnostics;
using System.Threading.Tasks;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
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
        [Description("The search text query to match against email body or subject")] string query,
        [Description("The mailbox folder to search, or ALL to search every folder (defaults to INBOX)")] string folder = "INBOX",
        [Description("Folder names or full paths to skip when folder is ALL")] string[]? excludeFolders = null)

    {
        Debug.WriteLine($"Username: {BridgeUser} - Searching emails in folder: {folder} with query: {query}");

        using var client = new ImapClient();

        // Proton Bridge uses self-signed certificates locally; bypass trust checks
        client.ServerCertificateValidationCallback = (s, c, h, e) => true;

        await client.ConnectAsync(BridgeHost, ImapPort, false);
        await client.AuthenticateAsync(BridgeUser, BridgePass);
        var messages = new List<string>();
        var mailboxes = string.Equals(folder, "ALL", StringComparison.OrdinalIgnoreCase)
            ? await GetAllFoldersAsync(client)
            : new List<IMailFolder>
            {
                string.Equals(folder, "INBOX", StringComparison.OrdinalIgnoreCase)
                    ? client.Inbox
                    : await client.GetFolderAsync(folder)
            };
        var excluded = excludeFolders?.ToHashSet(StringComparer.OrdinalIgnoreCase) ?? [];

        foreach (var mailbox in mailboxes)
        {
            if (excluded.Contains(mailbox.Name) || excluded.Contains(mailbox.FullName))
                continue;

            await mailbox.OpenAsync(MailKit.FolderAccess.ReadOnly);
            var results = await mailbox.SearchAsync(SearchQuery.BodyContains(query)
                .Or(SearchQuery.SubjectContains(query))
                .Or(SearchQuery.FromContains(query)));

            foreach (var uid in results)
            {
                var message = await mailbox.GetMessageAsync(uid);
                messages.Add($"Folder: {mailbox.FullName}\nSubject: {message.Subject}\nFrom: {message.From}\nDate: {message.Date:yyyy-MM-dd HH:mm:ss zzz}");
            }
        }
        await client.DisconnectAsync(true);

        return messages.Count == 0
            ? $"No emails matching '{query}' were found in {folder}."
            : $"Found {messages.Count} email(s) in {folder}:\n\n{string.Join("\n\n", messages)}";
    }

    [McpServerTool]
    [Description("Collect unique email addresses found in message sender and recipient headers across all Proton Mail folders")]
    public async Task<string> GetContactEmailsAsync()
    {
        using var client = new ImapClient();

        // Proton Bridge uses self-signed certificates locally; bypass trust checks
        client.ServerCertificateValidationCallback = (s, c, h, e) => true;

        await client.ConnectAsync(BridgeHost, ImapPort, false);
        await client.AuthenticateAsync(BridgeUser, BridgePass);

        var emailAddresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var mailbox in await GetAllFoldersAsync(client))
        {
            await mailbox.OpenAsync(MailKit.FolderAccess.ReadOnly);
            var messageIds = await mailbox.SearchAsync(SearchQuery.All);
            if (messageIds.Count == 0)
                continue;

            var summaries = await mailbox.FetchAsync(messageIds, MessageSummaryItems.Envelope);
            foreach (var summary in summaries)
            {
                var envelope = summary.Envelope;
                if (envelope is null)
                    continue;

                AddEmailAddresses(envelope.From, emailAddresses);
                AddEmailAddresses(envelope.ReplyTo, emailAddresses);
                AddEmailAddresses(envelope.To, emailAddresses);
                AddEmailAddresses(envelope.Cc, emailAddresses);
                AddEmailAddresses(envelope.Bcc, emailAddresses);
            }
        }

        await client.DisconnectAsync(true);

        return emailAddresses.Count == 0
            ? "No email addresses were found in message headers."
            : $"Found {emailAddresses.Count} unique email address(es):\n{string.Join("\n", emailAddresses.OrderBy(address => address, StringComparer.OrdinalIgnoreCase))}";
    }

    private static void AddEmailAddresses(InternetAddressList? addresses, HashSet<string> emailAddresses)
    {
        if (addresses is null)
            return;

        foreach (var address in addresses)
        {
            if (address is MailboxAddress mailbox)
                emailAddresses.Add(mailbox.Address);
            else if (address is GroupAddress group)
                AddEmailAddresses(group.Members, emailAddresses);
        }
    }

    private static async Task<List<IMailFolder>> GetAllFoldersAsync(ImapClient client)
    {
        var folders = new List<IMailFolder>();
        foreach (var personalNamespace in client.PersonalNamespaces)
        {
            var topLevelFolders = await client.GetFoldersAsync(personalNamespace, false);
            foreach (var folder in topLevelFolders)
                await AddFolderAndSubfoldersAsync(folder, folders);
        }

        return folders;
    }

    private static async Task AddFolderAndSubfoldersAsync(IMailFolder folder, List<IMailFolder> folders)
    {
        folders.Add(folder);
        foreach (var subfolder in await folder.GetSubfoldersAsync(false))
            await AddFolderAndSubfoldersAsync(subfolder, folders);
    }

    [McpServerTool]
    [Description("Move emails matching a sender from one folder to another")]
    public async Task<string> MoveEmailsAsync(
        [Description("The sender of the email(s) to move")] string from,
        [Description("The mailbox folder to search (defaults to INBOX)")] string sourceFolder = "INBOX",
        [Description("The mailbox folder to move matching emails into")] string destinationFolder = "Archive")
    {

        using var client = new ImapClient();

        // Proton Bridge uses self-signed certificates locally; bypass trust checks
        client.ServerCertificateValidationCallback = (s, c, h, e) => true;

        await client.ConnectAsync(BridgeHost, ImapPort, false);
        await client.AuthenticateAsync(BridgeUser, BridgePass);
        var folders = await GetAllFoldersAsync(client);

        try
        {
            var source = string.Equals(sourceFolder, "INBOX", StringComparison.OrdinalIgnoreCase)
                ? client.Inbox
                : await client.GetFolderAsync(sourceFolder);
            var destination = await client.GetFolderAsync(destinationFolder);

            await source.OpenAsync(MailKit.FolderAccess.ReadWrite);
            var results = await source.SearchAsync(SearchQuery.FromContains(from));

            if (results.Count > 0)
                await source.MoveToAsync(results, destination);

            await client.DisconnectAsync(true);

            return results.Count == 0
                ? $"No emails matching subject '{from}' were found in {sourceFolder}."
                : $"Moved {results.Count} email(s) matching subject '{from}' from {sourceFolder} to {destinationFolder}.";
        }
        catch (Exception ex)
        {
            return $"Available folders are {string.Join(", ", folders.Select(f => f.FullName))}. An error occurred: {ex.Message}";
        }
    }

    [McpServerTool]
    [Description("Move old emails from one folder to another")]
    public async Task<string> MoveOldEmailsAsync(
        [Description("The mailbox folder to search (defaults to INBOX)")] string sourceFolder = "INBOX",
        [Description("The mailbox folder to move old emails into")] string destinationFolder = "Archive",
        [Description("The age in days of emails to move")] int ageInDays = 30)
    {
        using var client = new ImapClient();

        // Proton Bridge uses self-signed certificates locally; bypass trust checks
        client.ServerCertificateValidationCallback = (s, c, h, e) => true;

        await client.ConnectAsync(BridgeHost, ImapPort, false);
        await client.AuthenticateAsync(BridgeUser, BridgePass);
        var folders = await GetAllFoldersAsync(client);

        try
        {
            var source = string.Equals(sourceFolder, "INBOX", StringComparison.OrdinalIgnoreCase)
                ? client.Inbox
                : await client.GetFolderAsync(sourceFolder);
            var destination = await client.GetFolderAsync(destinationFolder);

            await source.OpenAsync(MailKit.FolderAccess.ReadWrite);
            var results = await source.SearchAsync(SearchQuery.DeliveredBefore(DateTime.UtcNow.AddDays(-ageInDays)));

            if (results.Count > 0)
                await source.MoveToAsync(results, destination);

            await client.DisconnectAsync(true);

            return results.Count == 0
                ? $"No emails older than {ageInDays} days were found in {sourceFolder}."
                : $"Moved {results.Count} old email(s) from {sourceFolder} to {destinationFolder}.";
        }
        catch (Exception ex)
        {
            return $"Available folders are {string.Join(", ", folders.Select(f => f.FullName))}. An error occurred: {ex.Message}";
        }
    }

    [McpServerTool]
    [Description("Move emails matching a list of senders from one folder to another")]
    public async Task<string> MoveEmailContactsAsync(
        [Description("The list of senders of the email(s) to move")] List<string> fromList,
        [Description("The mailbox folder to search (defaults to INBOX)")] string sourceFolder = "INBOX",
        [Description("The mailbox folder to move matching emails into")] string destinationFolder = "Archive")
    {
        using var client = new ImapClient();

        // Proton Bridge uses self-signed certificates locally; bypass trust checks
        client.ServerCertificateValidationCallback = (s, c, h, e) => true;

        await client.ConnectAsync(BridgeHost, ImapPort, false);
        await client.AuthenticateAsync(BridgeUser, BridgePass);
        var folders = await GetAllFoldersAsync(client);

        try
        {
            var source = string.Equals(sourceFolder, "INBOX", StringComparison.OrdinalIgnoreCase)
                ? client.Inbox
                : await client.GetFolderAsync(sourceFolder);
            var destination = await client.GetFolderAsync(destinationFolder);

            await source.OpenAsync(MailKit.FolderAccess.ReadWrite);
            var results = new List<UniqueId>();
            foreach (var from in fromList)
            {
                var searchResults = await source.SearchAsync(SearchQuery.FromContains(from));
                results.AddRange(searchResults);
            }

            if (results.Count > 0)
                await source.MoveToAsync(results, destination);

            await client.DisconnectAsync(true);

            return results.Count == 0
                ? $"No emails matching senders '{string.Join(", ", fromList)}' were found in {sourceFolder}."
                : $"Moved {results.Count} email(s) matching senders '{string.Join(", ", fromList)}' from {sourceFolder} to {destinationFolder}.";
        }
        catch (Exception ex)
        {
            return $"Available folders are {string.Join(", ", folders.Select(f => f.FullName))}. An error occurred: {ex.Message}";
        }
    }
}