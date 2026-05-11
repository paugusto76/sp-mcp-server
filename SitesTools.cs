using Microsoft.Graph.Beta.Models;
using ModelContextProtocol.Server;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Text.Json;

namespace SPOMcpServer.Tools;

[McpServerToolType]
public static class SitesTools
{
    [McpServerTool, Description("Find SharePoint sites accessible to the user. Returns specific sites matching a search query, or the top 20 relevant sites if no query is provided.")]
    public static async Task<JsonDocument> FindSiteAsync(GraphService graphService, string searchQuery, string userBearerToken)
    {
        List<Site>? sites;

        if (string.IsNullOrWhiteSpace(searchQuery))
        {
            sites = await graphService.GetTopSitesAsync(userBearerToken);
        }
        else
        {
            sites = await graphService.SearchSiteAsync(searchQuery, userBearerToken);
        }

        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        return JsonDocument.Parse(JsonSerializer.Serialize(sites, options));
    }

    [McpServerTool, Description("Resolve a SharePoint site using its exact hostname and server-relative path. Use only when you have the complete site URL structure. Use FindSite when only a site name is known.")]
    public static async Task<JsonDocument> GetSiteByPath(GraphService graphService, string hostname, string serverRelativePath, string userBearerToken)
    {
        try
        {
            Debug.Assert(!string.IsNullOrWhiteSpace(hostname), "Hostname must be provided.");
            Debug.Assert(!string.IsNullOrWhiteSpace(serverRelativePath), "Server-relative path must be provided.");

            var site = await graphService.GetSiteByPathAsync(hostname, serverRelativePath, userBearerToken);

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            return JsonDocument.Parse(JsonSerializer.Serialize(site, options));
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error retrieving site with hostname '{hostname}' and path '{serverRelativePath}': {ex.Message}", ex);
        }
    }

    [McpServerTool, Description("List all subsites (child sites) of a SharePoint site.")]
    public static async Task<JsonDocument> GetSubsites(GraphService graphService, string siteId, string userBearerToken)
    {
        try
        {
            Debug.Assert(!string.IsNullOrWhiteSpace(siteId), "Site ID must be provided.");

            var subsites = await graphService.GetSubsitesAsync(siteId, userBearerToken);

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };

            return JsonDocument.Parse(JsonSerializer.Serialize(subsites, options));
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error retrieving subsites for site ID '{siteId}': {ex.Message}", ex);
        }
    }
}