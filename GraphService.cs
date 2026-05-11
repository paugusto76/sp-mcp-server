using Azure.Identity;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Beta;
using Microsoft.Graph.Beta.Models;

public class GraphService
{
    private readonly ILogger<GraphService> _logger;
    private readonly GraphServiceClient _graphClient;
    private DateTime _graphClientLastSet;

    private string _clientId = "";
    private string _clientSecret = "";
    private string _tenantId = "";

    public GraphService(ILogger<GraphService> logger)
    {
        _logger = logger;

        _graphClient = InitGraph();

        _logger.LogError("GraphService initialized.");
    }

    internal GraphServiceClient InitGraph()
    {
        string[] scopes = new string[] { "https://graph.microsoft.com/.default" };

        var options = new TokenCredentialOptions
        {
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
        };

        var clientSecretCredential = new ClientSecretCredential(
            _tenantId,
            _clientId,
            _clientSecret,
            options);

        var graphClient = new GraphServiceClient(clientSecretCredential, scopes);
        _graphClientLastSet = DateTime.Now;

        return graphClient;
    }

    public async Task<List<Site>?> SearchSiteAsync(string searchQuery)
    {
        try
        {
            _logger.LogError("Searching for sites with query: {SearchQuery}", searchQuery);
            var sites = await _graphClient.Sites.GetAsync(rc => rc.QueryParameters.Search = searchQuery);

            return sites?.Value;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error searching for site with query: {SearchQuery}", searchQuery);
            throw;
        }
    }

    public async Task<List<Site>?> GetTopSitesAsync()
    {
        try
        {
            _logger.LogError("Retrieving top 20 sites.");
            var sites = await _graphClient.Sites.GetAsync(rc => rc.QueryParameters.Top = 20);

            return sites?.Value;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving top sites.");
            throw;
        }
    }

    internal async Task<Site?> GetSiteByPathAsync(string hostname, string serverRelativePath)
    {
        try
        {
            _logger.LogError("Retrieving site with hostname: {Hostname} and path: {ServerRelativePath}", hostname, serverRelativePath);
            string sitePath = $"{hostname}:{(serverRelativePath.StartsWith("/") ? serverRelativePath : "/" + serverRelativePath)}";

            var site = await _graphClient.Sites[sitePath].GetAsync();
            return site;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving site with hostname: {Hostname} and path: {ServerRelativePath}", hostname, serverRelativePath);
            throw;
        }
    }

    internal async Task<List<Site>?> GetSubsitesAsync(string siteId)
    {
        try
        {
            _logger.LogError("Retrieving subsites for site ID: {SiteId}", siteId);
            var subsites = await _graphClient.Sites[siteId].Sites.GetAsync();
            return subsites?.Value;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving subsites for site ID: {SiteId}", siteId);
            throw;
        }
    }
}