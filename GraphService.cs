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



    public GraphService(ILogger<GraphService> logger)
    {
        _logger = logger;

        // _graphClient = InitGraph();
        // _graphClient = InitGraphWithDelegated();

        _logger.LogError("GraphService initialized.");
    }

    internal GraphServiceClient InitGraph()
    {

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

    internal GraphServiceClient InitGraphWithDelegated()
    {
        var scopes = new[] { "Sites.Read.All", "Files.Read.All" };   // or specific scopes like "Sites.ReadWrite.All"

        var credential = new InteractiveBrowserCredential(new InteractiveBrowserCredentialOptions
        {
            ClientId = _clientId,
            TenantId = _tenantId,
            // RedirectUri = "http://localhost",   // usually not needed
        });

        var graphClient = new GraphServiceClient(credential, scopes);
        _graphClientLastSet = DateTime.Now;

        return graphClient;
    }

    private GraphServiceClient UseGraphOnBehalfOfUser(string userBearerToken)
    {

        // Create a fresh credential for this request (or reuse with token provider)
        var credential = new OnBehalfOfCredential(
            tenantId: _tenantId,
            clientId: _clientId,
            clientSecret: _clientSecret,
            userAssertion: userBearerToken   // ← The token from MCP client
        );

        // Option 1: Create client per-request (simplest)
        var graphClient = new GraphServiceClient(credential, scopes);

        return graphClient;
    }


    public async Task<List<Site>?> SearchSiteAsync(string searchQuery, string userBearerToken)
    {
        try
        {
            _logger.LogError("Searching for sites with query: {SearchQuery}", searchQuery);
            var graphClient = UseGraphOnBehalfOfUser(userBearerToken);
            var sites = await graphClient.Sites.GetAsync(rc => rc.QueryParameters.Search = searchQuery);

            return sites?.Value;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error searching for site with query: {SearchQuery}", searchQuery);
            throw;
        }
    }

    public async Task<List<Site>?> GetTopSitesAsync(string userBearerToken)
    {
        try
        {
            _logger.LogError("Retrieving top 20 sites.");
            var graphClient = UseGraphOnBehalfOfUser(userBearerToken);
            var sites = await graphClient.Sites.GetAsync(rc => rc.QueryParameters.Top = 20);

            return sites?.Value;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving top sites.");
            throw;
        }
    }

    internal async Task<Site?> GetSiteByPathAsync(string hostname, string serverRelativePath, string userBearerToken)
    {
        try
        {
            _logger.LogError("Retrieving site with hostname: {Hostname} and path: {ServerRelativePath}", hostname, serverRelativePath);
            string sitePath = $"{hostname}:{(serverRelativePath.StartsWith("/") ? serverRelativePath : "/" + serverRelativePath)}";

            var graphClient = UseGraphOnBehalfOfUser(userBearerToken);
            var site = await graphClient.Sites[sitePath].GetAsync();
            return site;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving site with hostname: {Hostname} and path: {ServerRelativePath}", hostname, serverRelativePath);
            throw;
        }
    }

    internal async Task<List<Site>?> GetSubsitesAsync(string siteId, string userBearerToken)
    {
        try
        {
            _logger.LogError("Retrieving subsites for site ID: {SiteId}", siteId);
            var graphClient = UseGraphOnBehalfOfUser(userBearerToken);
            var subsites = await graphClient.Sites[siteId].Sites.GetAsync();
            return subsites?.Value;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error retrieving subsites for site ID: {SiteId}", siteId);
            throw;
        }
    }
}