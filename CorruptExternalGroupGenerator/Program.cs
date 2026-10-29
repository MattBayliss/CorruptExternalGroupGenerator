using Azure.Identity;
using CorruptExternalGroupGenerator;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph.Models.ExternalConnectors;
using Microsoft.Graph.Models.ODataErrors;
using static System.Diagnostics.Debug;

/*
 * repeat until a 'Corrupt' group is generated
 * repeat:
 *   - delete a group
 *   - create a group
 * until create a group produces a 409 error (NameAlreadyExists + UnclassifiedConflict)
 */

var configuration = new ConfigurationBuilder()
    .AddJsonFile("appsettings.json")
    .AddUserSecrets<GraphService>()
    .Build();

var tenantId = configuration["tenantId"];
var clientId = configuration["clientId"];
var clientSecret = configuration["clientSecret"];
var externalConnectionId = configuration["connectionId"];

// ensure all configs are set
List<string?> configs = [tenantId, clientId, clientSecret, externalConnectionId];
Assert(!configs.Any(string.IsNullOrEmpty));

var credentials = new ClientSecretCredential(tenantId, clientId, clientSecret);
var graphService = new GraphService(credentials, externalConnectionId!);

using var cancellationTokenSource = new CancellationTokenSource();
var ct = cancellationTokenSource.Token;

var connection = await graphService.GetConnectionAsync(ct);
if (connection?.Id == null)
{
    var created = await graphService.CreateConnectionAsync(ct);
    Assert(created);

    await graphService.SetSchemaAsync(SchemaHelper.BasicSchema, ct);

    var connectionInDraft = true;
    while (connectionInDraft)
    {
        await Task.Delay(TimeSpan.FromSeconds(10));
        connection = await graphService.GetConnectionAsync(ct);
        connectionInDraft = connection is ExternalConnection extCon ? extCon.State == ConnectionState.Draft : true;
    }
    Console.WriteLine($"Connection created: {externalConnectionId}");
}

int max_attempts = 1000;
int attempt = 0;

// just any valid ExternalGroupId will do
var groupId = Path.GetFileNameWithoutExtension(Path.GetRandomFileName());

ExternalGroup? corruptGroup = null;

ExternalGroup? testGroup = await graphService.CreateExternalGroupAsync(groupId, ct);
Assert(testGroup != null);

string? lastException = null;

while (corruptGroup is null && attempt < max_attempts)
{
    attempt++;
    Console.Clear();
    Console.WriteLine($"Trying to generate a 'corrupt' group. GroupId: `{groupId}`. Attempt: {attempt} of {max_attempts}");

    if(lastException != null)
    {
        Console.WriteLine();
        Console.WriteLine("Last error:");
        Console.WriteLine(lastException.ToString());
        lastException = null;
    }

    try
    {
        await graphService.DeleteExternalGroupAsync(groupId, ct);
        testGroup = await graphService.CreateExternalGroupAsync(groupId, ct);
    }
    /* 
      The "corrupt" group case we're watching for:
    
        {
          "error": {
            "code": "NameAlreadyExists",
            "message": "The specified resource name already exists",
            "innerError": {
              "code": "UnclassifiedConflict",
              "message": "Conflict occurred but could not be classified into a known category.",
            ...
     */
    catch (ODataError odex) when (
        odex.ResponseStatusCode == 409
        && odex.Error?.Code == "NameAlreadyExists"
        && (odex.Error?.InnerError?.AdditionalData?.TryGetValue("code", out var codeObj) == true)
        && codeObj is string innerCode && innerCode == "UnclassifiedConflict")
    {
        corruptGroup = testGroup;
    }
    catch(ODataError odex)
    {
        // just grabbing this to output to the console
        var innerCodeMessage = odex.Error?.InnerError?.AdditionalData?.TryGetValue("code", out var codeObj) == true
            && codeObj is string innerCode ? $" Inner Error [{innerCode}]" : "";

        lastException = $"Status code [{odex.ResponseStatusCode}], Error: [{odex.Error?.Code?? ""}]{innerCodeMessage}";
    }
}

Console.WriteLine();

if (corruptGroup != null)
{
    Console.WriteLine($"CORRUPT GROUP GENERATED: {corruptGroup.Id}");
} else
{
    Console.WriteLine($"FAILED TO GENERATE CORRUPT GROUP IN {max_attempts} ATTEMPTS. Problem fixed?");
}

