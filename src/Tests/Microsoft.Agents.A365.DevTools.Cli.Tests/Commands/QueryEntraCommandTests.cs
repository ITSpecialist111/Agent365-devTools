// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using Microsoft.Extensions.Logging;
using Microsoft.Agents.A365.DevTools.Cli.Commands;
using Microsoft.Agents.A365.DevTools.Cli.Services;
using NSubstitute;
using System.CommandLine;
using System.Linq;
using System.Text.Json;
using Xunit;

namespace Microsoft.Agents.A365.DevTools.Cli.Tests.Commands;

public class QueryEntraCommandTests
{
    private readonly ILogger<QueryEntraCommand> _mockLogger;
    private readonly IConfigService _mockConfigService;
    private readonly CommandExecutor _mockExecutor;
    private readonly GraphApiService _mockGraphApiService;
    private readonly AgentBlueprintService _mockBlueprintService;

    public QueryEntraCommandTests()
    {
        _mockLogger = Substitute.For<ILogger<QueryEntraCommand>>();
        _mockConfigService = Substitute.For<IConfigService>();
        // Create CommandExecutor with a mock logger dependency
        var mockExecutorLogger = Substitute.For<ILogger<CommandExecutor>>();
        _mockExecutor = new CommandExecutor(mockExecutorLogger);
        _mockGraphApiService = Substitute.For<GraphApiService>(Substitute.For<ILogger<GraphApiService>>(), _mockExecutor);
        _mockBlueprintService = Substitute.ForPartsOf<AgentBlueprintService>(Substitute.For<ILogger<AgentBlueprintService>>(), _mockGraphApiService);
    }

    [Fact]
    public void QueryEntraCommand_Should_Be_Created()
    {
        // Act
        var command = QueryEntraCommand.CreateCommand(
            _mockLogger,
            _mockConfigService,
            _mockExecutor,
            _mockGraphApiService, _mockBlueprintService);

        // Assert
        Assert.NotNull(command);
        Assert.Equal("query-entra", command.Name);
        Assert.Equal("Query Microsoft Entra ID for agent information (scopes, permissions, consent status)", command.Description);
    }

    [Fact]
    public void QueryEntraCommand_Should_Have_Correct_Subcommands()
    {
        // Arrange
        var command = QueryEntraCommand.CreateCommand(
            _mockLogger,
            _mockConfigService,
            _mockExecutor,
            _mockGraphApiService, _mockBlueprintService);

        // Assert
        Assert.Equal(3, command.Subcommands.Count);
        Assert.Contains(command.Subcommands, c => c.Name == "blueprints");
        Assert.Contains(command.Subcommands, c => c.Name == "blueprint-scopes");
        Assert.Contains(command.Subcommands, c => c.Name == "instance-scopes");
    }

    [Fact]
    public void QueryEntraCommand_Should_Have_Blueprints_Subcommand()
    {
        // Arrange
        var command = QueryEntraCommand.CreateCommand(
            _mockLogger,
            _mockConfigService,
            _mockExecutor,
            _mockGraphApiService, _mockBlueprintService);

        // Act
        var blueprintsSubcommand = command.Subcommands.FirstOrDefault(c => c.Name == "blueprints");

        // Assert
        Assert.NotNull(blueprintsSubcommand);
        Assert.Equal("List Agent Identity Blueprints in the tenant", blueprintsSubcommand.Description);
        Assert.Contains("list-blueprints", blueprintsSubcommand!.Aliases);
    }

    [Fact]
    public void QueryEntraCommand_Should_Have_BlueprintScopes_Subcommand()
    {
        // Arrange
        var command = QueryEntraCommand.CreateCommand(
            _mockLogger,
            _mockConfigService,
            _mockExecutor,
            _mockGraphApiService, _mockBlueprintService);

        // Act
        var blueprintScopesSubcommand = command.Subcommands.FirstOrDefault(c => c.Name == "blueprint-scopes");

        // Assert
        Assert.NotNull(blueprintScopesSubcommand);
        Assert.Equal("List configured scopes and consent status for the agent blueprint", blueprintScopesSubcommand.Description);
    }

    [Fact]
    public void QueryEntraCommand_Should_Have_InstanceScopes_Subcommand()
    {
        // Arrange
        var command = QueryEntraCommand.CreateCommand(
            _mockLogger,
            _mockConfigService,
            _mockExecutor,
            _mockGraphApiService, _mockBlueprintService);

        // Act
        var instanceScopesSubcommand = command.Subcommands.FirstOrDefault(c => c.Name == "instance-scopes");

        // Assert
        Assert.NotNull(instanceScopesSubcommand);
        Assert.Equal("List configured scopes and consent status for the agent instance", instanceScopesSubcommand.Description);
    }

    [Fact]
    public async Task QueryEntraCommand_Blueprints_WithTenantId_ListsBlueprints()
    {
        // Arrange
        using var blueprintsDoc = JsonDocument.Parse("""
            {
              "value": [
                {
                  "id": "blueprint-object-id",
                  "appId": "blueprint-app-id",
                  "displayName": "Contoso Agent Blueprint",
                  "createdDateTime": "2026-05-14T12:00:00Z"
                }
              ]
            }
            """);

        _mockGraphApiService
            .GraphGetAsync(
                Arg.Is("tenant-id"),
                Arg.Is<string>(path => path.Contains("microsoft.graph.agentIdentityBlueprint")),
                Arg.Any<CancellationToken>(),
                Arg.Any<IEnumerable<string>?>())
            .Returns(Task.FromResult<JsonDocument?>(blueprintsDoc));

        var root = new RootCommand();
        root.AddCommand(QueryEntraCommand.CreateCommand(
            _mockLogger,
            _mockConfigService,
            _mockExecutor,
            _mockGraphApiService, _mockBlueprintService));

        // Act
        var exitCode = await root.InvokeAsync("query-entra blueprints --tenant-id tenant-id");

        // Assert
        Assert.Equal(0, exitCode);
        _mockLogger.Received().Log(
            LogLevel.Information,
            Arg.Any<EventId>(),
            Arg.Is<object>(o => o.ToString()!.Contains("Contoso Agent Blueprint")),
            Arg.Any<Exception>(),
            Arg.Any<Func<object, Exception?, string>>());
        _mockLogger.Received().Log(
            LogLevel.Information,
            Arg.Any<EventId>(),
            Arg.Is<object>(o => o.ToString()!.Contains("Total blueprints: 1")),
            Arg.Any<Exception>(),
            Arg.Any<Func<object, Exception?, string>>());
    }

}
