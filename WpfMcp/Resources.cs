using System.ComponentModel;
using ModelContextProtocol.Server;

namespace WpfMcp;

/// <summary>
/// MCP resources that serve knowledge base content to AI agents.
/// Knowledge bases provide navigation context (automation IDs, keytips,
/// workflows, etc.) for specific applications.
/// </summary>
[McpServerResourceType]
public static class KnowledgeResources
{
    [McpServerResource(UriTemplate = "knowledge://{productName}", Name = "Application Knowledge Base")]
    [Description("Full knowledge base YAML for navigating a specific application via WPF MCP tools. Contains automation IDs, keytips, workflows, and navigation tips.")]
    public static string GetKnowledgeBase(string productName)
    {
        var knowledgeBases = WpfTools.GetKnowledgeBases();
        var kb = knowledgeBases.FirstOrDefault(k =>
            k.ProductName.Equals(productName, StringComparison.OrdinalIgnoreCase));

        if (kb == null)
        {
            var available = knowledgeBases.Select(k => k.ProductName).ToList();
            return available.Count > 0
                ? $"Knowledge base '{productName}' not found. Available: {string.Join(", ", available)}"
                : "No knowledge bases loaded. Place _knowledge.yaml files in the macros/ product subfolders.";
        }

        return kb.FullContent;
    }
}
