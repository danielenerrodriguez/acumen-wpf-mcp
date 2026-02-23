using System.Windows.Automation;

namespace WpfMcp;

/// <summary>
/// Thread-safe element reference cache shared across tools, proxy server, and CLI mode.
/// Stores AutomationElement references by generated keys (e.g., "e1", "e2").
/// </summary>
public class ElementCache
{
    private readonly Dictionary<string, AutomationElement> _cache = new();
    private int _counter;

    public string Add(AutomationElement element)
    {
        var key = $"e{Interlocked.Increment(ref _counter)}";
        lock (_cache)
        {
            _cache[key] = element;
        }
        return key;
    }

    public bool TryGet(string key, out AutomationElement? element)
    {
        lock (_cache)
        {
            return _cache.TryGetValue(key, out element);
        }
    }

    public int Count
    {
        get { lock (_cache) { return _cache.Count; } }
    }

    public void Clear()
    {
        lock (_cache)
        {
            _cache.Clear();
        }
    }
}
