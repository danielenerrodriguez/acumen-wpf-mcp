using System.Windows.Automation;

namespace WpfMcp;

/// <summary>
/// Thread-safe element reference cache shared across tools, proxy server, and CLI mode.
/// Stores AutomationElement references by generated keys (e.g., "e1", "e2").
/// Uses LRU eviction when the cache exceeds <see cref="MaxCapacity"/> to prevent
/// unbounded memory growth from stale COM references.
/// </summary>
public class ElementCache
{
    /// <summary>Maximum number of cached elements before LRU eviction kicks in.</summary>
    public const int MaxCapacity = 500;

    private readonly Dictionary<string, LinkedListNode<CacheEntry>> _map = new();
    private readonly LinkedList<CacheEntry> _order = new();
    private readonly object _lock = new();
    private int _counter;

    private record CacheEntry(string Key, AutomationElement Element);

    public string Add(AutomationElement element)
    {
        var key = $"e{Interlocked.Increment(ref _counter)}";
        lock (_lock)
        {
            // If key already exists (shouldn't with incrementing counter, but be safe)
            if (_map.TryGetValue(key, out var existing))
            {
                _order.Remove(existing);
                _map.Remove(key);
            }

            // Evict oldest entries if at capacity
            while (_map.Count >= MaxCapacity && _order.Count > 0)
            {
                var oldest = _order.Last!;
                _order.RemoveLast();
                _map.Remove(oldest.Value.Key);
            }

            var node = _order.AddFirst(new CacheEntry(key, element));
            _map[key] = node;
        }
        return key;
    }

    public bool TryGet(string key, out AutomationElement? element)
    {
        lock (_lock)
        {
            if (_map.TryGetValue(key, out var node))
            {
                // Move to front (most recently used)
                _order.Remove(node);
                _order.AddFirst(node);
                element = node.Value.Element;
                return true;
            }
            element = null;
            return false;
        }
    }

    public int Count
    {
        get { lock (_lock) { return _map.Count; } }
    }

    public void Clear()
    {
        lock (_lock)
        {
            _map.Clear();
            _order.Clear();
        }
    }
}
