using System;

namespace StyleSnooper
{
    public sealed record StyleModel(string DisplayName, object? ResourceKey, Type ElementType);
}
