using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.Reflection;
using s2industries.ZUGFeRD;

namespace ERechnung.Core.Zugferd;

/// <summary>
/// Converts between ZUGFeRD-csharp enums and the official code values.
/// The library marks every enum member with <see cref="EnumStringValueAttribute"/> (e.g. InvoiceType.Invoice → "380"),
/// but its own conversion helpers are internal, so we read the attribute ourselves.
/// </summary>
internal static class EnumCodes
{
    private sealed record Map(Dictionary<object, string> ToCode, Dictionary<string, object> FromCode);

    private static readonly ConcurrentDictionary<Type, Map> Cache = new();

    public static string ToCode<T>(T value) where T : struct, Enum =>
        GetMap(typeof(T)).ToCode.TryGetValue(value, out var code) ? code : value.ToString();

    public static bool TryFromCode<T>(string? code, out T value) where T : struct, Enum
    {
        value = default;
        if (string.IsNullOrWhiteSpace(code)) return false;
        if (GetMap(typeof(T)).FromCode.TryGetValue(code.Trim(), out var v))
        {
            value = (T)v;
            return true;
        }
        return false;
    }

    /// <summary>Like <see cref="TryFromCode{T}"/> but throws a German, user-facing message.</summary>
    public static T FromCode<T>(string? code, string fieldLabel) where T : struct, Enum =>
        TryFromCode<T>(code, out var v)
            ? v
            : throw new InvoiceDataException($"{fieldLabel}: Der Code \"{code}\" wird nicht unterstützt.");

    private static Map GetMap(Type t) => Cache.GetOrAdd(t, static type =>
    {
        var to = new Dictionary<object, string>();
        var from = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        foreach (var f in type.GetFields(BindingFlags.Public | BindingFlags.Static))
        {
            var value = f.GetValue(null)!;
            var attr = f.GetCustomAttributesData()
                .FirstOrDefault(a => a.AttributeType.Name == nameof(EnumStringValueAttribute));
            var code = attr?.ConstructorArguments.Count > 0 ? attr.ConstructorArguments[0].Value as string : null;
            code ??= f.Name;
            to.TryAdd(value, code);
            from.TryAdd(code, value);
            // Alternative spellings, e.g. QuantityCodes.C62 also accepts "PCE"
            if (attr?.ConstructorArguments.Count > 1 &&
                attr.ConstructorArguments[1].Value is ReadOnlyCollection<CustomAttributeTypedArgument> alts)
            {
                foreach (var alt in alts)
                    if (alt.Value is string s) from.TryAdd(s, value);
            }
        }
        return new Map(to, from);
    });
}
