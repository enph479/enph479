using System;

namespace ElectricalToolSuite.MECoordination
{
    internal static class NullableExtensions
    {
        public static T ValueOr<T>(this Nullable<T> nullable, T defaultValue) where T : struct
        {
            if (!nullable.HasValue)
                return defaultValue;

            return nullable.Value;
        }
    }
}
