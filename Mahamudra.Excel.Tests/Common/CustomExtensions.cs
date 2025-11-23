using System;

namespace Mahamudra.Excel.Tests.Common;

public static class CustomExtensions
{
    public static int GetRandomInteger(int maxValue = 10000)
    {
        var random = new Random();
        return random.Next(maxValue);
    }

    public static decimal GetRandomDecimal()
    {
        var random = new Random();
        return (decimal)random.NextDouble() * 100;
    }

    public static string GetRandomString(int length = 25)
    {
        var random = new Random();
        const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
        return new string(Enumerable.Repeat(chars, length)
            .Select(s => s[random.Next(s.Length)]).ToArray());
    }
}