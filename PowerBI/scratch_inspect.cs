using System;
using Microsoft.PowerBI.Api.Models;
using System.Reflection;

namespace TestNamespace
{
    public class Program
    {
        public static void Main()
        {
            try {
                var type = typeof(PowerBIReportExportConfiguration);
                Console.WriteLine($"Properties of {type.Name}:");
                foreach (var prop in type.GetProperties())
                {
                    Console.WriteLine($"- {prop.Name} ({prop.PropertyType.Name})");
                }
            } catch (Exception ex) {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
