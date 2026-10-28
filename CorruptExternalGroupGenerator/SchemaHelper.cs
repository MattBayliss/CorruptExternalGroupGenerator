using Microsoft.Graph.Models.ExternalConnectors;
using System;
using System.Collections.Generic;
using System.Text;

namespace CorruptExternalGroupGenerator
{
    internal static class SchemaHelper
    {
        public static Schema BasicSchema => new Schema
        {
            BaseType = "microsoft.graph.externalItem",
            Properties = 
                [
                    new Property {
                        Name = "title",
                        Type = PropertyType.String,
                        IsSearchable = true,
                        IsRetrievable = true,
                        IsQueryable = true,
                        IsRefinable = false,
                        Labels = [Label.Title]
                    }
                ]
        };
    }
}
