using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace QicBoqMapper
{
    public static class CategoryMappingService
    {
        private static readonly Dictionary<string, BuiltInCategory> Mappings = new Dictionary<string, BuiltInCategory>
        {
            // Architecture
            { "Walls", BuiltInCategory.OST_Walls },
            { "Floors", BuiltInCategory.OST_Floors },
            { "Ceilings", BuiltInCategory.OST_Ceilings },
            { "Roofs", BuiltInCategory.OST_Roofs },
            { "Doors", BuiltInCategory.OST_Doors },
            { "Windows", BuiltInCategory.OST_Windows },
            { "Rooms", BuiltInCategory.OST_Rooms },
            { "Areas", BuiltInCategory.OST_Areas },
            { "Furniture", BuiltInCategory.OST_Furniture },
            { "Furniture Systems", BuiltInCategory.OST_FurnitureSystems },
            { "Casework", BuiltInCategory.OST_Casework },
            { "Generic Models", BuiltInCategory.OST_GenericModel },
            { "Specialty Equipment", BuiltInCategory.OST_SpecialityEquipment },

            // Structure
            { "Structural Columns", BuiltInCategory.OST_StructuralColumns },
            { "Structural Framing", BuiltInCategory.OST_StructuralFraming },
            { "Structural Foundations", BuiltInCategory.OST_StructuralFoundation },
            { "Structural Connections", BuiltInCategory.OST_StructConnections },
            { "Rebar", BuiltInCategory.OST_Rebar },

            // Mechanical
            { "Mechanical Equipment", BuiltInCategory.OST_MechanicalEquipment },
            { "Ducts", BuiltInCategory.OST_DuctCurves },
            { "Flex Ducts", BuiltInCategory.OST_FlexDuctCurves },
            { "Air Terminals", BuiltInCategory.OST_DuctTerminal },
            { "Duct Fittings", BuiltInCategory.OST_DuctFitting },
            { "Duct Accessories", BuiltInCategory.OST_DuctAccessory },

            // Plumbing
            { "Pipes", BuiltInCategory.OST_PipeCurves },
            { "Flex Pipes", BuiltInCategory.OST_FlexPipeCurves },
            { "Pipe Fittings", BuiltInCategory.OST_PipeFitting },
            { "Pipe Accessories", BuiltInCategory.OST_PipeAccessory },
            { "Plumbing Fixtures", BuiltInCategory.OST_PlumbingFixtures },
            { "Sprinklers", BuiltInCategory.OST_Sprinklers },

            // Electrical
            { "Electrical Equipment", BuiltInCategory.OST_ElectricalEquipment },
            { "Electrical Fixtures", BuiltInCategory.OST_ElectricalFixtures },
            { "Lighting Fixtures", BuiltInCategory.OST_LightingFixtures },
            { "Lighting Devices", BuiltInCategory.OST_LightingDevices },
            { "Data Devices", BuiltInCategory.OST_DataDevices },
            { "Communication Devices", BuiltInCategory.OST_CommunicationDevices },
            { "Fire Alarm Devices", BuiltInCategory.OST_FireAlarmDevices },
            { "Security Devices", BuiltInCategory.OST_SecurityDevices },
            { "Cable Trays", BuiltInCategory.OST_CableTray },
            { "Cable Tray Fittings", BuiltInCategory.OST_CableTrayFitting },
            { "Conduits", BuiltInCategory.OST_Conduit },
            { "Conduit Fittings", BuiltInCategory.OST_ConduitFitting },

            // Site
            { "Parking", BuiltInCategory.OST_Parking },
            { "Planting", BuiltInCategory.OST_Planting },
            { "Topography", BuiltInCategory.OST_Topography },
            { "Site Components", BuiltInCategory.OST_Site }
        };

        public static IEnumerable<string> GetCategoryNames() => Mappings.Keys;

        public static BuiltInCategory GetBuiltInCategory(string name)
        {
            if (Mappings.TryGetValue(name, out var cat))
                return cat;
            throw new ArgumentException($"Category '{name}' is not supported.");
        }
    }
}
