using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace ElectricalToolSuite.MECoordination
{
    class MechanicalElectricalEquipmentCoordinator
    {
        private readonly DocumentHelper _document;
        private readonly Workset _electricalWorkset;

        public MechanicalElectricalEquipmentCoordinator(DocumentHelper document, Workset electricalWorkset)
        {
            _document = document;
            _electricalWorkset = electricalWorkset;
        }

        public void Coordinate(ICollection<ElementId> mechanicalInstances, ICollection<FamilySymbol> electricalSymbols)
        {
            foreach (var electricalSymbol in electricalSymbols)
            {
                foreach (var mechanicalId in mechanicalInstances)
                {
                    var mechanicalInstance = _document.Document.GetElement(mechanicalId) as FamilyInstance;
                    CreateElectricalInstance(mechanicalInstance, electricalSymbol);
                }
            }
        }

        private void CreateElectricalInstance(FamilyInstance mechanicalInstance, FamilySymbol electricalSymbol)
        {
            var host = mechanicalInstance.Host;
            var xyz = GetInstanceXYZ(mechanicalInstance);

            var newInstance = _document.CreateNonStructuralInstance(xyz, electricalSymbol, host);            

            CopyMatchingParameters(mechanicalInstance.GetOrderedParameters(), newInstance);
            SetWorkset(newInstance, _electricalWorkset);
        }

        private XYZ GetInstanceXYZ(FamilyInstance instance)
        {
            return ((LocationPoint) instance.Location).Point;
        }

        private static void SetWorkset(FamilyInstance newInstance, Workset workset)
        {
            if (workset == null)
                return;

            var worksetParameter = newInstance.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
            worksetParameter.Set(workset.Id.IntegerValue);
        }

        private static void CopyMatchingParameters(IEnumerable<Parameter> targetParameters, FamilyInstance newInstance)
        {
            foreach (var targetParameter in targetParameters)
            {
                var newInstanceParameter = newInstance.LookupParameter(targetParameter.Definition.Name);
                if (newInstanceParameter != null && !newInstanceParameter.IsReadOnly)
                {
                    CopyParameterValue(targetParameter, newInstanceParameter);
                }
            }
        }

        private static void CopyParameterValue(Parameter targetParameter, Parameter newInstanceParameter)
        {
            switch (targetParameter.StorageType)
            {
                case StorageType.Double:
                    newInstanceParameter.Set(targetParameter.AsDouble());
                    break;
                case StorageType.ElementId:
                    newInstanceParameter.Set(targetParameter.AsElementId());
                    break;
                case StorageType.Integer:
                    newInstanceParameter.Set(targetParameter.AsInteger());
                    break;
                case StorageType.String:
                    newInstanceParameter.Set(targetParameter.AsString());
                    break;
                default:
                    return;
            }
        }
    }
}
