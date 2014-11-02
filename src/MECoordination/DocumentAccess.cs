using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;

namespace ElectricalToolSuite.MECoordination
{
    public class DocumentAccess
    {
        public Document Document { get; private set; }

        public DocumentAccess(Document document)
        {
            Document = document;
        }

        public IEnumerable<FamilySymbol> FamilySymbols
        {
            get
            {
                return new FilteredElementCollector(Document)
                    .WherePasses(new ElementIsCurveDrivenFilter(true))
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>();
            }
        }

        public List<Workset> UserWorksets
        {
            get
            {
                return new FilteredWorksetCollector(Document)
                .OfKind(WorksetKind.UserWorkset)
                .ToList();
            }
        }

        public WorksetId ActiveWorksetId
        {
            get { return Document.GetWorksetTable().GetActiveWorksetId(); }
        }

        public FamilyInstance CreateNonStructuralInstance(XYZ location, FamilySymbol symbol, Element host)
        {
            if (host != null)
            {
                return Document.Create.NewFamilyInstance(location, symbol,
                            host, StructuralType.NonStructural);
            }

            return Document.Create.NewFamilyInstance(location, symbol,
                        StructuralType.NonStructural);
        }

        public IEnumerable<ElementId> GetInstancesOfFamilySymbols(IEnumerable<FamilySymbol> symbols)
        {
            var filters = symbols.Select(symbol => new FamilyInstanceFilter(Document, symbol.Id))
                .Cast<ElementFilter>()
                .ToList();
            var unionFilter = new LogicalOrFilter(filters);
            return new FilteredElementCollector(Document)
                .WherePasses(unionFilter)
                .ToElementIds();
        }
    }
}
