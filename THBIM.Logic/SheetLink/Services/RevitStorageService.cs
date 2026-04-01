using System;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using THBIM.Models;

namespace THBIM.Services
{
    /// <summary>
    /// Stores and retrieves SheetLink profile data in Revit extensible storage (DB).
    /// Data survives with the .rvt file so the last-used profile is remembered per project.
    /// </summary>
    public static class RevitStorageService
    {
        private static readonly Guid SchemaGuid = new("A1B2C3D4-E5F6-7890-ABCD-EF1234567890");
        private const string FieldName = "ProfileData";
        private const string SchemaName = "SheetLinkProfile";

        /// <summary>
        /// Save profile data into the Revit document's extensible storage.
        /// </summary>
        public static void SaveToDocument(Document doc, ProfileData profile)
        {
            if (doc == null || profile == null || doc.IsReadOnly) return;

            var txt = ProfileManager.Instance.SerializeProfile(profile);
            var schema = GetOrCreateSchema();
            var storage = GetOrCreateDataStorage(doc, schema);

            using var tx = new Transaction(doc, "SheetLink — Save Profile");
            tx.Start();
            try
            {
                var entity = new Entity(schema);
                entity.Set(FieldName, txt);
                storage.SetEntity(entity);
                tx.Commit();
            }
            catch
            {
                if (tx.HasStarted()) tx.RollBack();
            }
        }

        /// <summary>
        /// Load profile data from the Revit document's extensible storage.
        /// Returns null if no profile was previously saved.
        /// </summary>
        public static ProfileData LoadFromDocument(Document doc)
        {
            if (doc == null) return null;

            var schema = Schema.Lookup(SchemaGuid);
            if (schema == null) return null;

            var storage = new FilteredElementCollector(doc)
                .OfClass(typeof(DataStorage))
                .FirstOrDefault(ds => ds.GetEntity(schema).IsValid());

            if (storage == null) return null;

            var entity = storage.GetEntity(schema);
            if (!entity.IsValid()) return null;

            try
            {
                var txt = entity.Get<string>(FieldName);
                if (string.IsNullOrWhiteSpace(txt)) return null;
                return ProfileManager.Instance.DeserializeProfile(txt);
            }
            catch
            {
                return null;
            }
        }

        private static Schema GetOrCreateSchema()
        {
            var schema = Schema.Lookup(SchemaGuid);
            if (schema != null) return schema;

            var builder = new SchemaBuilder(SchemaGuid);
            builder.SetSchemaName(SchemaName);
            builder.SetReadAccessLevel(AccessLevel.Public);
            builder.SetWriteAccessLevel(AccessLevel.Public);
            builder.AddSimpleField(FieldName, typeof(string));
            return builder.Finish();
        }

        private static DataStorage GetOrCreateDataStorage(Document doc, Schema schema)
        {
            var existing = new FilteredElementCollector(doc)
                .OfClass(typeof(DataStorage))
                .FirstOrDefault(ds => ds.GetEntity(schema).IsValid());

            if (existing is DataStorage ds)
                return ds;

            return DataStorage.Create(doc);
        }
    }
}
