using System.Data.OleDb;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace MS.Access.MCP.Interop
{
    public class AccessInteropService : IDisposable
    {
        private OleDbConnection? _oleDbConnection;
        private string? _currentDatabasePath;
        private bool _disposed = false;

        #region 1. Connection Management

        public void Connect(string databasePath)
        {
            if (!File.Exists(databasePath))
                throw new FileNotFoundException($"Database file not found: {databasePath}");

            _currentDatabasePath = databasePath;
            
            // Create OleDb connection for direct data access
            var connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};";
            _oleDbConnection = new OleDbConnection(connectionString);
            _oleDbConnection.Open();
        }

        public void Disconnect()
        {
            _oleDbConnection?.Close();
            _oleDbConnection?.Dispose();
            _oleDbConnection = null;
            _currentDatabasePath = null;
        }

        public bool IsConnected => _oleDbConnection?.State == System.Data.ConnectionState.Open;

        #endregion

        #region 2. Data Access Object Models

        public List<TableInfo> GetTables()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var tables = new List<TableInfo>();
            
            // Use OleDb to get table information
            var schema = _oleDbConnection!.GetSchema("Tables");
            
            foreach (System.Data.DataRow row in schema.Rows)
            {
                var tableName = row["TABLE_NAME"].ToString();
                if (!string.IsNullOrEmpty(tableName) && !tableName.StartsWith("~"))
                {
                    var fields = GetTableFields(tableName);
                    tables.Add(new TableInfo
                    {
                        Name = tableName,
                        Fields = fields,
                        RecordCount = GetTableRecordCount(tableName)
                    });
                }
            }

            return tables;
        }

        public List<QueryInfo> GetQueries()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var queries = new List<QueryInfo>();
            
            // Use OleDb to get query information
            var schema = _oleDbConnection!.GetSchema("Views");
            
            foreach (System.Data.DataRow row in schema.Rows)
            {
                var queryName = row["TABLE_NAME"].ToString();
                if (!string.IsNullOrEmpty(queryName))
                {
                    queries.Add(new QueryInfo
                    {
                        Name = queryName,
                        SQL = "", // SQL not available through schema
                        Type = "Query"
                    });
                }
            }

            return queries;
        }

        public List<RelationshipInfo> GetRelationships()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var relationships = new List<RelationshipInfo>();
            
            // Use OleDb to get relationship information
            var schema = _oleDbConnection!.GetSchema("ForeignKeys");
            
            foreach (System.Data.DataRow row in schema.Rows)
            {
                relationships.Add(new RelationshipInfo
                {
                    Name = row["FK_NAME"]?.ToString() ?? "",
                    Table = row["TABLE_NAME"]?.ToString() ?? "",
                    ForeignTable = row["REFERENCED_TABLE_NAME"]?.ToString() ?? "",
                    Attributes = ""
                });
            }

            return relationships;
        }

        public void CreateTable(string tableName, List<FieldInfo> fields)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var fieldDefinitions = new List<string>();
            foreach (var field in fields)
            {
                var fieldDef = $"[{field.Name}] {field.Type}";
                if (field.Size > 0 && field.Type.ToLower() == "text")
                    fieldDef += $"({field.Size})";
                if (field.Required)
                    fieldDef += " NOT NULL";
                fieldDefinitions.Add(fieldDef);
            }

            var createSql = $"CREATE TABLE [{tableName}] ({string.Join(", ", fieldDefinitions)})";
            var command = new OleDbCommand(createSql, _oleDbConnection);
            command.ExecuteNonQuery();
        }

        public void DeleteTable(string tableName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            var command = new OleDbCommand($"DROP TABLE [{tableName}]", _oleDbConnection);
            command.ExecuteNonQuery();
        }

        #endregion

        #region 3. COM Automation (Simplified)

        public void LaunchAccess()
        {
            // This would require full COM interop - simplified for now
            Console.WriteLine("Access launch functionality requires full COM interop");
        }

        public void CloseAccess()
        {
            // This would require full COM interop - simplified for now
            Console.WriteLine("Access close functionality requires full COM interop");
        }

        public List<FormInfo> GetForms()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var forms = new List<FormInfo>();
            
            // Try to get forms from system tables
            try
            {
                var command = new OleDbCommand("SELECT Name FROM MSysObjects WHERE Type = -32768", _oleDbConnection);
                using var reader = command.ExecuteReader();
                
                while (reader.Read())
                {
                    forms.Add(new FormInfo
                    {
                        Name = reader["Name"]?.ToString() ?? "",
                        FullName = reader["Name"]?.ToString() ?? "",
                        Type = "Form"
                    });
                }
            }
            catch
            {
                // MSysObjects might not be accessible
            }

            return forms;
        }

        public List<ReportInfo> GetReports()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var reports = new List<ReportInfo>();
            
            // Try to get reports from system tables
            try
            {
                var command = new OleDbCommand("SELECT Name FROM MSysObjects WHERE Type = -32764", _oleDbConnection);
                using var reader = command.ExecuteReader();
                
                while (reader.Read())
                {
                    reports.Add(new ReportInfo
                    {
                        Name = reader["Name"]?.ToString() ?? "",
                        FullName = reader["Name"]?.ToString() ?? "",
                        Type = "Report"
                    });
                }
            }
            catch
            {
                // MSysObjects might not be accessible
            }

            return reports;
        }

        public List<MacroInfo> GetMacros()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var macros = new List<MacroInfo>();
            
            // Try to get macros from system tables
            try
            {
                var command = new OleDbCommand("SELECT Name FROM MSysObjects WHERE Type = -32766", _oleDbConnection);
                using var reader = command.ExecuteReader();
                
                while (reader.Read())
                {
                    macros.Add(new MacroInfo
                    {
                        Name = reader["Name"]?.ToString() ?? "",
                        FullName = reader["Name"]?.ToString() ?? "",
                        Type = "Macro"
                    });
                }
            }
            catch
            {
                // MSysObjects might not be accessible
            }

            return macros;
        }

        public List<ModuleInfo> GetModules()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var modules = new List<ModuleInfo>();
            
            // Try to get modules from system tables
            try
            {
                var command = new OleDbCommand("SELECT Name FROM MSysObjects WHERE Type = -32761", _oleDbConnection);
                using var reader = command.ExecuteReader();
                
                while (reader.Read())
                {
                    modules.Add(new ModuleInfo
                    {
                        Name = reader["Name"]?.ToString() ?? "",
                        FullName = reader["Name"]?.ToString() ?? "",
                        Type = "Module"
                    });
                }
            }
            catch
            {
                // MSysObjects might not be accessible
            }

            return modules;
        }

        public void OpenForm(string formName)
        {
            // This would require full COM interop - simplified for now
            Console.WriteLine($"Form {formName} open functionality requires full COM interop");
        }

        public void CloseForm(string formName)
        {
            // This would require full COM interop - simplified for now
            Console.WriteLine($"Form {formName} close functionality requires full COM interop");
        }

        #endregion

        #region 4. VBA Extensibility (Simplified)

        public List<VBAProjectInfo> GetVBAProjects()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var projects = new List<VBAProjectInfo>();
            
            // Simplified VBA project discovery
            try
            {
                var command = new OleDbCommand("SELECT Name FROM MSysObjects WHERE Type = -32761", _oleDbConnection);
                using var reader = command.ExecuteReader();
                
                var modules = new List<VBAModuleInfo>();
                while (reader.Read())
                {
                    modules.Add(new VBAModuleInfo
                    {
                        Name = reader["Name"]?.ToString() ?? "",
                        Type = "Module",
                        HasCode = true
                    });
                }

                projects.Add(new VBAProjectInfo
                {
                    Name = "CurrentProject",
                    Description = "Current Access Project",
                    Modules = modules
                });
            }
            catch
            {
                // MSysObjects might not be accessible
            }

            return projects;
        }

        public string GetVBACode(string projectName, string moduleName)
        {
            // This would require full COM interop - simplified for now
            return $"// VBA code for {moduleName} would be retrieved here";
        }

        public void SetVBACode(string projectName, string moduleName, string code)
        {
            // This would require full COM interop - simplified for now
            Console.WriteLine($"VBA code for {moduleName} would be set here");
        }

        public void AddVBAProcedure(string projectName, string moduleName, string procedureName, string code)
        {
            // This would require full COM interop - simplified for now
            Console.WriteLine($"VBA procedure {procedureName} would be added to {moduleName}");
        }

        public void CompileVBA()
        {
            // This would require full COM interop - simplified for now
            Console.WriteLine("VBA compilation would be performed here");
        }

        #endregion

        #region 5. System Table Metadata Access

        public List<SystemTableInfo> GetSystemTables()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var systemTables = new List<SystemTableInfo>();
            var schema = _oleDbConnection!.GetSchema("Tables");
            
            foreach (System.Data.DataRow row in schema.Rows)
            {
                var tableName = row["TABLE_NAME"].ToString();
                if (!string.IsNullOrEmpty(tableName) && (tableName.StartsWith("~") || tableName.StartsWith("MSys")))
                {
                    systemTables.Add(new SystemTableInfo
                    {
                        Name = tableName,
                        DateCreated = DateTime.Now, // Not available through OleDb
                        LastUpdated = DateTime.Now, // Not available through OleDb
                        RecordCount = GetTableRecordCount(tableName)
                    });
                }
            }

            return systemTables;
        }

        public List<MetadataInfo> GetObjectMetadata()
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var metadata = new List<MetadataInfo>();
            
            try
            {
                // Query MSysObjects table for object metadata
                var command = new OleDbCommand("SELECT * FROM MSysObjects", _oleDbConnection);
                using var reader = command.ExecuteReader();
                
                while (reader.Read())
                {
                    metadata.Add(new MetadataInfo
                    {
                        Name = reader["Name"]?.ToString() ?? "",
                        Type = reader["Type"]?.ToString() ?? "",
                        Flags = reader["Flags"]?.ToString() ?? "",
                        DateCreated = reader["DateCreate"]?.ToString() ?? "",
                        DateModified = reader["DateUpdate"]?.ToString() ?? ""
                    });
                }
            }
            catch
            {
                // MSysObjects might not be accessible, return empty list
            }

            return metadata;
        }

        #endregion

        #region 6. Form & Control Discovery & Editing APIs (Simplified)

        public bool FormExists(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            
            try
            {
                var command = new OleDbCommand("SELECT COUNT(*) FROM MSysObjects WHERE Name = ? AND Type = -32768", _oleDbConnection);
                command.Parameters.AddWithValue("@Name", formName);
                var count = Convert.ToInt32(command.ExecuteScalar());
                return count > 0;
            }
            catch
            {
                return false;
            }
        }

        public List<ControlInfo> GetFormControls(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var controls = new List<ControlInfo>();
            
            // Simplified control discovery - would require full COM interop for actual control enumeration
            // For now, return a placeholder control
            controls.Add(new ControlInfo
            {
                Name = "PlaceholderControl",
                Type = "TextBox",
                Left = 100,
                Top = 100,
                Width = 200,
                Height = 25,
                Visible = true,
                Enabled = true
            });

            return controls;
        }

        public ControlProperties GetControlProperties(string formName, string controlName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            // Simplified control properties - would require full COM interop for actual properties
            return new ControlProperties
            {
                Name = controlName,
                Type = "TextBox",
                Left = 100,
                Top = 100,
                Width = 200,
                Height = 25,
                Visible = true,
                Enabled = true,
                BackColor = 16777215, // White
                ForeColor = 0, // Black
                FontName = "Arial",
                FontSize = 10,
                FontBold = false,
                FontItalic = false
            };
        }

        public void SetControlProperty(string formName, string controlName, string propertyName, object value)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            // This would require full COM interop - simplified for now
            Console.WriteLine($"Property {propertyName} of control {controlName} would be set to {value}");
        }

        #endregion

        #region 7. Persistence & Versioning

        public string ExportFormToText(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var formData = new
            {
                Name = formName,
                ExportedAt = DateTime.UtcNow,
                Controls = GetFormControls(formName),
                VBA = GetVBACode("CurrentProject", formName)
            };

            return JsonSerializer.Serialize(formData, new JsonSerializerOptions { WriteIndented = true });
        }

        public void ImportFormFromText(string formData)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var formInfo = JsonSerializer.Deserialize<FormExportData>(formData);
            if (formInfo == null) throw new ArgumentException("Invalid form data");

            // Simplified form import - would require full COM interop for actual form creation
            Console.WriteLine($"Form {formInfo.Name} would be imported here");
        }

        public void DeleteForm(string formName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            
            // This would require full COM interop - simplified for now
            Console.WriteLine($"Form {formName} would be deleted here");
        }

        public string ExportReportToText(string reportName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var reportData = new
            {
                Name = reportName,
                ExportedAt = DateTime.UtcNow,
                Controls = GetFormControls(reportName) // Reuse form controls for reports
            };

            return JsonSerializer.Serialize(reportData, new JsonSerializerOptions { WriteIndented = true });
        }

        public void ImportReportFromText(string reportData)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");

            var reportInfo = JsonSerializer.Deserialize<ReportExportData>(reportData);
            if (reportInfo == null) throw new ArgumentException("Invalid report data");

            // Simplified report import - would require full COM interop for actual report creation
            Console.WriteLine($"Report {reportInfo.Name} would be imported here");
        }

        public void DeleteReport(string reportName)
        {
            if (!IsConnected) throw new InvalidOperationException("Not connected to database");
            
            // This would require full COM interop - simplified for now
            Console.WriteLine($"Report {reportName} would be deleted here");
        }

        #endregion

        #region Helper Methods

        private List<FieldInfo> GetTableFields(string tableName)
        {
            var fields = new List<FieldInfo>();
            
            try
            {
                var schema = _oleDbConnection!.GetSchema("Columns", new string[] { null!, null!, tableName });
                
                foreach (System.Data.DataRow row in schema.Rows)
                {
                    fields.Add(new FieldInfo
                    {
                        Name = row["COLUMN_NAME"]?.ToString() ?? "",
                        Type = row["DATA_TYPE"]?.ToString() ?? "",
                        Size = Convert.ToInt32(row["CHARACTER_MAXIMUM_LENGTH"] ?? 0),
                        Required = row["IS_NULLABLE"]?.ToString() == "NO",
                        AllowZeroLength = true // Default value
                    });
                }
            }
            catch
            {
                // Return empty list if table doesn't exist or can't be accessed
            }

            return fields;
        }

        private long GetTableRecordCount(string tableName)
        {
            try
            {
                var command = new OleDbCommand($"SELECT COUNT(*) FROM [{tableName}]", _oleDbConnection);
                return Convert.ToInt64(command.ExecuteScalar());
            }
            catch
            {
                return 0;
            }
        }

        #endregion

        public void Dispose()
        {
            if (!_disposed)
            {
                Disconnect();
                _disposed = true;
            }
        }
    }

    #region Data Models

    public class TableInfo
    {
        public string Name { get; set; } = "";
        public List<FieldInfo> Fields { get; set; } = new();
        public long RecordCount { get; set; }
    }

    public class FieldInfo
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public int Size { get; set; }
        public bool Required { get; set; }
        public bool AllowZeroLength { get; set; }
    }

    public class QueryInfo
    {
        public string Name { get; set; } = "";
        public string SQL { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class RelationshipInfo
    {
        public string Name { get; set; } = "";
        public string Table { get; set; } = "";
        public string ForeignTable { get; set; } = "";
        public string Attributes { get; set; } = "";
    }

    public class FormInfo
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class ReportInfo
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class MacroInfo
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class ModuleInfo
    {
        public string Name { get; set; } = "";
        public string FullName { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class VBAProjectInfo
    {
        public string Name { get; set; } = "";
        public string Description { get; set; } = "";
        public List<VBAModuleInfo> Modules { get; set; } = new();
    }

    public class VBAModuleInfo
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public bool HasCode { get; set; }
    }

    public class SystemTableInfo
    {
        public string Name { get; set; } = "";
        public DateTime DateCreated { get; set; }
        public DateTime LastUpdated { get; set; }
        public long RecordCount { get; set; }
    }

    public class MetadataInfo
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public string Flags { get; set; } = "";
        public string DateCreated { get; set; } = "";
        public string DateModified { get; set; } = "";
    }

    public class ControlInfo
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public int Left { get; set; }
        public int Top { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public bool Visible { get; set; }
        public bool Enabled { get; set; }
    }

    public class ControlProperties
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public int Left { get; set; }
        public int Top { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public bool Visible { get; set; }
        public bool Enabled { get; set; }
        public int BackColor { get; set; }
        public int ForeColor { get; set; }
        public string FontName { get; set; } = "";
        public int FontSize { get; set; }
        public bool FontBold { get; set; }
        public bool FontItalic { get; set; }
    }

    public class FormExportData
    {
        public string Name { get; set; } = "";
        public DateTime ExportedAt { get; set; }
        public List<ControlInfo> Controls { get; set; } = new();
        public string VBA { get; set; } = "";
    }

    public class ReportExportData
    {
        public string Name { get; set; } = "";
        public DateTime ExportedAt { get; set; }
        public List<ControlInfo> Controls { get; set; } = new();
    }

    #endregion
} 