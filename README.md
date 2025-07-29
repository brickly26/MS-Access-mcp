# Microsoft Access MCP Server

A comprehensive Model Context Protocol (MCP) server that provides full Microsoft Access automation capabilities for Claude Desktop and other MCP clients.

## Overview

This MCP server enables AI assistants to interact with Microsoft Access databases through a complete set of tools that implement all seven required capabilities:

1. **Connection Management** - Establish and manage database connections
2. **Data Access Object Models** - Discover and manipulate tables, queries, and relationships
3. **COM Automation** - Launch Access and manage forms, reports, macros, and modules
4. **VBA Extensibility** - Read, write, and compile VBA code
5. **System Table Metadata Access** - Access hidden system tables and metadata
6. **Form & Control Discovery & Editing APIs** - Discover and modify form controls
7. **Persistence & Versioning** - Export/import database objects for version control

## Architecture

The solution consists of two main components:

1. **MS.Access.MCP.Interop** - A .NET 8.0 interop library that handles COM interactions with Microsoft Access
2. **MS.Access.MCP.Official** - An MCP server using the official ModelContextProtocol package that exposes Access functionality as MCP tools

## Prerequisites

- Windows operating system
- Microsoft Access installed (for COM interop)
- .NET 8.0 runtime
- Access database file (.accdb or .mdb)

## Installation

### Building from Source

1. Clone or download this repository
2. Build the solution:
   ```bash
   dotnet build
   ```
3. Publish the official server:
   ```bash
   dotnet publish MS.Access.MCP.Official/MS.Access.MCP.Official.csproj -c Release -o ./mcp-server-official --self-contained
   ```

## Configuration

### Claude Desktop Configuration

Add the following to your Claude Desktop configuration file (`%APPDATA%\Claude\claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "access-mcp-server": {
      "command": "C:\\Users\\brickly\\Desktop\\MS-Access-MCP\\mcp-server-official\\MS.Access.MCP.Official.exe",
      "args": []
    }
  }
}
```

**Note**: Use the absolute path to your MCP server executable and double backslashes.

## Available Tools

The MCP server provides comprehensive tools across all seven capability areas:

### 1. Connection Management
- **connect_access**: Connect to an Access database (requires `database_path` parameter)
- **disconnect_access**: Disconnect from the current Access database
- **is_connected**: Check if connected to an Access database

### 2. Data Access Object Models
- **get_tables**: Get all tables from the connected database
- **get_queries**: Get all queries from the connected database
- **get_relationships**: Get all relationships from the connected database
- **create_table**: Create a new table in the database
- **delete_table**: Delete a table from the database

### 3. COM Automation
- **launch_access**: Launch Microsoft Access application
- **close_access**: Close Microsoft Access application
- **get_forms**: Get all forms in the database
- **get_reports**: Get all reports in the database
- **get_macros**: Get all macros in the database
- **get_modules**: Get all modules in the database
- **open_form**: Open a form in Access
- **close_form**: Close a form in Access

### 4. VBA Extensibility
- **get_vba_projects**: Get all VBA projects in the database
- **get_vba_code**: Get VBA code from a module
- **set_vba_code**: Set VBA code in a module
- **add_vba_procedure**: Add a VBA procedure to a module
- **compile_vba**: Compile VBA code

### 5. System Table Metadata Access
- **get_system_tables**: Get system tables from the database
- **get_object_metadata**: Get object metadata from system tables

### 6. Form & Control Discovery & Editing APIs
- **form_exists**: Check if a form exists
- **get_form_controls**: Get all controls in a form
- **get_control_properties**: Get properties of a control
- **set_control_property**: Set a property of a control

### 7. Persistence & Versioning
- **export_form_to_text**: Export a form to text representation
- **import_form_from_text**: Import a form from text representation
- **delete_form**: Delete a form from the database
- **export_report_to_text**: Export a report to text representation
- **import_report_from_text**: Import a report from text representation
- **delete_report**: Delete a report from the database

## Usage Examples

### Basic Connection and Discovery
```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "tools/call",
  "params": {
    "name": "connect_access",
    "arguments": {
      "database_path": "C:\\path\\to\\database.accdb"
    }
  }
}
```

### Table Operations
```json
{
  "jsonrpc": "2.0",
  "id": 2,
  "method": "tools/call",
  "params": {
    "name": "get_tables",
    "arguments": {}
  }
}
```

### VBA Code Management
```json
{
  "jsonrpc": "2.0",
  "id": 3,
  "method": "tools/call",
  "params": {
    "name": "set_vba_code",
    "arguments": {
      "project_name": "CurrentProject",
      "module_name": "Module1",
      "code": "Sub TestProcedure()\n    MsgBox \"Hello World\"\nEnd Sub"
    }
  }
}
```

### Form Export/Import
```json
{
  "jsonrpc": "2.0",
  "id": 4,
  "method": "tools/call",
  "params": {
    "name": "export_form_to_text",
    "arguments": {
      "form_name": "MyForm"
    }
  }
}
```

## Testing

### Running the Integration Test

The project includes a comprehensive integration test that validates all seven capabilities:

```bash
dotnet run --project IntegrationTest.csproj
```

The integration test:
1. Opens a sample database
2. Tests connection management (connect, disconnect, status check)
3. Discovers and lists tables, queries, and relationships
4. Launches Access, loads forms, then closes it
5. Inserts and reads back VBA procedures
6. Queries system table metadata
7. Reads, updates, and verifies control properties
8. Exports forms to text, deletes them, re-imports them, and confirms existence

### Manual Testing

To test the server manually:

```bash
echo {"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"1.0"}}} | dotnet run --project MS.Access.MCP.Official/MS.Access.MCP.Official.csproj
```

This should return a proper MCP initialize response with server information.

## Development

The server is built using the official `ModelContextProtocol` package (version 0.3.0-preview.3) and implements a JSON-RPC message handler that:

1. Reads JSON-RPC messages from stdin
2. Processes MCP protocol methods (initialize, tools/list, tools/call)
3. Executes Access database operations through the Interop library
4. Returns results as JSON-RPC responses on stdout

### Key Features

- **Full COM Interop**: Direct access to Microsoft Access COM objects
- **Comprehensive Coverage**: All seven required capabilities implemented
- **Error Handling**: Robust error handling and resource cleanup
- **Type Safety**: Strongly typed C# interfaces for all operations
- **Extensibility**: Easy to add new tools and capabilities

### Data Models

The Interop library includes comprehensive data models for:
- Tables, fields, and relationships
- Forms, reports, macros, and modules
- VBA projects and code
- System tables and metadata
- Form controls and properties
- Export/import data structures

## Troubleshooting

### Common Issues

1. **Access not installed**: Ensure Microsoft Access is installed on the system
2. **Database not found**: Verify the database path is correct and accessible
3. **Permission errors**: Ensure the application has permission to access the database file
4. **COM errors**: May indicate Access is not properly installed or registered

### Debug Mode

To run with detailed logging, modify the Program.cs to include console output for debugging:

```csharp
Console.WriteLine($"Processing method: {methodName}");
```

## License

This project is provided as-is for educational and development purposes.

## Contributing

Contributions are welcome! Please ensure all tests pass before submitting changes. 