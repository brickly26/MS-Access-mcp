# Setting up Microsoft Access MCP Server with Claude Desktop

## Prerequisites

1. **Microsoft Access** installed on your system
2. **.NET 8.0 Runtime** installed
3. **Claude Desktop** installed and configured

## Step 1: Build the MCP Server

First, ensure the server builds successfully:

```bash
dotnet build MS.Access.MCP.Interop/MS.Access.MCP.Interop.csproj
dotnet build MS.Access.MCP.Official/MS.Access.MCP.Official.csproj
```

## Step 2: Create Standalone Executable

Create a standalone executable to avoid build output issues:

```bash
dotnet publish MS.Access.MCP.Official/MS.Access.MCP.Official.csproj -c Release -r win-x64 --self-contained false -o ./publish
```

## Step 3: Test the MCP Server

Run the test script to verify the server works:

```bash
test-mcp-server.bat
```

This should return a proper MCP initialize response.

## Step 4: Configure Claude Desktop

### Option A: Direct Executable Reference (Recommended)

1. Open Claude Desktop
2. Go to Settings → MCP Servers
3. Add a new server with these settings:

**Server Name:** `access-mcp-server`

**Command:** `C:\Users\brickly\Desktop\MS-Access-MCP\publish\MS.Access.MCP.Official.exe`

**Working Directory:** `C:\Users\brickly\Desktop\MS-Access-MCP\publish`

### Option B: Using Configuration File

1. Find your Claude Desktop configuration file:
   - Windows: `%APPDATA%\Claude\claude_desktop_config.json`
   - macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`

2. Add the MCP server configuration:

```json
{
  "mcpServers": {
    "access-mcp-server": {
      "command": "C:\\Users\\brickly\\Desktop\\MS-Access-MCP\\publish\\MS.Access.MCP.Official.exe",
      "cwd": "C:\\Users\\brickly\\Desktop\\MS-Access-MCP\\publish"
    }
  }
}
```

## Step 5: Verify Connection

1. Restart Claude Desktop
2. Open a new conversation
3. Ask Claude to list available tools: "What MCP tools do you have access to?"

You should see a response listing all the Access MCP tools.

## Step 6: Test with a Real Database

1. Create or obtain an Access database file (.accdb)
2. Ask Claude to connect to it:

```
Connect to the Access database at C:\path\to\your\database.accdb
```

3. Then ask Claude to list the tables:

```
List all tables in the connected database
```

## Available Tools

The MCP server provides these tools:

### Connection Management
- `connect_access` - Connect to an Access database
- `disconnect_access` - Disconnect from database
- `is_connected` - Check connection status

### Data Access
- `get_tables` - List all tables
- `get_queries` - List all queries
- `get_relationships` - List all relationships
- `create_table` - Create a new table
- `delete_table` - Delete a table

### COM Automation
- `launch_access` - Launch Access application
- `close_access` - Close Access application
- `get_forms` - List all forms
- `get_reports` - List all reports
- `get_macros` - List all macros
- `get_modules` - List all modules
- `open_form` - Open a form
- `close_form` - Close a form

### VBA Extensibility
- `get_vba_projects` - List VBA projects
- `get_vba_code` - Get VBA code from module
- `set_vba_code` - Set VBA code in module
- `add_vba_procedure` - Add VBA procedure
- `compile_vba` - Compile VBA code

### System Tables
- `get_system_tables` - List system tables
- `get_object_metadata` - Get object metadata

### Form Controls
- `form_exists` - Check if form exists
- `get_form_controls` - List form controls
- `get_control_properties` - Get control properties
- `set_control_property` - Set control property

### Persistence & Versioning
- `export_form_to_text` - Export form to text
- `import_form_from_text` - Import form from text
- `delete_form` - Delete form
- `export_report_to_text` - Export report to text
- `import_report_from_text` - Import report from text
- `delete_report` - Delete report

## Troubleshooting

### Common Issues

1. **"Command not found"**
   - Ensure .NET 8.0 is installed
   - Verify the executable path is correct

2. **"Access not installed"**
   - Install Microsoft Access
   - Ensure it's properly registered

3. **"Database not found"**
   - Verify the database path is correct
   - Ensure the file has .accdb or .mdb extension

4. **"Permission denied"**
   - Run Claude Desktop as administrator
   - Check file permissions

5. **"Unexpected token" errors**
   - This indicates the server is outputting non-JSON text
   - Use the standalone executable instead of `dotnet run`
   - Ensure no build output is being sent to stdout

### Debug Mode

To see detailed server output, modify the configuration to include:

```json
{
  "mcpServers": {
    "access-mcp-server": {
      "command": "C:\\Users\\brickly\\Desktop\\MS-Access-MCP\\publish\\MS.Access.MCP.Official.exe",
      "cwd": "C:\\Users\\brickly\\Desktop\\MS-Access-MCP\\publish"
    }
  }
}
```

## Example Conversations

### Basic Database Operations

```
User: Connect to the database at C:\temp\sample.accdb
Claude: I'll connect to the Access database for you.

User: List all tables in the database
Claude: Here are the tables in your database:
- Customers (15 records)
- Orders (45 records)
- Products (23 records)

User: Show me the structure of the Customers table
Claude: The Customers table has these fields:
- CustomerID (AutoNumber, Primary Key)
- CompanyName (Text, 255 chars)
- ContactName (Text, 255 chars)
- Phone (Text, 50 chars)
```

### VBA Operations

```
User: List all VBA modules in the database
Claude: I found these VBA modules:
- Module1 (Has code: true)
- Form_Module (Has code: false)

User: Show me the code in Module1
Claude: Here's the VBA code in Module1:
Sub TestProcedure()
    MsgBox "Hello World"
End Sub
```

### Form Operations

```
User: List all forms in the database
Claude: I found these forms:
- CustomerForm
- OrderForm
- ProductForm

User: Show me the controls in CustomerForm
Claude: The CustomerForm contains these controls:
- txtCompanyName (TextBox)
- txtContactName (TextBox)
- btnSave (CommandButton)
- btnCancel (CommandButton)
``` 