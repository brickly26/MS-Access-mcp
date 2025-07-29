# Claude Desktop Configuration for Microsoft Access MCP Server

## 🎯 Configuration for Source Code Execution

Since you want Claude Desktop to run the MCP server directly from the source code (not a published executable), here's how to configure it:

## 📋 Claude Desktop Configuration

### **Method 1: Claude Desktop Settings UI**

1. **Open Claude Desktop**
2. **Go to Settings** → **MCP Servers**
3. **Add new server**:
   - **Name**: `access-mcp-server`
   - **Command**: `dotnet`
   - **Arguments**: `run --project "C:\Users\brickly\Desktop\MS-Access-MCP\MS.Access.MCP.Official" --configuration Release`
   - **Working Directory**: `C:\Users\brickly\Desktop\MS-Access-MCP`

### **Method 2: Configuration File**

1. **Find your Claude Desktop config file**:
   - Windows: `%APPDATA%\Claude\claude_desktop_config.json`

2. **Add this configuration**:
```json
{
  "mcpServers": {
    "access-mcp-server": {
      "command": "dotnet",
      "args": [
        "run",
        "--project",
        "C:\\Users\\brickly\\Desktop\\MS-Access-MCP\\MS.Access.MCP.Official",
        "--configuration",
        "Release"
      ],
      "cwd": "C:\\Users\\brickly\\Desktop\\MS-Access-MCP"
    }
  }
}
```

## ✅ **What This Configuration Does**

- **Runs the MCP server directly from source code** using `dotnet run`
- **Uses Release configuration** for better performance
- **Points to your specific project directory**
- **No published executable needed**

## 🚀 **Usage in Claude Desktop**

Once configured, you can use:

```
connect_access
{}
```

The server will automatically connect to your database at `C:\Users\brickly\Documents\Database1.accdb`.

## 🔧 **Prerequisites**

1. **.NET 8.0 SDK** must be installed on your system
2. **Microsoft Access Database Engine** must be installed (for database connectivity)
3. **Claude Desktop** must be restarted after configuration

## 🎯 **Benefits of Source Code Execution**

- ✅ **No published executable needed**
- ✅ **Always runs latest code changes**
- ✅ **Easier debugging and development**
- ✅ **Direct access to source code**

## 🚨 **Troubleshooting**

### **Issue: "dotnet command not found"**
**Solution**: Install .NET 8.0 SDK from Microsoft

### **Issue: "Project not found"**
**Solution**: Verify the project path is correct

### **Issue: "Database connection failed"**
**Solution**: Install Microsoft Access Database Engine

### **Issue: "Tool not found"**
**Solution**: Restart Claude Desktop after configuration

## 📋 **Test Configuration**

After setting up, test with:

```
connect_access
{}
```

Expected response (if Access Database Engine is installed):
```json
{
  "success": true,
  "message": "Connected to C:\\Users\\brickly\\Documents\\Database1.accdb",
  "connected": true
}
```

The MCP server is now configured to run directly from your source code! 🚀 