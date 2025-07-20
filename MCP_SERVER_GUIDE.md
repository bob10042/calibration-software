# 🌐 MCP Server Setup & Usage Guide

## 📋 **Overview**

Model Context Protocol (MCP) servers provide enhanced functionality to Claude Code by enabling access to external tools, data sources, and project-specific capabilities. This guide covers the complete MCP setup for the CalTestGUI project.

---

## ✅ **MCP Configuration Status: COMPLETE**

### 🎯 **What's Configured:**

#### **1. Custom CalTestGUI MCP Server**
- **File**: `caltest_mcp_server.py`
- **Purpose**: Project-specific tools and information
- **Capabilities**:
  - Build automation
  - Project analysis
  - Code change tracking
  - Resource access

#### **2. Standard MCP Servers**
- **Filesystem Server**: Enhanced file operations
- **Git Server**: Version control integration
- **Web Search Server**: Enhanced search capabilities

---

## 🔧 **Installation & Setup**

### **Prerequisites**
- ✅ Node.js (for NPX-based servers)
- ✅ Python (for custom server)
- ✅ PowerShell 7 (already installed)

### **Setup Command**
```powershell
# Run the setup script
"C:\Program Files\PowerShell\7\pwsh.exe" -ExecutionPolicy Bypass -File setup_mcp.ps1
```

### **Manual Installation** (if needed)
```bash
# Install MCP servers globally
npm install -g @modelcontextprotocol/server-filesystem
npm install -g @modelcontextprotocol/server-git
npm install -g @modelcontextprotocol/server-web-search
```

---

## 📂 **Configuration Files**

### **MCP Configuration: `.claude/mcp.json`**
```json
{
  "servers": {
    "caltest-project": {
      "command": "python",
      "args": ["caltest_mcp_server.py"],
      "cwd": "/mnt/c/Users/bob43/Downloads/gui/CalTestGUIW",
      "env": {
        "PROJECT_ROOT": "/mnt/c/Users/bob43/Downloads/gui/CalTestGUIW"
      }
    },
    "filesystem": {
      "command": "npx",
      "args": ["-y", "@modelcontextprotocol/server-filesystem"],
      "env": {
        "ALLOWED_DIRECTORIES": "/mnt/c/Users/bob43/Downloads/gui"
      }
    },
    "git": {
      "command": "npx",
      "args": ["-y", "@modelcontextprotocol/server-git"],
      "env": {
        "GIT_REPO_PATH": "/mnt/c/Users/bob43/Downloads/gui/CalTestGUIW"
      }
    },
    "web-search": {
      "command": "npx",
      "args": ["-y", "@modelcontextprotocol/server-web-search"]
    }
  }
}
```

---

## 🚀 **Usage Guide**

### **Resource References**

#### **Project Information**
```bash
# Access project details
@caltest-project:caltest://project/info

# Read project changelog
@caltest-project:caltest://changelog

# Check current build status
@caltest-project:caltest://build/status
```

#### **File System Operations**
```bash
# Access project files
@filesystem:/mnt/c/Users/bob43/Downloads/gui/CalTestGUIW

# Read specific files
@filesystem:/mnt/c/Users/bob43/Downloads/gui/CalTestGUIW/frmMain.cs
```

### **Available Tools**

#### **1. Build Project Tool**
- **Name**: `build_project`
- **Description**: Build the CalTestGUI project
- **Parameters**:
  - `configuration`: "Debug" or "Release" (default: "Debug")
  - `clean`: Boolean to clean before building (default: false)

**Usage Example**:
```text
Use the build_project tool to compile CalTestGUI in Release mode with clean
```

#### **2. Get Project Info Tool**
- **Name**: `get_project_info`
- **Description**: Get comprehensive project information
- **Returns**: JSON with project details, technologies, recent improvements

**Usage Example**:
```text
Use the get_project_info tool to show me the current project status
```

#### **3. Analyze Code Changes Tool**
- **Name**: `analyze_code_changes`
- **Description**: Analyze recent code improvements and changes
- **Returns**: JSON with detailed analysis of recent modifications

**Usage Example**:
```text
Use the analyze_code_changes tool to review what improvements were made
```

---

## 📊 **Available Resources**

### **Project Resources**

| Resource URI | Description | Content |
|-------------|-------------|---------|
| `caltest://project/info` | Current project status | Technologies, improvements, structure |
| `caltest://changelog` | Recent changes | Complete changelog in markdown |
| `caltest://build/status` | Build information | Build status, file size, last modified |

### **File System Resources**
- Access any file in the allowed directories
- Read, analyze, and work with project files
- Enhanced file operations through MCP

---

## 🛠️ **Custom MCP Server Features**

### **CalTestGUI-Specific Capabilities**

#### **Project Information**
```json
{
  "name": "CalTestGUI",
  "description": "Calibration Test GUI Application",
  "recent_improvements": [
    "Fixed menu label colors (white → black)",
    "Standardized font sizes (12pt Bold)",
    "Added Meters landing page",
    "Upgraded to PowerShell 7.5.2"
  ],
  "technologies": [
    "C# Windows Forms",
    ".NET Framework 4.7.2",
    "National Instruments VISA",
    "Microsoft Office Interop"
  ]
}
```

#### **Build Integration**
- Automatic detection of PowerShell 7 build scripts
- Fallback to MSBuild if needed
- Real-time build status and error reporting
- Build time measurement and optimization

#### **Change Analysis**
- Track UI improvements
- Monitor technical enhancements
- Code quality assessment
- Documentation completeness

---

## 💡 **Usage Examples**

### **Example 1: Project Overview**
```text
Show me the project information using @caltest-project:caltest://project/info
```

### **Example 2: Build the Project**
```text
Use the build_project tool to build CalTestGUI in Debug mode with clean enabled
```

### **Example 3: Check Recent Changes**
```text
Show me @caltest-project:caltest://changelog and use analyze_code_changes to summarize improvements
```

### **Example 4: File Analysis**
```text
Read @filesystem:/mnt/c/Users/bob43/Downloads/gui/CalTestGUIW/CalTestGUIW/frmMain.cs and explain the recent menu fixes
```

---

## 🔍 **Troubleshooting**

### **Common Issues**

#### **MCP Server Not Found**
```bash
# Check if servers are installed
npm list -g | grep modelcontextprotocol

# Reinstall if needed
npm install -g @modelcontextprotocol/server-filesystem
```

#### **Python Server Issues**
```bash
# Test the custom server
cd /mnt/c/Users/bob43/Downloads/gui/CalTestGUIW/CalTestGUIW
python caltest_mcp_server.py
```

#### **Configuration Issues**
- Ensure `.claude/mcp.json` exists in the project root
- Verify file paths are correct for your system
- Check that all environment variables are properly set

### **Restart Requirements**
After MCP configuration changes:
1. **Restart Claude Code** to reload MCP servers
2. **Verify connection** by testing resource access
3. **Check server status** if tools aren't working

---

## 📈 **Benefits**

### **Enhanced Development Experience**
- **Context-Aware**: Project-specific tools and information
- **Automated Tasks**: Build, analyze, and manage project
- **Real-Time Data**: Live project status and build information
- **Integrated Workflow**: Seamless tool integration

### **Improved Productivity**
- **Quick Access**: Instant project information
- **Automated Building**: PowerShell 7 integration
- **Change Tracking**: Monitor improvements and modifications
- **Resource Management**: Enhanced file and data access

---

## 🎯 **Next Steps**

### **Immediate Actions**
1. ✅ MCP servers configured
2. ✅ Custom project server created
3. ✅ Setup script provided
4. 🎯 **Ready to use enhanced MCP functionality!**

### **Advanced Usage**
- **Create additional custom tools** for specific CalTestGUI tasks
- **Integrate with CI/CD** for automated testing
- **Extend resource types** for meter-specific data
- **Add authentication** for remote servers if needed

---

## 📚 **Additional Resources**

### **MCP Documentation**
- [MCP Official Documentation](https://modelcontextprotocol.io/)
- [Claude Code MCP Guide](https://docs.anthropic.com/en/docs/claude-code/mcp)

### **Server Repositories**
- [MCP Filesystem Server](https://github.com/modelcontextprotocol/servers/tree/main/src/filesystem)
- [MCP Git Server](https://github.com/modelcontextprotocol/servers/tree/main/src/git)
- [MCP Web Search Server](https://github.com/modelcontextprotocol/servers/tree/main/src/web-search)

---

*Last Updated: June 29, 2025*  
*MCP Configuration Version: 1.0*  
*Status: ✅ Production Ready*

---

## 🔧 **Quick Reference**

### **Most Common Commands**
```bash
# View project info
@caltest-project:caltest://project/info

# Build project
Use build_project tool with configuration="Debug" and clean=true

# Check changes
@caltest-project:caltest://changelog

# Read files
@filesystem:/path/to/file
```

### **Setup Commands**
```powershell
# Run MCP setup
"C:\Program Files\PowerShell\7\pwsh.exe" -ExecutionPolicy Bypass -File setup_mcp.ps1

# Test custom server
python caltest_mcp_server.py
```

MCP servers are now your powerful development companions, providing enhanced context and automation for the CalTestGUI project! 🚀