# MCP-Server for Data Analysis
This server mainly focuses on the tools. It also includes some demos of resources and prompt. FastMCP is used to define and run the server, as well as to expose each tool as an endpoint.

# Features
- Read excel file from the desktop
- Perform descriptive analysis
- Create plots (i.e. bar chart, pie chart, line chart, scatter plot, histogram and box plot)

# Tools
## Listing Excel 
Fileslist_excel_files(): Scans the desktop for .xlsx and .xls files and returns a list of filenames.
## Reading Excel Files
read_excel_file(filename, sheet_name=None): Reads the specified Excel file from the desktop. By default, reads the first sheet if sheet_name is not provided. Returns basic info about rows, columns, a preview of the first 5 rows, and data types.
## Analyzing Excel Files
analyze_excel_data(filename, sheet_name=None): Computes descriptive statistics for numeric, categorical, and date columns. Also calculates a correlation matrix for numeric columns if applicable.
## Chart Generation
Each of these tools reads from an Excel file and generates a plot using Matplotlib. The plots are returned as Image objects (in-memory PNG data) and can be handled downstream by whichever client is interacting with the MCP server.

create_bar_chart(filename, x_column, y_column, sheet_name=None, title=None, limit=10)

create_pie_chart(filename, labels_column, values_column, sheet_name=None, title=None, limit=8)

create_line_chart(filename, x_column, y_columns, sheet_name=None, title=None)

create_scatter_plot(filename, x_column, y_column, color_column=None, sheet_name=None, title=None)

create_histogram(filename, column, bins=10, sheet_name=None, title=None)

create_box_plot(filename, columns, sheet_name=None, title=None)

## MCP Server
When the script is run directly (i.e., python excel_analytics_server.py), it starts the FastMCP server and registers all the tools above. These tools can then be called by a client that supports the MCP interface.

# Usage with Claude Desktop
## Prerequisites
- Download Claude Desktop

## Step 1: Adding MCP to your python project
```
# Install uv
curl -LsSf https://astral.sh/uv/install.sh | sh

# Create a new project directory
uv init data_analysis_tool
cd data_analysis_tool

# Create and activate a virtual environment
uv venv
source .venv/bin/activate

# Install dependencies
uv add "mcp[cli]" httpx

```

## Step 2: Create the MCP Server
- Place the file excel_analytics_server.py under the directory data_analysis_tool
- Run
```
pip install "mcp[cli]" pandas matplotlib openpyxl
```

## Step 3: Test the Server
```
mcp dev excel_analytics_server.py
```

## Step 4: Connect MCP to Claude
- Add this to your claude_desktop_config.json
```
"data_analysis_tool": {
      "command": "/Users/username/.local/bin/uv", # remember to change the username
      "args": [
        "--directory",
        "xxx/data_analysis_tool", # change it to the directory of data_analysis_tool
        "run",
        "excel_analytics_server.py"
      ]
    }
```

## Step 5: Done!
- Try it in Claude!

### Sample 1:
![image](https://github.com/LindseyyyLi/MCP-Server/blob/main/img/sample_use.png)

### Sample 2:
![image](https://github.com/LindseyyyLi/MCP-Server/blob/main/img/sample_use1.png)

### Sample 3:
![image](https://github.com/LindseyyyLi/MCP-Server/blob/main/img/sample_use2.png)


# Resources
Demo:
https://github.com/user-attachments/assets/6bf023ae-7d73-4bfb-95dd-d648a19133a1





# Prompt
Demo:
https://github.com/user-attachments/assets/59596eae-b8d1-4e5f-951b-15d94d8d261f




















