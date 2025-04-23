from mcp.server.fastmcp import FastMCP, Context, Image
import pandas as pd
import matplotlib.pyplot as plt
import os
import io
from pathlib import Path
import numpy as np
from typing import List, Dict, Any, Optional

# Create an MCP server
mcp = FastMCP("Excel Analytics")

# Configure matplotlib to use a non-interactive backend
plt.switch_backend('Agg')

def get_desktop_path():
    """Get the path to the user's desktop"""
    username = os.path.expanduser("~").split(os.sep)[-1]
    return os.path.join("/Users", username, "Desktop")

@mcp.tool()
def list_excel_files() -> List[str]:
    """List all Excel files available on the desktop"""
    desktop_path = get_desktop_path()
    excel_files = []
    
    # Look for .xlsx and .xls files
    for ext in ['.xlsx', '.xls']:
        excel_files.extend([f for f in os.listdir(desktop_path) 
                           if f.endswith(ext)])
    
    return excel_files

@mcp.tool()
def read_excel_file(filename: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """
    Read an Excel file from the desktop
    
    Args:
        filename: Name of the Excel file (must include .xlsx or .xls extension)
        sheet_name: Optional sheet name to read (if not provided, reads first sheet)
        
    Returns:
        Dictionary containing file info and preview of the data
    """
    desktop_path = get_desktop_path()
    file_path = os.path.join(desktop_path, filename)
    
    if not os.path.exists(file_path):
        return {"error": f"File {filename} not found on desktop"}
    
    try:
        # Read the Excel file
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)
        
        # Get sheet names
        with pd.ExcelFile(file_path) as xls:
            all_sheets = xls.sheet_names
        
        # Get basic info and a preview of the data
        info = {
            "filename": filename,
            "sheet_name": sheet_name if sheet_name else all_sheets[0],
            "all_sheets": all_sheets,
            "rows": len(df),
            "columns": len(df.columns),
            "column_names": df.columns.tolist(),
            "preview": df.head(5).to_dict(orient="records"),
            "dtypes": {col: str(dtype) for col, dtype in df.dtypes.items()}
        }
        
        return info
    
    except Exception as e:
        return {"error": f"Error reading Excel file: {str(e)}"}

@mcp.tool()
def analyze_excel_data(filename: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """
    Perform descriptive analysis on an Excel file
    
    Args:
        filename: Name of the Excel file (must include .xlsx or .xls extension)
        sheet_name: Optional sheet name to analyze (if not provided, analyzes first sheet)
        
    Returns:
        Dictionary containing descriptive statistics and column analysis
    """
    desktop_path = get_desktop_path()
    file_path = os.path.join(desktop_path, filename)
    
    if not os.path.exists(file_path):
        return {"error": f"File {filename} not found on desktop"}
    
    try:
        # Read the Excel file
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)
        
        # Basic statistics for numeric columns
        numeric_stats = {}
        for col in df.select_dtypes(include=['number']).columns:
            numeric_stats[col] = {
                "mean": float(df[col].mean()) if not pd.isna(df[col].mean()) else None,
                "median": float(df[col].median()) if not pd.isna(df[col].median()) else None,
                "std": float(df[col].std()) if not pd.isna(df[col].std()) else None,
                "min": float(df[col].min()) if not pd.isna(df[col].min()) else None,
                "max": float(df[col].max()) if not pd.isna(df[col].max()) else None,
                "count": int(df[col].count()),
                "null_count": int(df[col].isna().sum()),
                "unique_values": int(df[col].nunique())
            }
        
        # Category analysis for categorical columns
        categorical_stats = {}
        for col in df.select_dtypes(include=['object', 'category']).columns:
            value_counts = df[col].value_counts().head(10).to_dict()
            categorical_stats[col] = {
                "unique_values": int(df[col].nunique()),
                "null_count": int(df[col].isna().sum()),
                "top_values": value_counts
            }
        
        # Date analysis for datetime columns
        date_stats = {}
        for col in df.select_dtypes(include=['datetime']).columns:
            if not df[col].empty and not df[col].isna().all():
                date_stats[col] = {
                    "min_date": df[col].min().isoformat() if not pd.isna(df[col].min()) else None,
                    "max_date": df[col].max().isoformat() if not pd.isna(df[col].max()) else None,
                    "range_days": (df[col].max() - df[col].min()).days if not pd.isna(df[col].min()) and not pd.isna(df[col].max()) else None,
                    "null_count": int(df[col].isna().sum())
                }
        
        # Correlation matrix for numeric columns
        corr_matrix = None
        numeric_cols = df.select_dtypes(include=['number']).columns
        if len(numeric_cols) >= 2:
            corr_df = df[numeric_cols].corr()
            corr_matrix = corr_df.to_dict()
        
        return {
            "filename": filename,
            "sheet_name": sheet_name,
            "row_count": len(df),
            "column_count": len(df.columns),
            "numeric_stats": numeric_stats,
            "categorical_stats": categorical_stats,
            "date_stats": date_stats,
            "correlation": corr_matrix
        }
    
    except Exception as e:
        return {"error": f"Error analyzing Excel file: {str(e)}"}

@mcp.tool()
def create_bar_chart(filename: str, x_column: str, y_column: str, 
                    sheet_name: Optional[str] = None, title: Optional[str] = None,
                    limit: int = 10) -> Image:
    """
    Create a bar chart from Excel data
    
    Args:
        filename: Name of the Excel file
        x_column: Column to use for x-axis categories
        y_column: Column to use for y-axis values
        sheet_name: Optional sheet name (uses first sheet if not specified)
        title: Optional chart title
        limit: Maximum number of bars to display (default 10)
        
    Returns:
        Bar chart image
    """
    desktop_path = get_desktop_path()
    file_path = os.path.join(desktop_path, filename)
    
    try:
        # Read the Excel file
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)
        
        # Ensure columns exist
        if x_column not in df.columns or y_column not in df.columns:
            raise ValueError(f"Columns {x_column} or {y_column} not found in the Excel file")
        
        # Sort and limit data points
        if pd.api.types.is_numeric_dtype(df[y_column]):
            df_sorted = df.sort_values(by=y_column, ascending=False).head(limit)
        else:
            df_sorted = df.head(limit)
        
        # Create the figure
        plt.figure(figsize=(10, 6))
        plt.bar(df_sorted[x_column], df_sorted[y_column])
        
        # Add labels and title
        plt.xlabel(x_column)
        plt.ylabel(y_column)
        plt.title(title if title else f"{y_column} by {x_column}")
        
        # Rotate x-axis labels if there are many categories
        if len(df_sorted) > 5:
            plt.xticks(rotation=45, ha='right')
        
        plt.tight_layout()
        
        # Save the figure to a bytes buffer
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        plt.close()
        
        # Return as an Image
        return Image(data=buf.getvalue(), format="png")
    
    except Exception as e:
        # Create an error image
        plt.figure(figsize=(10, 6))
        plt.text(0.5, 0.5, f"Error creating bar chart: {str(e)}", 
                 horizontalalignment='center', verticalalignment='center', fontsize=12)
        plt.axis('off')
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        plt.close()
        return Image(data=buf.getvalue(), format="png")

@mcp.tool()
def create_pie_chart(filename: str, labels_column: str, values_column: str,
                    sheet_name: Optional[str] = None, title: Optional[str] = None,
                    limit: int = 8) -> Image:
    """
    Create a pie chart from Excel data
    
    Args:
        filename: Name of the Excel file
        labels_column: Column to use for pie slice labels
        values_column: Column to use for pie slice values
        sheet_name: Optional sheet name (uses first sheet if not specified)
        title: Optional chart title
        limit: Maximum number of slices to display (default 8, others grouped as 'Other')
        
    Returns:
        Pie chart image
    """
    desktop_path = get_desktop_path()
    file_path = os.path.join(desktop_path, filename)
    
    try:
        # Read the Excel file
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)
        
        # Ensure columns exist
        if labels_column not in df.columns or values_column not in df.columns:
            raise ValueError(f"Columns {labels_column} or {values_column} not found in the Excel file")
        
        # Group by labels column and sum values
        pie_data = df.groupby(labels_column)[values_column].sum().reset_index()
        
        # Sort and prepare data
        pie_data = pie_data.sort_values(by=values_column, ascending=False)
        
        # If more than limit categories, group the rest as "Other"
        if len(pie_data) > limit:
            top_data = pie_data.iloc[:limit-1]
            other_sum = pie_data.iloc[limit-1:][values_column].sum()
            other_row = pd.DataFrame({labels_column: ['Other'], values_column: [other_sum]})
            pie_data = pd.concat([top_data, other_row])
        
        # Create the figure
        plt.figure(figsize=(10, 8))
        plt.pie(pie_data[values_column], labels=pie_data[labels_column], 
                autopct='%1.1f%%', startangle=90, shadow=True)
        
        # Equal aspect ratio ensures pie is circular
        plt.axis('equal')
        
        # Add title
        plt.title(title if title else f"Distribution of {values_column} by {labels_column}")
        
        # Save the figure to a bytes buffer
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        plt.close()
        
        # Return as an Image
        return Image(data=buf.getvalue(), format="png")
    
    except Exception as e:
        # Create an error image
        plt.figure(figsize=(10, 6))
        plt.text(0.5, 0.5, f"Error creating pie chart: {str(e)}", 
                 horizontalalignment='center', verticalalignment='center', fontsize=12)
        plt.axis('off')
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        plt.close()
        return Image(data=buf.getvalue(), format="png")

@mcp.tool()
def create_line_chart(filename: str, x_column: str, y_columns: List[str],
                     sheet_name: Optional[str] = None, title: Optional[str] = None) -> Image:
    """
    Create a line chart from Excel data
    
    Args:
        filename: Name of the Excel file
        x_column: Column to use for x-axis (typically a date or ordered category)
        y_columns: List of columns to plot as lines (can specify multiple for comparison)
        sheet_name: Optional sheet name (uses first sheet if not specified)
        title: Optional chart title
        
    Returns:
        Line chart image
    """
    desktop_path = get_desktop_path()
    file_path = os.path.join(desktop_path, filename)
    
    try:
        # Read the Excel file
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)
        
        # Ensure columns exist
        if x_column not in df.columns:
            raise ValueError(f"Column {x_column} not found in the Excel file")
        
        for col in y_columns:
            if col not in df.columns:
                raise ValueError(f"Column {col} not found in the Excel file")
        
        # Sort by x_column if it's a datetime
        if pd.api.types.is_datetime64_any_dtype(df[x_column]):
            df = df.sort_values(by=x_column)
        
        # Create the figure
        plt.figure(figsize=(12, 6))
        
        # Plot each Y column
        for col in y_columns:
            plt.plot(df[x_column], df[col], marker='o', linestyle='-', label=col)
        
        # Add labels, title and legend
        plt.xlabel(x_column)
        plt.ylabel("Value")
        plt.title(title if title else f"Trend of {', '.join(y_columns)} over {x_column}")
        plt.legend()
        
        # Rotate x-axis labels if there are many points
        if len(df) > 10:
            plt.xticks(rotation=45, ha='right')
        
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.tight_layout()
        
        # Save the figure to a bytes buffer
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        plt.close()
        
        # Return as an Image
        return Image(data=buf.getvalue(), format="png")
    
    except Exception as e:
        # Create an error image
        plt.figure(figsize=(10, 6))
        plt.text(0.5, 0.5, f"Error creating line chart: {str(e)}", 
                 horizontalalignment='center', verticalalignment='center', fontsize=12)
        plt.axis('off')
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        plt.close()
        return Image(data=buf.getvalue(), format="png")

@mcp.tool()
def create_scatter_plot(filename: str, x_column: str, y_column: str, 
                       color_column: Optional[str] = None,
                       sheet_name: Optional[str] = None, title: Optional[str] = None) -> Image:
    """
    Create a scatter plot from Excel data
    
    Args:
        filename: Name of the Excel file
        x_column: Column to use for x-axis
        y_column: Column to use for y-axis
        color_column: Optional column to color points by categories
        sheet_name: Optional sheet name (uses first sheet if not specified)
        title: Optional chart title
        
    Returns:
        Scatter plot image
    """
    desktop_path = get_desktop_path()
    file_path = os.path.join(desktop_path, filename)
    
    try:
        # Read the Excel file
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)
        
        # Ensure columns exist
        if x_column not in df.columns or y_column not in df.columns:
            raise ValueError(f"Columns {x_column} or {y_column} not found in the Excel file")
        
        if color_column and color_column not in df.columns:
            raise ValueError(f"Column {color_column} not found in the Excel file")
        
        # Create the figure
        plt.figure(figsize=(10, 6))
        
        # Plot with or without color categories
        if color_column:
            categories = df[color_column].unique()
            for category in categories:
                subset = df[df[color_column] == category]
                plt.scatter(subset[x_column], subset[y_column], label=str(category), alpha=0.7)
            plt.legend(title=color_column)
        else:
            plt.scatter(df[x_column], df[y_column], alpha=0.7)
        
        # Add labels and title
        plt.xlabel(x_column)
        plt.ylabel(y_column)
        plt.title(title if title else f"{y_column} vs {x_column}")
        
        plt.grid(True, linestyle='--', alpha=0.3)
        plt.tight_layout()
        
        # Save the figure to a bytes buffer
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        plt.close()
        
        # Return as an Image
        return Image(data=buf.getvalue(), format="png")
    
    except Exception as e:
        # Create an error image
        plt.figure(figsize=(10, 6))
        plt.text(0.5, 0.5, f"Error creating scatter plot: {str(e)}", 
                 horizontalalignment='center', verticalalignment='center', fontsize=12)
        plt.axis('off')
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        plt.close()
        return Image(data=buf.getvalue(), format="png")

@mcp.tool()
def create_histogram(filename: str, column: str, bins: int = 10,
                    sheet_name: Optional[str] = None, title: Optional[str] = None) -> Image:
    """
    Create a histogram from Excel data
    
    Args:
        filename: Name of the Excel file
        column: Column to create histogram from (must be numeric)
        bins: Number of bins to use (default 10)
        sheet_name: Optional sheet name (uses first sheet if not specified)
        title: Optional chart title
        
    Returns:
        Histogram image
    """
    desktop_path = get_desktop_path()
    file_path = os.path.join(desktop_path, filename)
    
    try:
        # Read the Excel file
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)
        
        # Ensure column exists and is numeric
        if column not in df.columns:
            raise ValueError(f"Column {column} not found in the Excel file")
        
        if not pd.api.types.is_numeric_dtype(df[column]):
            raise ValueError(f"Column {column} must be numeric to create a histogram")
        
        # Create the figure
        plt.figure(figsize=(10, 6))
        
        # Create histogram
        plt.hist(df[column].dropna(), bins=bins, alpha=0.7, edgecolor='black')
        
        # Add labels and title
        plt.xlabel(column)
        plt.ylabel("Frequency")
        plt.title(title if title else f"Distribution of {column}")
        
        # Add a mean line
        if not df[column].dropna().empty:
            mean_val = df[column].mean()
            plt.axvline(mean_val, color='red', linestyle='dashed', linewidth=1)
            plt.text(mean_val*1.01, plt.ylim()[1]*0.9, f'Mean: {mean_val:.2f}', color='red')
        
        plt.grid(True, linestyle='--', alpha=0.3)
        plt.tight_layout()
        
        # Save the figure to a bytes buffer
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        plt.close()
        
        # Return as an Image
        return Image(data=buf.getvalue(), format="png")
    
    except Exception as e:
        # Create an error image
        plt.figure(figsize=(10, 6))
        plt.text(0.5, 0.5, f"Error creating histogram: {str(e)}", 
                 horizontalalignment='center', verticalalignment='center', fontsize=12)
        plt.axis('off')
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        plt.close()
        return Image(data=buf.getvalue(), format="png")

@mcp.tool()
def create_box_plot(filename: str, columns: List[str], 
                   sheet_name: Optional[str] = None, title: Optional[str] = None) -> Image:
    """
    Create a box plot from Excel data
    
    Args:
        filename: Name of the Excel file
        columns: List of numeric columns to include in the box plot
        sheet_name: Optional sheet name (uses first sheet if not specified)
        title: Optional chart title
        
    Returns:
        Box plot image
    """
    desktop_path = get_desktop_path()
    file_path = os.path.join(desktop_path, filename)
    
    try:
        # Read the Excel file
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)
        
        # Check if columns exist
        missing_cols = [col for col in columns if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Columns not found: {', '.join(missing_cols)}")
        
        # Check if columns are numeric
        non_numeric = [col for col in columns if not pd.api.types.is_numeric_dtype(df[col])]
        if non_numeric:
            raise ValueError(f"Non-numeric columns: {', '.join(non_numeric)}")
        
        # Create the figure
        plt.figure(figsize=(10, 6))
        
        # Create box plot
        plt.boxplot([df[col].dropna() for col in columns], labels=columns)
        
        # Add labels and title
        plt.ylabel("Value")
        plt.title(title if title else "Box Plot Comparison")
        
        # Rotate x-axis labels if there are many columns
        if len(columns) > 5:
            plt.xticks(rotation=45, ha='right')
        
        plt.grid(True, linestyle='--', alpha=0.3, axis='y')
        plt.tight_layout()
        
        # Save the figure to a bytes buffer
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        plt.close()
        
        # Return as an Image
        return Image(data=buf.getvalue(), format="png")
    
    except Exception as e:
        # Create an error image
        plt.figure(figsize=(10, 6))
        plt.text(0.5, 0.5, f"Error creating box plot: {str(e)}", 
                 horizontalalignment='center', verticalalignment='center', fontsize=12)
        plt.axis('off')
        buf = io.BytesIO()
        plt.savefig(buf, format='png')
        buf.seek(0)
        plt.close()
        return Image(data=buf.getvalue(), format="png")

# Run the server when this script is executed directly
if __name__ == "__main__":
    mcp.run()
