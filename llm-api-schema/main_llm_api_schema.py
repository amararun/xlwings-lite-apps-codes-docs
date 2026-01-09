import xlwings as xw
from xlwings import script
import json
import pandas as pd
import traceback
import requests  # Import requests instead of pyfetch
import matplotlib.pyplot as plt
import seaborn as sns
import tempfile
import os
import numpy as np
from scipy import stats
import httpx  # Replace aiohttp with httpx

@script
async def analyze_table_schema_gemini(book: xw.Book):
    """Analyze table schema using Gemini API to identify categorical and numeric variables."""
    print("⭐⭐⭐ STARTING analyze_table_schema ⭐⭐⭐")
    
    # Read MASTER sheet parameters
    master_sheet = book.sheets["MASTER"]
    table_name = master_sheet["B6"].value
    gemini_api_key = master_sheet["B4"].value
    gemini_model = master_sheet["B3"].value
    
    # Find table in workbook
    table = None
    for sheet in book.sheets:
        if table_name in [t.name for t in sheet.tables]:
            table = sheet.tables[table_name]
            break
    
    if not table:
        return
    
    # Get sample data
    sample_data = table.range[:10, :].options(pd.DataFrame, index=False).value
    
    print("\n📊 Sample Data Preview:")
    print("-" * 50)
    print(sample_data.head(3).to_string())
    print("-" * 50)
    
    # First API call for categorical/numeric classification
    first_prompt = f"""
    Analyze these rows from a dataset (including headers):
    
    {sample_data.to_string()}
    
    Task: Identify categorical and numeric variables for EDA.
    Rules:
    1. Exclude ID columns, meaningless numbers, and dates
    2. Categorical variables: Text columns and numeric codes representing categories
    3. Numeric variables: Continuous or discrete numbers suitable for statistical analysis
    
    Return ONLY a JSON object with exactly this structure:
    {{
        "categorical_variables": ["col1", "col2"],
        "numeric_variables": ["col3", "col4"]
    }}
    """
    
    # Second API call for PostgreSQL schema generation
    second_prompt = f"""
    You are a PostgreSQL schema generator. Analyze the data and output ONLY a JSON schema.
    
    Sample data (first few rows including headers):
    {sample_data.to_string()}
    
    Requirements:
    1. Use standard PostgreSQL types (TEXT, NUMERIC, DATE, TIMESTAMP. Don't use INTEGER)
    2. Column names must be SQL-safe (alphanumeric and underscores only)
    3. Use NUMERIC for all values that can have decimal points
    4. Use DATE or TIMESTAMP for dates
    5. Use TEXT when unsure
    
    For descriptions:
    1. Provide detailed descriptions up to 50 words
    2. Include observed data patterns (e.g., "Contains categorical values like 'High', 'Medium', 'Low'")
    3. Mention if values appear continuous, discrete, categorical, or temporal
    4. Include example values from the data where helpful
    5. Note any patterns like value ranges or common formats
    6. Mention if it appears to be an ID, code, or reference field
    
    Return ONLY a JSON object with exactly this structure:
    {{
        "columns": [
            {{"name": "column_name", "type": "postgresql_type", "description": "detailed description of data content and patterns"}}
        ]
    }}
    """
    
    # Prepare API payloads
    first_payload = {
        "contents": [{"parts": [{"text": first_prompt}]}],
        "generationConfig": {
            "temperature": 0,
            "response_mime_type": "application/json",
            "response_schema": {
                "type": "OBJECT",
                "properties": {
                    "categorical_variables": {
                        "type": "ARRAY",
                        "items": {"type": "STRING"}
                    },
                    "numeric_variables": {
                        "type": "ARRAY",
                        "items": {"type": "STRING"}
                    }
                },
                "required": ["categorical_variables", "numeric_variables"]
            }
        }
    }
    
    second_payload = {
        "contents": [{"parts": [{"text": second_prompt}]}],
        "generationConfig": {
            "temperature": 0,
            "response_mime_type": "application/json",
            "response_schema": {
                "type": "OBJECT",
                "properties": {
                    "columns": {
                        "type": "ARRAY",
                        "items": {
                            "type": "OBJECT",
                            "properties": {
                                "name": {"type": "STRING"},
                                "type": {"type": "STRING"},
                                "description": {"type": "STRING"}
                            },
                            "required": ["name", "type", "description"]
                        }
                    }
                },
                "required": ["columns"]
            }
        }
    }
    
    # Make API calls
    api_url = f"https://generativelanguage.googleapis.com/v1beta/models/{gemini_model}:generateContent?key={gemini_api_key}"
    
    first_response = requests.post(
        api_url,
        headers={"Content-Type": "application/json"},
        json=first_payload
    )
    
    second_response = requests.post(
        api_url,
        headers={"Content-Type": "application/json"},
        json=second_payload
    )
    
    if first_response.status_code == 200 and second_response.status_code == 200:
        # Process first response
        first_response_json = first_response.json()
        if 'candidates' in first_response_json:
            schema = first_response_json['candidates'][0]['content']['parts'][0]['text']
            print("\n📄 First API Response (Variable Classification):")
            print("-" * 50)
            print(schema)
            print("-" * 50)
            
            # Parse first schema
            schema = json.loads(schema)
            
            # Process second response
            second_response_json = second_response.json()
            if 'candidates' in second_response_json:
                pg_schema = second_response_json['candidates'][0]['content']['parts'][0]['text']
                print("\n📄 Second API Response (PostgreSQL Schema):")
                print("-" * 50)
                print(pg_schema)
                print("-" * 50)
                
                # Parse second schema
                pg_schema = json.loads(pg_schema)
                
                # Update MASTER sheet with both results
                master_sheet["B8"].value = str(schema['categorical_variables'])
                master_sheet["B9"].value = str(schema['numeric_variables'])
                
                # Format PostgreSQL schema information
                title_cell = master_sheet["B12"]
                title_cell.value = "Column Schema Information"
                title_range = master_sheet["B12:D12"]
                title_range.color = "#A7D9AB"
                
                headers = ["Column Name", "PostgreSQL Type", "Description"]
                master_sheet["B13"].value = headers
                
                schema_data = [[col["name"], col["type"], col["description"]] for col in pg_schema["columns"]]
                master_sheet["B14"].value = schema_data
                
                # Format as table
                table_range = master_sheet["B13"].resize(len(schema_data) + 1, len(headers))
                master_sheet.tables.add(table_range)
                
                print("\n✅ Schema analysis completed successfully!")
                return schema
        else:
            print(f"❌ API call failed with status {first_response.status_code} or {second_response.status_code}")
            print("Error:", first_response.text, second_response.text)
            return
    else:
        print(f"❌ API call failed with status {first_response.status_code} or {second_response.status_code}")
        print("Error:", first_response.text, second_response.text)
        return

@script
async def analyze_table_schema_openai(book: xw.Book):
    """Analyze table schema using OpenAI API to identify categorical and numeric variables."""
    print("⭐⭐⭐ STARTING OpenAI SCHEMA ANALYSIS ⭐⭐⭐")
    
    # Read MASTER sheet parameters
    master_sheet = book.sheets["MASTER"]
    table_name = master_sheet["B6"].value
    openai_api_key = master_sheet["D4"].value
    openai_model = master_sheet["D3"].value
    
    # Find table in workbook
    table = None
    for sheet in book.sheets:
        if table_name in [t.name for t in sheet.tables]:
            table = sheet.tables[table_name]
            break
    
    if not table:
        return
    
    # Get sample data
    sample_data = table.range[:10, :].options(pd.DataFrame, index=False).value
    
    print("\n📊 Sample Data Preview:")
    print("-" * 50)
    print(sample_data.head(3).to_string())
    print("-" * 50)
    
    # First API call for categorical/numeric classification
    first_payload = {
        "model": openai_model,
        "messages": [
            {
                "role": "system",
                "content": "You are an expert at analyzing data schemas. You will be given sample data and should identify categorical and numeric variables suitable for analysis."
            },
            {
                "role": "user",
                "content": f"""Analyze these rows from a dataset (including headers):
                
                {sample_data.to_string()}
                
                Task: Identify categorical and numeric variables for EDA.
                Rules:
                1. Exclude ID columns, meaningless numbers, and dates
                2. Categorical variables: Text columns and numeric codes representing categories
                3. Numeric variables: Continuous or discrete numbers suitable for statistical analysis"""
            }
        ],
        "response_format": {
            "type": "json_schema",
            "json_schema": {
                "name": "variable_classification",
                "schema": {
                    "type": "object",
                    "properties": {
                        "categorical_variables": {
                            "type": "array",
                            "items": {"type": "string"}
                        },
                        "numeric_variables": {
                            "type": "array",
                            "items": {"type": "string"}
                        }
                    },
                    "required": ["categorical_variables", "numeric_variables"],
                    "additionalProperties": False
                },
                "strict": True
            }
        }
    }
    
    # Second API call for PostgreSQL schema generation
    second_payload = {
        "model": openai_model,
        "messages": [
            {
                "role": "system",
                "content": "You are a PostgreSQL schema generator. You will analyze sample data and generate appropriate column definitions with detailed descriptions."
            },
            {
                "role": "user",
                "content": f"""Analyze this sample data and generate a PostgreSQL schema:
                
                {sample_data.to_string()}
                
                Requirements:
                1. Use standard PostgreSQL types (TEXT, NUMERIC, DATE, TIMESTAMP. Don't use INTEGER)
                2. Column names must be SQL-safe (alphanumeric and underscores only)
                3. Use NUMERIC for all values that can have decimal points
                4. Use DATE or TIMESTAMP for dates
                5. Use TEXT when unsure
                6. Provide detailed descriptions (up to 50 words) including data patterns, value types, and examples"""
            }
        ],
        "response_format": {
            "type": "json_schema",
            "json_schema": {
                "name": "postgresql_schema",
                "schema": {
                    "type": "object",
                    "properties": {
                        "columns": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "name": {"type": "string"},
                                    "type": {"type": "string"},
                                    "description": {"type": "string"}
                                },
                                "required": ["name", "type", "description"],
                                "additionalProperties": False
                            }
                        }
                    },
                    "required": ["columns"],
                    "additionalProperties": False
                },
                "strict": True
            }
        }
    }
    
    # Make API calls using httpx
    async with httpx.AsyncClient() as client:
        # First API call
        first_response = await client.post(
            "https://api.openai.com/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {openai_api_key}",
                "Content-Type": "application/json"
            },
            json=first_payload,
            timeout=30.0
        )
        
        if first_response.status_code == 200:
            first_response_json = first_response.json()
            # Extract the actual data from the response
            schema = json.loads(first_response_json['choices'][0]['message']['content'])
            print("\n📄 First API Response (Variable Classification):")
            print("-" * 50)
            print(json.dumps(schema, indent=2))
            print("-" * 50)
            
            # Second API call
            second_response = await client.post(
                "https://api.openai.com/v1/chat/completions",
                headers={
                    "Authorization": f"Bearer {openai_api_key}",
                    "Content-Type": "application/json"
                },
                json=second_payload,
                timeout=30.0
            )
            
            if second_response.status_code == 200:
                second_response_json = second_response.json()
                # Extract the actual data from the response
                pg_schema = json.loads(second_response_json['choices'][0]['message']['content'])
                print("\n📄 Second API Response (PostgreSQL Schema):")
                print("-" * 50)
                print(json.dumps(pg_schema, indent=2))
                print("-" * 50)
                
                # Update MASTER sheet with both results
                master_sheet["B8"].value = str(schema['categorical_variables'])
                master_sheet["B9"].value = str(schema['numeric_variables'])
                
                # Format PostgreSQL schema information
                title_cell = master_sheet["B12"]
                title_cell.value = "Column Schema Information"
                title_range = master_sheet["B12:D12"]
                title_range.color = "#A7D9AB"
                
                headers = ["Column Name", "PostgreSQL Type", "Description"]
                master_sheet["B13"].value = headers
                
                schema_data = [[col["name"], col["type"], col["description"]] for col in pg_schema["columns"]]
                master_sheet["B14"].value = schema_data
                
                # Format as table
                table_range = master_sheet["B13"].resize(len(schema_data) + 1, len(headers))
                master_sheet.tables.add(table_range)
                
                print("\n✅ Schema analysis completed successfully!")
                return schema
            else:
                print(f"❌ Second API call failed with status {second_response.status_code}")
                print("Error:", second_response.text)
        else:
            print(f"❌ First API call failed with status {first_response.status_code}")
            print("Error:", first_response.text)
    
    return None

@script
def perform_eda(book: xw.Book):
    """Perform comprehensive Exploratory Data Analysis (EDA) on the specified table."""
    print("⭐⭐⭐ STARTING EDA ANALYSIS ⭐⭐⭐")
    
    try:
        # Read parameters from MASTER sheet
        master_sheet = book.sheets["MASTER"]
        table_name = master_sheet["B6"].value
        
        # Get string values and convert them to lists
        cat_vars_str = master_sheet["B8"].value
        num_vars_str = master_sheet["B9"].value
        
        # Check if variables are properly set
        if not cat_vars_str or not num_vars_str:
            print("❌ ERROR: Missing variable classifications in MASTER sheet cells B8 and B9")
            print("Please run schema analysis first")
            return
        
        # Convert string representations of lists to actual lists
        categorical_vars = [var.strip().strip("'") for var in cat_vars_str.strip('[]').split(',')]
        numeric_vars = [var.strip().strip("'") for var in num_vars_str.strip('[]').split(',')]
        
        # Validate we have variables to analyze
        if not categorical_vars and not numeric_vars:
            print("❌ ERROR: No variables found for analysis")
            return
        
        # Find table in workbook
        table = None
        
        # First try to find the table directly
        try:
            table = book.tables[table_name]
        except Exception:
            # If not found, search in sheets
            for sheet in book.sheets:
                try:
                    if table_name in [t.name for t in sheet.tables]:
                        table = sheet.tables[table_name]
                        break
                except Exception:
                    continue
        
        if not table:
            print("❌ ERROR: Table not found")
            return
        
        # Load data into DataFrame
        df = table.range.options(pd.DataFrame, index=False).value
        
        # Create or get DISTROS and PLOT sheets
        distros_sheet = None
        plot_sheet = None
        
        # List of sheets to handle with table name suffix
        sheets_to_handle = [f"DISTROS_{table_name}", f"PLOT_{table_name}"]
        
        # First, delete existing sheets if they exist
        for sheet_name in sheets_to_handle:
            try:
                if sheet_name in [s.name for s in book.sheets]:
                    book.sheets[sheet_name].delete()
            except Exception:
                continue
        
        # Then create new sheets
        for sheet_name in sheets_to_handle:
            try:
                sheet = book.sheets.add(name=sheet_name)
                if "DISTROS" in sheet_name:
                    distros_sheet = sheet
                else:
                    plot_sheet = sheet
            except Exception as e:
                raise
        
        # Step 1: Numeric Variables Analysis
        print("\n📊 Writing numeric statistics...")
        current_row = 1
        
        # Calculate comprehensive statistics
        stats_df = pd.DataFrame(index=['count', 'mean', 'std', 'min', '25%', '50%', '75%', 'max',
                                     'skewness', 'kurtosis', 'variance',
                                     '0%', '1%', '5%', '10%', '20%', '30%', '40%', '50%', '60%',
                                     '70%', '80%', '90%', '95%', '99%', '100%',
                                     'missing_count', 'missing_pct'],
                              columns=numeric_vars)
        
        # Basic statistics
        numeric_df = df[numeric_vars]
        stats_df.loc['count'] = numeric_df.count()
        stats_df.loc['mean'] = numeric_df.mean()
        stats_df.loc['std'] = numeric_df.std()
        stats_df.loc['min'] = numeric_df.min()
        stats_df.loc['25%'] = numeric_df.quantile(0.25).values
        stats_df.loc['50%'] = numeric_df.quantile(0.50).values
        stats_df.loc['75%'] = numeric_df.quantile(0.75).values
        stats_df.loc['max'] = numeric_df.max()
        
        # Additional statistics
        stats_df.loc['skewness'] = numeric_df.skew()
        stats_df.loc['kurtosis'] = numeric_df.kurtosis()
        stats_df.loc['variance'] = numeric_df.var()
        
        # Percentiles
        for p in [0, 1, 5, 10, 20, 30, 40, 50, 60, 70, 80, 90, 95, 99, 100]:
            stats_df.loc[f'{p}%'] = numeric_df.quantile(p/100).values
        
        # Missing values
        stats_df.loc['missing_count'] = numeric_df.isnull().sum()
        stats_df.loc['missing_pct'] = (numeric_df.isnull().sum() / len(numeric_df) * 100).round(2)
        
        # Write numeric statistics
        try:
            # Write and format header
            title_cell = distros_sheet["A1"]
            title_cell.value = "Numeric Variables Analysis"
            title_range = distros_sheet["A1:C1"]
            title_range.color = "#A7D9AB"
            
            # Write column headers
            headers = ['Statistic'] + list(stats_df.columns)
            distros_sheet["A2"].value = headers
            
            # Write data starting from row 3
            stats_df.reset_index(inplace=True)
            stats_df.rename(columns={'index': 'Statistic'}, inplace=True)
            distros_sheet["A3"].value = stats_df.values.tolist()
            
            # Format as table including headers
            stats_range = distros_sheet["A2"].resize(len(stats_df) + 1, len(stats_df.columns))
            distros_sheet.tables.add(stats_range)
            
        except Exception:
            pass
        
        current_row = len(stats_df) + 5  # Add extra space after numeric stats
        
        # Step 2: Categorical Variables Analysis
        print("\n📊 Analyzing categorical variables...")
        
        for col in categorical_vars:
            try:
                value_counts = df[col].value_counts()
                value_pcts = df[col].value_counts(normalize=True) * 100
                
                # Create a DataFrame with counts and percentages
                cat_stats = pd.DataFrame({
                    'Value': value_counts.index,
                    'Count': value_counts.values,
                    'Percentage': value_pcts.values.round(2)
                })
                
                # Write and format category header
                title_cell = distros_sheet[f"A{current_row}"]
                title_cell.value = f"{col} Distribution"
                title_range = distros_sheet[f"A{current_row}:C{current_row}"]
                title_range.color = "#A7D9AB"
                
                # Write column headers
                headers = ['Value', 'Count', 'Percentage']
                distros_sheet[f"A{current_row + 1}"].value = headers
                
                # Write data starting from next row
                distros_sheet[f"A{current_row + 2}"].value = cat_stats.values.tolist()
                
                # Format as table including headers
                cat_range = distros_sheet[f"A{current_row + 1}"].resize(len(cat_stats) + 1, len(cat_stats.columns))
                distros_sheet.tables.add(cat_range)
                
            except Exception:
                pass
            
            current_row += len(cat_stats) + 4  # Add extra space between categories
        
        # Step 3: Correlation Matrix
        print("\n📊 Calculating correlation matrix...")
        current_row += 2
        
        try:
            # Write and format correlation header
            title_cell = distros_sheet[f"A{current_row}"]
            title_cell.value = "Correlation Matrix"
            title_range = distros_sheet[f"A{current_row}:C{current_row}"]
            title_range.color = "#A7D9AB"
            
            # Calculate correlation matrix
            corr_matrix = numeric_df.corr().round(2)
            corr_df = corr_matrix.reset_index().rename(columns={'index': 'Variable'})
            
            # Write column headers
            headers = ['Variable'] + list(corr_df.columns[1:])
            distros_sheet[f"A{current_row + 1}"].value = headers
            
            # Write data starting from next row
            distros_sheet[f"A{current_row + 2}"].value = corr_df.values.tolist()
            
            # Format as table including headers
            corr_range = distros_sheet[f"A{current_row + 1}"].resize(len(corr_df) + 1, len(corr_df.columns))
            distros_sheet.tables.add(corr_range)
            
        except Exception:
            pass
        
        # Step 4: Visualizations
        print("\n📊 Creating visualizations...")
        
        def plot_and_insert(fig, sheet, top_left_cell, name):
            try:
                temp_dir = tempfile.gettempdir()
                temp_path = os.path.join(temp_dir, f"{name}.png")
                fig.savefig(temp_path, dpi=150, bbox_inches='tight')
                anchor_cell = sheet[top_left_cell]
                sheet.pictures.add(temp_path, name=name, update=True, anchor=anchor_cell, format="png")
            except Exception:
                pass
            finally:
                plt.close(fig)
        
        # Write plots header
        plot_sheet["A1"].value = "Data Visualizations"
        
        # Set default style parameters
        plt.rcParams.update({
            'figure.facecolor': 'white',
            'axes.facecolor': 'white',
            'axes.grid': True,
            'grid.alpha': 0.3,
            'axes.labelsize': 8,
            'axes.titlesize': 10,
            'xtick.labelsize': 8,
            'ytick.labelsize': 8
        })
        
        # 1. Correlation Heatmap
        fig, ax = plt.subplots(figsize=(5, 3.5))
        im = ax.imshow(corr_matrix, cmap='coolwarm', aspect='auto', vmin=-1, vmax=1)
        
        # Add correlation values
        for i in range(len(corr_matrix)):
            for j in range(len(corr_matrix)):
                text = ax.text(j, i, f'{corr_matrix.iloc[i, j]:.2f}',
                             ha='center', va='center', color='black',
                             fontsize=8)
        
        # Customize appearance
        ax.set_xticks(range(len(corr_matrix.columns)))
        ax.set_yticks(range(len(corr_matrix.columns)))
        ax.set_xticklabels(corr_matrix.columns, rotation=45, ha='right')
        ax.set_yticklabels(corr_matrix.columns)
        ax.set_title("Correlation Matrix", pad=10, fontsize=8, weight='bold')
        
        # Add colorbar
        plt.colorbar(im, ax=ax, shrink=0.8)
        plt.tight_layout()
        plot_and_insert(fig, plot_sheet, "A2", "correlation_heatmap")
        
        # Start grid layout from row 20
        base_row = 20
        current_row = base_row
        col_positions = ['A', 'E', 'I']
        current_col_idx = 0
        
        # Function to get next plot position
        def get_next_position():
            nonlocal current_row, current_col_idx
            position = f"{col_positions[current_col_idx]}{current_row}"
            current_col_idx += 1
            if current_col_idx >= len(col_positions):
                current_col_idx = 0
                current_row += 15
            return position
        
        # 2. Box Plots for Numeric Variables
        for col in numeric_vars:
            fig, ax = plt.subplots(figsize=(3, 2.5))
            ax.boxplot(numeric_df[col].dropna(), vert=True, widths=0.7,
                      patch_artist=True,
                      boxprops=dict(facecolor='lightblue', color='gray'),
                      whiskerprops=dict(color='gray'),
                      capprops=dict(color='gray'),
                      medianprops=dict(color='red'))
            
            ax.set_title(col, pad=10, fontsize=8, weight='bold')
            ax.set_ylabel(col, fontsize=8)
            ax.grid(True, alpha=0.3)
            plt.tight_layout()
            plot_and_insert(fig, plot_sheet, get_next_position(), f"box_{col}")
        
        # Add some space between different types of plots
        if current_col_idx != 0:
            current_row += 15
            current_col_idx = 0
        
        # 3. Distribution Plots for Numeric Variables
        for col in numeric_vars:
            fig, ax = plt.subplots(figsize=(3, 2.5))
            ax.hist(numeric_df[col].dropna(), bins=30, density=True, alpha=0.7,
                   color='lightblue', edgecolor='gray')
            
            # Add KDE
            kde = stats.gaussian_kde(numeric_df[col].dropna())
            x_range = np.linspace(numeric_df[col].min(), numeric_df[col].max(), 100)
            ax.plot(x_range, kde(x_range), 'r-', lw=1.5)
            
            ax.set_title(col, pad=10, fontsize=8, weight='bold')
            ax.set_xlabel(col, fontsize=8)
            ax.set_ylabel('Density', fontsize=8)
            ax.grid(True, alpha=0.3)
            plt.tight_layout()
            plot_and_insert(fig, plot_sheet, get_next_position(), f"dist_{col}")
        
        # Add some space between different types of plots
        if current_col_idx != 0:
            current_row += 15
            current_col_idx = 0
        
        # 4. Bar Plots for Categorical Variables
        for col in categorical_vars:
            fig, ax = plt.subplots(figsize=(3, 2.5))
            value_counts = df[col].value_counts()
            
            bars = ax.bar(range(len(value_counts)), value_counts.values,
                         color='lightblue', edgecolor='gray', alpha=0.7)
            
            ax.set_title(col, pad=10, fontsize=8, weight='bold')
            ax.set_xticks(range(len(value_counts)))
            ax.set_xticklabels(value_counts.index, rotation=45, ha='right')
            ax.set_ylabel('Count', fontsize=8)
            ax.grid(True, alpha=0.3)
            
            # Add value labels on top of bars
            for bar in bars:
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width()/2., height,
                       f'{int(height)}',
                       ha='center', va='bottom', fontsize=8)
            
            plt.tight_layout()
            plot_and_insert(fig, plot_sheet, get_next_position(), f"bar_{col}")
        
        print("\n✅ EDA analysis completed successfully!")
        print("\n⭐⭐⭐ ENDING EDA ANALYSIS ⭐⭐⭐")
        
    except Exception as e:
        print(f"❌ ERROR during EDA: {str(e)}")
        print(traceback.format_exc())
        return 