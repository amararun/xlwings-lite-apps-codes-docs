import xlwings as xw
import pandas as pd
from pyodide.ffi import to_js
from pyodide.http import pyfetch
import json
import urllib.parse
import js
import tempfile
import os
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split
from sklearn.metrics import roc_curve, auc, roc_auc_score, confusion_matrix, precision_score, recall_score, f1_score, accuracy_score, classification_report
from xgboost import XGBClassifier
import seaborn as sns

# FastAPI CORS Configuration for Excel/xlwings:
# When setting up a FastAPI server to work with xlwings in Excel, configure CORS as follows:
#
# from fastapi.middleware.cors import CORSMiddleware
#
# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=["*"],  # For development/testing, using "*" is simplest
#     # If specific origins are required, include:
#     # allow_origins=["https://addin.xlwings.org"],
#     allow_credentials=True,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )
#
# Note: The actual requests come from the Excel WebView2 browser component.
# If using "*" for origins, no additional configuration is needed.
# If specifying origins explicitly, ensure "https://addin.xlwings.org" is included.

@script
async def list_tables(book: xw.Book):
    """List all tables in the database using connection details from MASTER sheet."""
    print("⭐⭐⭐ STARTING list_tables ⭐⭐⭐")
    
    try:
        # Read connection details from MASTER sheet
        try:
            master_sheet = book.sheets["MASTER"]
            
            # Read API URL from B2
            api_url = master_sheet["B2"].value
            if not api_url:
                raise Exception("No API URL provided in cell B2")
            
            # Read connection parameters from MASTER sheet with updated cell references
            connection_params = {}
            connection_params["host"] = master_sheet["B3"].value
            connection_params["database"] = master_sheet["B4"].value
            connection_params["user"] = master_sheet["B5"].value
            connection_params["password"] = master_sheet["B6"].value
            
            # Add validation for port number - now in B7
            try:
                port_value = master_sheet["B7"].value
                if isinstance(port_value, (int, float)):
                    connection_params["port"] = int(port_value)
                else:
                    # Default to standard ports if invalid
                    db_type = str(master_sheet["B8"].value).lower()  # DBTYPE is in B8
                    connection_params["port"] = 5432 if db_type == "postgresql" else 3306
                    print(f"⚠️ Invalid port number, using default port: {connection_params['port']}")
            except Exception as port_err:
                # Default to PostgreSQL port if error
                connection_params["port"] = 5432
                print(f"⚠️ Error reading port, using default port 5432: {port_err}")
            
            connection_params["db_type"] = master_sheet["B8"].value  # DBTYPE is in B8
            schema = master_sheet["B9"].value or "public"  # Schema is in B9
            
            print(f"📊 Read connection details from MASTER sheet")
            print(f"📊 API URL: {api_url}")
            print(f"📊 Host: {connection_params['host']}")
            print(f"📊 Database: {connection_params['database']}")
            print(f"📊 User: {connection_params['user']}")
            print(f"📊 Port: {connection_params['port']}")
            print(f"📊 DB Type: {connection_params['db_type']}")
            print(f"📊 Schema: {schema}")
            
        except Exception as e:
            print(f"❌ ERROR: Failed to read from MASTER sheet: {str(e)}")
            print(f"❌ Make sure the MASTER sheet exists with connection details")
            return
        
        # Query to get all tables with column counts - different for MySQL vs PostgreSQL
        if connection_params["db_type"].lower() == "mysql":
            # MySQL query - doesn't use schema in the same way
            sql_query = f"""
            SELECT 
                TABLE_NAME as table_name,
                (SELECT COUNT(*) FROM information_schema.columns c 
                 WHERE c.table_schema = '{connection_params["database"]}' AND c.table_name = t.table_name) AS column_count
            FROM 
                information_schema.tables t
            WHERE 
                table_schema = '{connection_params["database"]}'
            ORDER BY
                table_name;
            """
            print(f"📊 Using MySQL query for database {connection_params['database']}")
        else:
            # PostgreSQL query
            sql_query = f"""
            SELECT 
                table_name,
                (SELECT COUNT(*) FROM information_schema.columns c 
                 WHERE c.table_schema = t.table_schema AND c.table_name = t.table_name) AS column_count
            FROM 
                information_schema.tables t
            WHERE 
                table_schema = '{schema}'
            ORDER BY
                table_name;
            """
            print(f"📊 Using PostgreSQL query for schema {schema}")
        
        # Add SQL query to connection parameters
        connection_params["sqlquery"] = sql_query
        
        # Create query string with proper encoding
        query_parts = []
        for k, v in connection_params.items():
            encoded_value = urllib.parse.quote(str(v))
            query_parts.append(f"{k}={encoded_value}")
        
        query_string = "&".join(query_parts)
        full_url = f"{api_url}?{query_string}"
        
        print(f"📤 DEBUG: Sending GET request to API...")
        
        # Make API request with blob response type
        response = await pyfetch(
            full_url,
            method="GET",
            headers={"Accept": "text/plain,application/json"},
            response_type="blob"
        )
        
        print(f"📥 DEBUG: Response status: {response.status}")
        
        # Create a result sheet
        sheet_name = f"{connection_params['database']}_TABLES"  # Using underscore instead of space
        sheet_name = sheet_name.upper()  # Convert to uppercase
        if sheet_name in [s.name for s in book.sheets]:
            book.sheets[sheet_name].delete()
        
        sheet = book.sheets.add(name=sheet_name)
        
        # Simple header - different based on db type
        if connection_params["db_type"].lower() == "mysql":
            sheet["A1"].value = f"MySQL Database Tables in {connection_params['database']}"
        else:
            sheet["A1"].value = f"Database Tables in {schema} schema"
        
        # Process the response
        if response.ok:
            # Get text content directly
            text_content = await response.text()
            print(f"📥 DEBUG: Response text preview: {text_content[:100]}...")
            
            # Process pipe-delimited content
            if "|" in text_content:
                try:
                    # Parse pipe-delimited content into rows and columns
                    lines = text_content.strip().split("\n")
                    headers = [h.strip() for h in lines[0].split("|")]
                    
                    data_rows = []
                    for line in lines[1:]:
                        if line.strip():  # Skip empty lines
                            data_rows.append([cell.strip() for cell in line.split("|")])
                    
                    # Create a DataFrame
                    if data_rows:
                        df = pd.DataFrame(data_rows, columns=headers)
                        print(f"📊 Found {len(df)} tables in database")
                        
                        # Convert numeric columns (especially the column_count)
                        for col in df.columns:
                            try:
                                df[col] = pd.to_numeric(df[col])
                            except:
                                pass  # Keep as string if not numeric
                        
                        # Display as table
                        sheet["A3"].value = df
                        
                        # Format as table - wrapped in try/except
                        try:
                            table_range = sheet["A3"].resize(len(df) + 1, len(df.columns))
                            sheet.tables.add(table_range)
                        except Exception as table_err:
                            print(f"⚠️ Could not format as table: {table_err}")
                            
                        # Update MASTER sheet table dropdown if it exists
                        try:
                            if len(df) > 0:
                                table_list = df["table_name"].tolist()
                                print(f"📊 Updating MASTER sheet with table list")
                                
                                # Update a dropdown list validation for the TABLE NAME cell
                                # This would create a dropdown in the MASTER sheet
                                table_cell = master_sheet["B10"]
                                
                                # Create dropdown list validation
                                if hasattr(table_cell, "validation"):
                                    # This works in desktop Excel but might not in xlwings lite
                                    try:
                                        table_cell.api.Validation.Delete()
                                        table_cell.api.Validation.Add(
                                            Type=3,  # xlValidateList
                                            Formula1=",".join(table_list)
                                        )
                                        print("✅ Added dropdown list validation to TABLE NAME cell")
                                    except Exception as val_err:
                                        print(f"⚠️ Could not add validation: {val_err}")
                                        # Alternative approach for xlwings lite
                                        print("⚠️ Tables found but dropdown validation not supported in xlwings lite")
                        except Exception as master_err:
                            print(f"⚠️ Could not update MASTER sheet: {master_err}")
                        
                        # Add explanation for how to view table data
                        sheet["A" + str(6 + len(df))].value = "To view data for a specific table:"
                        sheet["A" + str(7 + len(df))].value = "1. Select a table name from the list above"
                        sheet["A" + str(8 + len(df))].value = "2. Enter it in the TABLE NAME field on the MASTER sheet"
                        sheet["A" + str(9 + len(df))].value = "3. Run the 'get_table_data' function"
                    else:
                        sheet["A3"].value = "No tables found in database"
                
                except Exception as parse_err:
                    print(f"❌ DEBUG: Table parsing error: {parse_err}")
                    import traceback
                    print(traceback.format_exc())
                    sheet["A3"].value = "Error parsing table data:"
                    sheet["B3"].value = str(parse_err)
                    sheet["A4"].value = "Raw response:"
                    sheet["B4"].value = text_content
            else:
                sheet["A3"].value = "Unexpected format - raw content:"
                sheet["B3"].value = text_content
        else:
            sheet["A3"].value = "Error retrieving tables:"
            sheet["B3"].value = f"Status: {response.status}"
        
        print("✅ Table list retrieved successfully")
        
    except Exception as e:
        print(f"❌ DEBUG ERROR: {str(e)}")
        import traceback
        print(traceback.format_exc())
        print("⭐⭐⭐ ENDING list_tables with ERROR ⭐⭐⭐")
        return
    
    print("⭐⭐⭐ ENDING list_tables SUCCESSFULLY ⭐⭐⭐")

@script
async def get_table_data(book: xw.Book):
    """Get the data and metadata for a specific table using connection details from MASTER sheet."""
    print("⭐⭐⭐ STARTING get_table_data ⭐⭐⭐")
    
    try:
        # Read connection details from MASTER sheet
        try:
            master_sheet = book.sheets["MASTER"]
            
            # Read API URL from B2
            api_url = master_sheet["B2"].value
            if not api_url:
                raise Exception("No API URL provided in cell B2")
            
            # Read connection parameters from MASTER sheet with updated cell references
            connection_params = {}
            connection_params["host"] = master_sheet["B3"].value
            connection_params["database"] = master_sheet["B4"].value
            connection_params["user"] = master_sheet["B5"].value
            connection_params["password"] = master_sheet["B6"].value
            
            # Add validation for port number - now in B7
            try:
                port_value = master_sheet["B7"].value
                if isinstance(port_value, (int, float)):
                    connection_params["port"] = int(port_value)
                else:
                    # Default to standard ports if invalid
                    db_type = str(master_sheet["B8"].value).lower()  # DBTYPE is in B8
                    connection_params["port"] = 5432 if db_type == "postgresql" else 3306
                    print(f"⚠️ Invalid port number, using default port: {connection_params['port']}")
            except Exception as port_err:
                # Default to PostgreSQL port if error
                connection_params["port"] = 5432
                print(f"⚠️ Error reading port, using default port 5432: {port_err}")
            
            connection_params["db_type"] = master_sheet["B8"].value  # DBTYPE is in B8
            schema = master_sheet["B9"].value or "public"  # Schema is in B9
            
            # Read table name from B11
            table_name = master_sheet["B11"].value
            
            # Get number of records to pull from B12
            num_records = 100  # Default value
            try:
                if master_sheet["B12"].value:
                    num_records = int(master_sheet["B12"].value)
            except:
                print(f"⚠️ Could not parse number of records, using default: {num_records}")
            
            # Check if table name is provided
            if not table_name:
                # Prompt user for table name if not in MASTER sheet
                table_name = js.prompt("Enter table name to query:", "")
                if not table_name:
                    raise Exception("No table name provided")
                # Update MASTER sheet with table name
                master_sheet["B11"].value = table_name
            
            # Check for custom query in B14
            custom_query = None
            try:
                if "B14" in master_sheet.used_range.address:
                    custom_query = master_sheet["B14"].value
                    if custom_query:
                        print(f"📊 Using custom query from MASTER sheet")
            except Exception as cq_err:
                print(f"⚠️ Could not read custom query: {cq_err}")
            
            print(f"📊 Read connection details from MASTER sheet")
            print(f"📊 API URL: {api_url}")
            print(f"📊 Host: {connection_params['host']}")
            print(f"📊 Database: {connection_params['database']}")
            print(f"📊 User: {connection_params['user']}")
            print(f"📊 Port: {connection_params['port']}")
            print(f"📊 DB Type: {connection_params['db_type']}")
            print(f"📊 Schema: {schema}")
            print(f"📊 Table: {table_name}")
            
        except Exception as e:
            print(f"❌ ERROR: Failed to read from MASTER sheet: {str(e)}")
            import traceback
            print(traceback.format_exc())
            print(f"❌ Make sure the MASTER sheet exists with connection details")
            return
            
        # Create a result sheet
        sheet_name = f"INFO_{table_name}"  # Already using underscore
        sheet_name = sheet_name[:31].upper()  # Excel has a 31 character limit for sheet names, convert to uppercase
        
        if sheet_name in [s.name for s in book.sheets]:
            book.sheets[sheet_name].delete()
            print(f"🗑️ Deleted existing sheet: {sheet_name}")
        
        sheet = book.sheets.add(name=sheet_name)
        print(f"📄 Created new sheet named '{sheet_name}' for table information")
        
        # Prepare to get table metadata first
        is_mysql = connection_params["db_type"].lower() == "mysql"
        
        # STEP 1: Get table structure/metadata query
        if is_mysql:
            # MySQL table metadata query
            metadata_query = f"""
            SELECT 
                COLUMN_NAME as column_name,
                DATA_TYPE as data_type,
                CHARACTER_MAXIMUM_LENGTH as max_length,
                IS_NULLABLE as is_nullable,
                COLUMN_KEY as column_key,
                COLUMN_DEFAULT as default_value,
                EXTRA as extra
            FROM 
                INFORMATION_SCHEMA.COLUMNS
            WHERE 
                TABLE_SCHEMA = '{connection_params["database"]}' 
                AND TABLE_NAME = '{table_name}'
            ORDER BY 
                ORDINAL_POSITION;
            """
            print(f"📊 Using MySQL metadata query for table {table_name}")
        else:
            # PostgreSQL table metadata query
            metadata_query = f"""
            SELECT 
                column_name,
                data_type,
                character_maximum_length as max_length,
                is_nullable,
                column_default as default_value,
                (SELECT 
                    pg_catalog.pg_get_constraintdef(con.oid)
                FROM 
                    pg_catalog.pg_constraint con
                    INNER JOIN pg_catalog.pg_class rel ON rel.oid = con.conrelid
                    INNER JOIN pg_catalog.pg_namespace nsp ON nsp.oid = rel.relnamespace
                WHERE 
                    con.contype = 'p' 
                    AND rel.relname = '{table_name}'
                    AND nsp.nspname = '{schema}'
                    AND array_position(con.conkey, cols.ordinal_position) > 0
                LIMIT 1) as constraint_def
            FROM 
                information_schema.columns cols
            WHERE 
                table_schema = '{schema}' 
                AND table_name = '{table_name}'
            ORDER BY 
                ordinal_position;
            """
            print(f"📊 Using PostgreSQL metadata query for table {schema}.{table_name}")
        
        # Add metadata query to connection parameters
        metadata_params = connection_params.copy()
        metadata_params["sqlquery"] = metadata_query
        
        # Create query string with proper encoding
        query_parts = []
        for k, v in metadata_params.items():
            encoded_value = urllib.parse.quote(str(v))
            query_parts.append(f"{k}={encoded_value}")
        
        metadata_query_string = "&".join(query_parts)
        metadata_url = f"{api_url}?{metadata_query_string}"
        
        print(f"📤 DEBUG: Sending metadata GET request to API...")
        
        # Make API request for metadata
        metadata_response = await pyfetch(
            metadata_url,
            method="GET",
            headers={"Accept": "text/plain,application/json"},
            response_type="blob"
        )
        
        print(f"📥 DEBUG: Metadata response status: {metadata_response.status}")
        
        # Process metadata response
        if metadata_response.ok:
            metadata_content = await metadata_response.text()
            print(f"📥 DEBUG: Received metadata content with length: {len(metadata_content)}")
            
            # Parse metadata content and display it in the sheet
            if "|" in metadata_content:
                try:
                    lines = metadata_content.strip().split("\n")
                    headers = [h.strip() for h in lines[0].split("|")]
                    
                    data_rows = []
                    for line in lines[1:]:
                        if line.strip():  # Skip empty lines
                            data_rows.append([cell.strip() for cell in line.split("|")])
                    
                    # Create a DataFrame for metadata
                    if data_rows:
                        metadata_df = pd.DataFrame(data_rows, columns=headers)
                        print(f"📊 Found {len(metadata_df)} columns in table {table_name}")
                        
                        # Add a title for the metadata section
                        sheet["A1"].value = f"Table Structure: {table_name}"
                        
                        # STEP 1: Get table statistics first
                        if is_mysql:
                            # MySQL table statistics - using direct COUNT(*) for accuracy
                            stats_query = f"""
                            SELECT 
                                (SELECT COUNT(*) FROM {table_name}) as row_count,
                                t.DATA_LENGTH as data_size,
                                t.INDEX_LENGTH as index_size,
                                t.DATA_LENGTH + t.INDEX_LENGTH as total_size
                            FROM 
                                information_schema.TABLES t
                            WHERE 
                                t.TABLE_SCHEMA = '{connection_params["database"]}' 
                                AND t.TABLE_NAME = '{table_name}';
                            """
                        else:
                            # PostgreSQL table statistics - using direct COUNT(*) for accuracy
                            stats_query = f"""
                            SELECT 
                                (SELECT COUNT(*) FROM {schema}.{table_name}) as row_count,
                                pg_size_pretty(pg_relation_size(pg_class.oid)) AS table_size,
                                pg_size_pretty(pg_total_relation_size(pg_class.oid) - pg_relation_size(pg_class.oid)) AS index_size,
                                pg_size_pretty(pg_total_relation_size(pg_class.oid)) AS total_size
                            FROM 
                                pg_class
                            JOIN 
                                pg_namespace ON pg_namespace.oid = pg_class.relnamespace
                            WHERE 
                                pg_class.relname = '{table_name}'
                                AND pg_namespace.nspname = '{schema}';
                            """
                        
                        # Add stats query to connection parameters
                        stats_params = connection_params.copy()
                        stats_params["sqlquery"] = stats_query
                        
                        # Create query string for stats
                        stats_query_parts = []
                        for k, v in stats_params.items():
                            encoded_value = urllib.parse.quote(str(v))
                            stats_query_parts.append(f"{k}={encoded_value}")
                        
                        stats_query_string = "&".join(stats_query_parts)
                        stats_url = f"{api_url}?{stats_query_string}"
                        
                        print(f"📤 DEBUG: Sending stats GET request to API...")
                        
                        # Make API request for stats
                        stats_response = await pyfetch(
                            stats_url,
                            method="GET",
                            headers={"Accept": "text/plain,application/json"},
                            response_type="blob"
                        )
                        
                        # Add title for the stats section
                        sheet["A2"].value = "Table Statistics:"
                        current_row = 3
                        
                        if stats_response.ok:
                            stats_content = await stats_response.text()
                            if "|" in stats_content:
                                try:
                                    stats_lines = stats_content.strip().split("\n")
                                    if len(stats_lines) > 1:
                                        stats_headers = [h.strip() for h in stats_lines[0].split("|")]
                                        stats_values = [v.strip() for v in stats_lines[1].split("|")]
                                        
                                        # Display stats as key-value pairs
                                        for header, value in zip(stats_headers, stats_values):
                                            label = header.replace("_", " ").title()
                                            if header == "row_count":
                                                label = "Row Count (Exact)"
                                            elif "size" in header:
                                                label = label + " (Approximate)"
                                            sheet[f"A{current_row}"].value = label + ":"
                                            sheet[f"B{current_row}"].value = value
                                            current_row += 1
                                            
                                        print(f"📊 Added table statistics")
                                except Exception as stats_err:
                                    print(f"⚠️ Could not parse statistics: {stats_err}")
                                    sheet["A3"].value = "Error parsing statistics"
                                    current_row = 4
                        
                        # Add spacing after statistics
                        current_row += 2
                        
                        # Add Column Structure section
                        sheet[f"A{current_row}"].value = "Column Structure:"
                        current_row += 1
                        
                        # Display metadata as table
                        sheet[f"A{current_row}"].options(index=False).value = metadata_df
                        
                        try:
                            # Format as table
                            metadata_range = sheet[f"A{current_row}"].resize(len(metadata_df) + 1, len(metadata_df.columns))
                            sheet.tables.add(metadata_range)
                            print(f"📊 Formatted metadata as table")
                        except Exception as table_err:
                            print(f"⚠️ Could not format metadata as table: {table_err}")
                        
                        # Calculate where to start the data section
                        data_start_row = current_row + len(metadata_df) + 3  # Leave some space after metadata
                        
                        # STEP 2: Get sample data query
                        if custom_query:
                            data_query = custom_query
                            print(f"📊 Using custom query for data: {custom_query}")
                        else:
                            if is_mysql:
                                data_query = f"SELECT * FROM {table_name} LIMIT 20"
                                print(f"📊 Using MySQL query for data sample from {table_name}")
                            else:
                                data_query = f"SELECT * FROM {schema}.{table_name} LIMIT 20"
                                print(f"📊 Using PostgreSQL query for data sample from {schema}.{table_name}")
                        
                        # Add data query to connection parameters
                        data_params = connection_params.copy()
                        data_params["sqlquery"] = data_query
                        
                        # Create query string for data
                        data_query_parts = []
                        for k, v in data_params.items():
                            encoded_value = urllib.parse.quote(str(v))
                            data_query_parts.append(f"{k}={encoded_value}")
                        
                        data_query_string = "&".join(data_query_parts)
                        data_url = f"{api_url}?{data_query_string}"
                        
                        print(f"📤 DEBUG: Sending data GET request to API...")
                        
                        # Make API request for data
                        data_response = await pyfetch(
                            data_url,
                            method="GET",
                            headers={"Accept": "text/plain,application/json"},
                            response_type="blob"
                        )
                        
                        print(f"📥 DEBUG: Data response status: {data_response.status}")
                        
                        # Add title for the data section
                        sheet[f"A{data_start_row - 1}"].value = "Sample Data (First 20 Rows):"
                        
                        if data_response.ok:
                            data_content = await data_response.text()
                            print(f"📥 DEBUG: Received data content with length: {len(data_content)}")
                            
                            if "|" in data_content:
                                try:
                                    # Parse data content
                                    data_lines = data_content.strip().split("\n")
                                    data_headers = [h.strip() for h in data_lines[0].split("|")]
                                    
                                    sample_rows = []
                                    for line in data_lines[1:]:
                                        if line.strip():  # Skip empty lines
                                            sample_rows.append([cell.strip() for cell in line.split("|")])
                                    
                                    # Create a DataFrame for sample data
                                    if sample_rows:
                                        data_df = pd.DataFrame(sample_rows, columns=data_headers)
                                        print(f"📊 Retrieved {len(data_df)} sample rows from {table_name}")
                                        
                                        # Display sample data
                                        sheet[f"A{data_start_row}"].options(index=False).value = data_df
                                        
                                        try:
                                            # Format as table
                                            data_range = sheet[f"A{data_start_row}"].resize(len(data_df) + 1, len(data_df.columns))
                                            sheet.tables.add(data_range)
                                            print(f"📊 Formatted sample data as table")
                                        except Exception as table_err:
                                            print(f"⚠️ Could not format sample data as table: {table_err}")
                                    else:
                                        sheet[f"A{data_start_row}"].value = "No sample data found in table"
                                except Exception as parse_err:
                                    print(f"❌ DEBUG: Data parsing error: {parse_err}")
                                    sheet[f"A{data_start_row}"].value = "Error parsing sample data:"
                                    sheet[f"B{data_start_row}"].value = str(parse_err)
                            else:
                                sheet[f"A{data_start_row}"].value = "Sample data not in expected format"
                                sheet[f"A{data_start_row + 1}"].value = data_content[:1000]
                        else:
                            data_error = await data_response.text()
                            sheet[f"A{data_start_row}"].value = "Error retrieving sample data:"
                            sheet[f"B{data_start_row}"].value = f"Status: {data_response.status}"
                            sheet[f"A{data_start_row + 1}"].value = data_error[:1000]
                        
                        print("✅ Table data and metadata retrieved successfully")
                
                except Exception as parse_err:
                    print(f"❌ DEBUG: Metadata parsing error: {parse_err}")
                    import traceback
                    print(traceback.format_exc())
                    sheet["A1"].value = "Error parsing table metadata:"
                    sheet["B1"].value = str(parse_err)
            else:
                sheet["A1"].value = "Metadata not in expected format"
                sheet["B1"].value = metadata_content[:1000]
        else:
            metadata_error = await metadata_response.text()
            sheet["A1"].value = "Error retrieving table metadata:"
            sheet["B1"].value = f"Status: {metadata_response.status}"
            sheet["A2"].value = "Error details:"
            sheet["B2"].value = metadata_error[:1000]
            print(f"❌ Metadata API Error: {metadata_response.status}")
        
        print("✅ Table data and metadata retrieved successfully")
        
    except Exception as e:
        print(f"❌ DEBUG ERROR: {str(e)}")
        import traceback
        print(traceback.format_exc())
        print("⭐⭐⭐ ENDING get_table_data with ERROR ⭐⭐⭐")
        return
    
    print("⭐⭐⭐ ENDING get_table_data SUCCESSFULLY ⭐⭐⭐")

@script
async def get_random_records(book: xw.Book):
    """Get random records from the specified table using connection details from MASTER sheet."""
    print("⭐⭐⭐ STARTING get_random_records ⭐⭐⭐")
    
    try:
        # Read connection details from MASTER sheet
        try:
            master_sheet = book.sheets["MASTER"]
            
            # Read API URL from B2
            api_url = master_sheet["B2"].value
            if not api_url:
                raise Exception("No API URL provided in cell B2")
            
            # Read connection parameters from MASTER sheet with updated cell references
            connection_params = {}
            connection_params["host"] = master_sheet["B3"].value
            connection_params["database"] = master_sheet["B4"].value
            connection_params["user"] = master_sheet["B5"].value
            connection_params["password"] = master_sheet["B6"].value
            
            # Add validation for port number - now in B7
            try:
                port_value = master_sheet["B7"].value
                if isinstance(port_value, (int, float)):
                    connection_params["port"] = int(port_value)
                else:
                    # Default to standard ports if invalid
                    db_type = str(master_sheet["B8"].value).lower()  # DBTYPE is in B8
                    connection_params["port"] = 5432 if db_type == "postgresql" else 3306
                    print(f"⚠️ Invalid port number, using default port: {connection_params['port']}")
            except Exception as port_err:
                # Default to PostgreSQL port if error
                connection_params["port"] = 5432
                print(f"⚠️ Error reading port, using default port 5432: {port_err}")
            
            connection_params["db_type"] = master_sheet["B8"].value  # DBTYPE is in B8
            schema = master_sheet["B9"].value or "public"  # Schema is in B9
            
            # Read table name from B11
            table_name = master_sheet["B11"].value
            
            # Get number of records to pull from B12
            num_records = 100  # Default value
            try:
                if master_sheet["B12"].value:
                    num_records = int(master_sheet["B12"].value)
            except:
                print(f"⚠️ Could not parse number of records, using default: {num_records}")
            
            # Check if table name is provided
            if not table_name:
                # Prompt user for table name if not in MASTER sheet
                table_name = js.prompt("Enter table name to query:", "")
                if not table_name:
                    raise Exception("No table name provided")
                # Update MASTER sheet with table name
                master_sheet["B11"].value = table_name
            
            print(f"📊 Read connection details from MASTER sheet")
            print(f"📊 API URL: {api_url}")
            print(f"📊 Host: {connection_params['host']}")
            print(f"📊 Database: {connection_params['database']}")
            print(f"📊 User: {connection_params['user']}")
            print(f"📊 Port: {connection_params['port']}")
            print(f"📊 DB Type: {connection_params['db_type']}")
            print(f"📊 Schema: {schema}")
            print(f"�� Table: {table_name} (from cell B11)")
            print(f"📊 Number of records: {num_records} (from cell B12)")
            
        except Exception as e:
            print(f"❌ ERROR: Failed to read from MASTER sheet: {str(e)}")
            print(f"❌ Make sure the MASTER sheet exists with connection details")
            return
            
        print(f"📊 Querying {num_records} random records from table: {table_name}")
        
        # Create SQL query for random records - different for MySQL vs PostgreSQL
        if connection_params["db_type"].lower() == "mysql":
            # MySQL random records query
            sql_query = f"SELECT * FROM {table_name} ORDER BY RAND() LIMIT {num_records}"
            print(f"📊 Using MySQL random query for table {table_name}")
        else:
            # PostgreSQL random records query
            sql_query = f"SELECT * FROM {schema}.{table_name} ORDER BY RANDOM() LIMIT {num_records}"
            print(f"📊 Using PostgreSQL random query for table {schema}.{table_name}")
        
        # Add SQL query to connection parameters
        connection_params["sqlquery"] = sql_query
        
        # Create query string with proper encoding
        query_parts = []
        for k, v in connection_params.items():
            encoded_value = urllib.parse.quote(str(v))
            query_parts.append(f"{k}={encoded_value}")
        
        query_string = "&".join(query_parts)
        full_url = f"{api_url}?{query_string}"
        
        print(f"📤 DEBUG: Sending GET request to API...")
        
        # Make API request
        response = await pyfetch(
            full_url,
            method="GET",
            headers={"Accept": "text/plain,application/json"},
            response_type="blob"
        )
        
        print(f"📥 DEBUG: Response status: {response.status}")
        
        # Create a result sheet with the table name
        sheet_name = table_name.replace(" ", "_")  # Replace any spaces with underscores
        sheet_name = sheet_name[:31].upper()  # Excel has a 31 character limit for sheet names, convert to uppercase
        
        # Check if sheet exists, delete if it does
        if sheet_name in [s.name for s in book.sheets]:
            book.sheets[sheet_name].delete()
            print(f"🗑️ Deleted existing sheet: {sheet_name}")
        
        sheet = book.sheets.add(name=sheet_name)
        print(f"📄 Created new sheet: {sheet_name}")
        
        # Header with table information
        if connection_params["db_type"].lower() == "mysql":
            sheet["A1"].value = f"{num_records} Random Records from {table_name}"
        else:
            sheet["A1"].value = f"{num_records} Random Records from {schema}.{table_name}"
        
        # Process the response
        if response.ok:
            # Get text content directly
            text_content = await response.text()
            print(f"📥 DEBUG: Received content with length: {len(text_content)}")
            print(f"📥 DEBUG: Content preview: {text_content[:100]}...")
            
            # Process pipe-delimited content
            if "|" in text_content:
                try:
                    # Parse pipe-delimited content into rows and columns
                    lines = text_content.strip().split("\n")
                    headers = [h.strip() for h in lines[0].split("|")]
                    
                    data_rows = []
                    for line in lines[1:]:
                        if line.strip():  # Skip empty lines
                            data_rows.append([cell.strip() for cell in line.split("|")])
                    
                    # Create a DataFrame
                    if data_rows:
                        df = pd.DataFrame(data_rows, columns=headers)
                        print(f"📊 Created DataFrame with {len(df)} rows and {len(df.columns)} columns")
                        
                        # Convert numeric columns
                        for col in df.columns:
                            try:
                                df[col] = pd.to_numeric(df[col])
                            except:
                                pass  # Keep as string if not numeric
                        
                        # Display as table - DIRECTLY at A1 with index=False to skip the index column
                        sheet["A1"].options(index=False).value = df
                        
                        # Format as table if there are rows - starting at A1 not A3
                        if len(df) > 0:
                            try:
                                table_range = sheet["A1"].resize(len(df) + 1, len(df.columns))
                                sheet.tables.add(table_range)
                                print(f"📊 Formatted data as table starting at A1")
                            except Exception as table_err:
                                print(f"⚠️ Could not format as table: {table_err}")
                        
                        print(f"✅ Successfully displayed {len(df)} records from {table_name}")
                    else:
                        sheet["A1"].value = "No data rows found in response"
                        print("⚠️ No data rows found in the response")
                
                except Exception as parse_err:
                    print(f"❌ DEBUG: Table parsing error: {parse_err}")
                    import traceback
                    print(traceback.format_exc())
                    sheet["A1"].value = "Error parsing table data:"
                    sheet["B1"].value = str(parse_err)
            else:
                # Display raw text if not pipe-delimited
                sheet["A1"].value = "Raw Response (preview):"
                sheet["B1"].value = text_content[:1000]  # Show first 1000 chars only
                print("⚠️ Response not in expected pipe-delimited format")
        else:
            error_text = await response.text()
            sheet["A1"].value = "Error:"
            sheet["B1"].value = f"Status: {response.status}"
            print(f"❌ API Error: {response.status}")
        
        print("✅ Random records process completed")
        
    except Exception as e:
        print(f"❌ DEBUG ERROR: {str(e)}")
        import traceback
        print(traceback.format_exc())
        print("⭐⭐⭐ ENDING get_random_records with ERROR ⭐⭐⭐")
        return
    
    print("⭐⭐⭐ ENDING get_random_records SUCCESSFULLY ⭐⭐⭐")
    print("⭐⭐⭐ ENDING get_table_data SUCCESSFULLY ⭐⭐⭐")

@script
async def get_first_n_records(book: xw.Book):
    """Get the first N records from the specified table using connection details from MASTER sheet."""
    print("⭐⭐⭐ STARTING get_first_n_records ⭐⭐⭐")
    
    try:
        # Read connection details from MASTER sheet
        try:
            master_sheet = book.sheets["MASTER"]
            
            # Read API URL from B2
            api_url = master_sheet["B2"].value
            if not api_url:
                raise Exception("No API URL provided in cell B2")
            
            # Read connection parameters from MASTER sheet with updated cell references
            connection_params = {}
            connection_params["host"] = master_sheet["B3"].value
            connection_params["database"] = master_sheet["B4"].value
            connection_params["user"] = master_sheet["B5"].value
            connection_params["password"] = master_sheet["B6"].value
            
            # Add validation for port number - now in B7
            try:
                port_value = master_sheet["B7"].value
                if isinstance(port_value, (int, float)):
                    connection_params["port"] = int(port_value)
                else:
                    # Default to standard ports if invalid
                    db_type = str(master_sheet["B8"].value).lower()  # DBTYPE is in B8
                    connection_params["port"] = 5432 if db_type == "postgresql" else 3306
                    print(f"⚠️ Invalid port number, using default port: {connection_params['port']}")
            except Exception as port_err:
                # Default to PostgreSQL port if error
                connection_params["port"] = 5432
                print(f"⚠️ Error reading port, using default port 5432: {port_err}")
            
            connection_params["db_type"] = master_sheet["B8"].value  # DBTYPE is in B8
            schema = master_sheet["B9"].value or "public"  # Schema is in B9
            
            # Read table name from B11
            table_name = master_sheet["B11"].value
            
            # Get number of records to pull from B12
            num_records = 100  # Default value
            try:
                if master_sheet["B12"].value:
                    num_records = int(master_sheet["B12"].value)
            except:
                print(f"⚠️ Could not parse number of records, using default: {num_records}")
            
            # Check if table name is provided
            if not table_name:
                # Prompt user for table name if not in MASTER sheet
                table_name = js.prompt("Enter table name to query:", "")
                if not table_name:
                    raise Exception("No table name provided")
                # Update MASTER sheet with table name
                master_sheet["B11"].value = table_name
            
            print(f"📊 Read connection details from MASTER sheet")
            print(f"📊 API URL: {api_url}")
            print(f"📊 Host: {connection_params['host']}")
            print(f"📊 Database: {connection_params['database']}")
            print(f"📊 User: {connection_params['user']}")
            print(f"📊 Port: {connection_params['port']}")
            print(f"📊 DB Type: {connection_params['db_type']}")
            print(f"📊 Schema: {schema}")
            print(f"�� Table: {table_name} (from cell B11)")
            print(f"📊 Number of records: {num_records} (from cell B12)")
            
        except Exception as e:
            print(f"❌ ERROR: Failed to read from MASTER sheet: {str(e)}")
            print(f"❌ Make sure the MASTER sheet exists with connection details")
            return
            
        print(f"📊 Querying first {num_records} records from table: {table_name}")
        
        # Create SQL query for first N records - different for MySQL vs PostgreSQL
        if connection_params["db_type"].lower() == "mysql":
            # MySQL query - simple LIMIT clause
            sql_query = f"SELECT * FROM {table_name} LIMIT {num_records}"
            print(f"📊 Using MySQL query for first {num_records} records from table {table_name}")
        else:
            # PostgreSQL query with schema
            sql_query = f"SELECT * FROM {schema}.{table_name} LIMIT {num_records}"
            print(f"📊 Using PostgreSQL query for first {num_records} records from table {schema}.{table_name}")
        
        # Add SQL query to connection parameters
        connection_params["sqlquery"] = sql_query
        
        # Create query string with proper encoding
        query_parts = []
        for k, v in connection_params.items():
            encoded_value = urllib.parse.quote(str(v))
            query_parts.append(f"{k}={encoded_value}")
        
        query_string = "&".join(query_parts)
        full_url = f"{api_url}?{query_string}"
        
        print(f"📤 DEBUG: Sending GET request to API...")
        
        # Make API request
        response = await pyfetch(
            full_url,
            method="GET",
            headers={"Accept": "text/plain,application/json"},
            response_type="blob"
        )
        
        print(f"📥 DEBUG: Response status: {response.status}")
        
        # Create a result sheet with the table name
        sheet_name = table_name.replace(" ", "_")  # Replace any spaces with underscores
        sheet_name = sheet_name[:31].upper()  # Excel has a 31 character limit for sheet names, convert to uppercase
        
        # Check if sheet exists, delete if it does
        if sheet_name in [s.name for s in book.sheets]:
            book.sheets[sheet_name].delete()
            print(f"🗑️ Deleted existing sheet: {sheet_name}")
        
        sheet = book.sheets.add(name=sheet_name)
        print(f"📄 Created new sheet: {sheet_name}")
        
        # Header with table information
        if connection_params["db_type"].lower() == "mysql":
            sheet["A1"].value = f"{num_records} Random Records from {table_name}"
        else:
            sheet["A1"].value = f"{num_records} Random Records from {schema}.{table_name}"
        
        # Process the response
        if response.ok:
            # Get text content directly
            text_content = await response.text()
            print(f"📥 DEBUG: Received content with length: {len(text_content)}")
            print(f"📥 DEBUG: Content preview: {text_content[:100]}...")
            
            # Process pipe-delimited content
            if "|" in text_content:
                try:
                    # Parse pipe-delimited content into rows and columns
                    lines = text_content.strip().split("\n")
                    headers = [h.strip() for h in lines[0].split("|")]
                    
                    data_rows = []
                    for line in lines[1:]:
                        if line.strip():  # Skip empty lines
                            data_rows.append([cell.strip() for cell in line.split("|")])
                    
                    # Create a DataFrame
                    if data_rows:
                        df = pd.DataFrame(data_rows, columns=headers)
                        print(f"📊 Created DataFrame with {len(df)} rows and {len(df.columns)} columns")
                        
                        # Convert numeric columns
                        for col in df.columns:
                            try:
                                df[col] = pd.to_numeric(df[col])
                            except:
                                pass  # Keep as string if not numeric
                        
                        # Display as table - DIRECTLY at A1 with index=False to skip the index column
                        sheet["A1"].options(index=False).value = df
                        
                        # Format as table if there are rows - starting at A1 not A3
                        if len(df) > 0:
                            try:
                                table_range = sheet["A1"].resize(len(df) + 1, len(df.columns))
                                sheet.tables.add(table_range)
                                print(f"📊 Formatted data as table starting at A1")
                            except Exception as table_err:
                                print(f"⚠️ Could not format as table: {table_err}")
                        
                        print(f"✅ Successfully displayed {len(df)} records from {table_name}")
                    else:
                        sheet["A1"].value = "No data rows found in response"
                        print("⚠️ No data rows found in the response")
                
                except Exception as parse_err:
                    print(f"❌ DEBUG: Table parsing error: {parse_err}")
                    import traceback
                    print(traceback.format_exc())
                    sheet["A1"].value = "Error parsing table data:"
                    sheet["B1"].value = str(parse_err)
            else:
                # Display raw text if not pipe-delimited
                sheet["A1"].value = "Raw Response (preview):"
                sheet["B1"].value = text_content[:1000]  # Show first 1000 chars only
                print("⚠️ Response not in expected pipe-delimited format")
        else:
            error_text = await response.text()
            sheet["A1"].value = "Error:"
            sheet["B1"].value = f"Status: {response.status}"
            print(f"❌ API Error: {response.status}")
        
        print("✅ First N records process completed")
        
    except Exception as e:
        print(f"❌ DEBUG ERROR: {str(e)}")
        import traceback
        print(traceback.format_exc())
        print("⭐⭐⭐ ENDING get_first_n_records with ERROR ⭐⭐⭐")
        return
    
    print("⭐⭐⭐ ENDING get_first_n_records SUCCESSFULLY ⭐⭐⭐")

@script
async def get_custom_query(book: xw.Book):
    """Execute a custom SQL query using connection details from MASTER sheet."""
    print("⭐⭐⭐ STARTING get_custom_query ⭐⭐⭐")
    
    try:
        # Read connection details from MASTER sheet
        try:
            master_sheet = book.sheets["MASTER"]
            
            # Read API URL from B2
            api_url = master_sheet["B2"].value
            if not api_url:
                raise Exception("No API URL provided in cell B2")
            
            # Read connection parameters from MASTER sheet with updated cell references
            connection_params = {}
            connection_params["host"] = master_sheet["B3"].value
            connection_params["database"] = master_sheet["B4"].value
            connection_params["user"] = master_sheet["B5"].value
            connection_params["password"] = master_sheet["B6"].value
            
            # Add validation for port number - now in B7
            try:
                port_value = master_sheet["B7"].value
                if isinstance(port_value, (int, float)):
                    connection_params["port"] = int(port_value)
                else:
                    # Default to standard ports if invalid
                    db_type = str(master_sheet["B8"].value).lower()  # DBTYPE is in B8
                    connection_params["port"] = 5432 if db_type == "postgresql" else 3306
                    print(f"⚠️ Invalid port number, using default port: {connection_params['port']}")
            except Exception as port_err:
                # Default to PostgreSQL port if error
                connection_params["port"] = 5432
                print(f"⚠️ Error reading port, using default port 5432: {port_err}")
            
            connection_params["db_type"] = master_sheet["B8"].value  # DBTYPE is in B8
            schema = master_sheet["B9"].value or "public"  # Schema is in B9
            
            # Read table name from B11
            table_name = master_sheet["B11"].value
            
            # Read custom query from B14
            custom_query = master_sheet["B14"].value
            
            if not custom_query:
                raise Exception("No custom query provided in cell B14")
            
            # Remove any trailing semicolon from the query
            custom_query = custom_query.strip().rstrip(';')
            
            print(f"📊 Read connection details from MASTER sheet")
            print(f"📊 API URL: {api_url}")
            print(f"📊 Host: {connection_params['host']}")
            print(f"📊 Database: {connection_params['database']}")
            print(f"📊 User: {connection_params['user']}")
            print(f"📊 Port: {connection_params['port']}")
            print(f"📊 DB Type: {connection_params['db_type']}")
            print(f"📊 Schema: {schema}")
            print(f"📊 Table: {table_name}")
            print(f"📊 Custom Query: {custom_query}")
            
        except Exception as e:
            print(f"❌ ERROR: Failed to read from MASTER sheet: {str(e)}")
            import traceback
            print(traceback.format_exc())
            print(f"❌ Make sure the MASTER sheet exists with connection details")
            return
            
        # Create a result sheet
        sheet_name = f"{table_name}_CUSTOM"  # Using underscore
        sheet_name = sheet_name[:31].upper()  # Excel has a 31 character limit for sheet names, convert to uppercase
        
        if sheet_name in [s.name for s in book.sheets]:
            book.sheets[sheet_name].delete()
            print(f"🗑️ Deleted existing sheet: {sheet_name}")
        
        sheet = book.sheets.add(name=sheet_name)
        print(f"📄 Created new sheet named '{sheet_name}' for custom query results")
        
        # Add title for the query section
        sheet["A1"].value = f"Custom Query Results:"
        sheet["A2"].value = f"Query: {custom_query}"
        
        # Add query to connection parameters
        connection_params["sqlquery"] = custom_query
        
        # Create query string with proper encoding
        query_parts = []
        for k, v in connection_params.items():
            encoded_value = urllib.parse.quote(str(v))
            query_parts.append(f"{k}={encoded_value}")
        
        query_string = "&".join(query_parts)
        full_url = f"{api_url}?{query_string}"
        
        print(f"📤 DEBUG: Sending custom query GET request to API...")
        
        # Make API request
        response = await pyfetch(
            full_url,
            method="GET",
            headers={"Accept": "text/plain,application/json"},
            response_type="blob"
        )
        
        print(f"📥 DEBUG: Response status: {response.status}")
        
        if response.ok:
            # Get text content directly
            text_content = await response.text()
            print(f"📥 DEBUG: Received content with length: {len(text_content)}")
            
            # Process pipe-delimited content
            if "|" in text_content:
                try:
                    # Parse pipe-delimited content into rows and columns
                    lines = text_content.strip().split("\n")
                    headers = [h.strip() for h in lines[0].split("|")]
                    
                    data_rows = []
                    for line in lines[1:]:
                        if line.strip():  # Skip empty lines
                            data_rows.append([cell.strip() for cell in line.split("|")])
                    
                    # Create a DataFrame
                    if data_rows:
                        df = pd.DataFrame(data_rows, columns=headers)
                        print(f"📊 Created DataFrame with {len(df)} rows and {len(df.columns)} columns")
                        
                        # Convert numeric columns
                        for col in df.columns:
                            try:
                                df[col] = pd.to_numeric(df[col])
                            except:
                                pass  # Keep as string if not numeric
                        
                        # Display as table - starting at A4 (after query text)
                        sheet["A4"].options(index=False).value = df
                        
                        # Format as table if there are rows
                        if len(df) > 0:
                            try:
                                table_range = sheet["A4"].resize(len(df) + 1, len(df.columns))
                                sheet.tables.add(table_range)
                                print(f"📊 Formatted data as table")
                            except Exception as table_err:
                                print(f"⚠️ Could not format as table: {table_err}")
                        
                        print(f"✅ Successfully displayed {len(df)} records from custom query")
                    else:
                        sheet["A4"].value = "No data rows found in response"
                        print("⚠️ No data rows found in the response")
                
                except Exception as parse_err:
                    print(f"❌ DEBUG: Table parsing error: {parse_err}")
                    import traceback
                    print(traceback.format_exc())
                    sheet["A4"].value = "Error parsing query results:"
                    sheet["B4"].value = str(parse_err)
            else:
                # Display raw text if not pipe-delimited
                sheet["A4"].value = "Raw Response (preview):"
                sheet["B4"].value = text_content[:1000]  # Show first 1000 chars only
                print("⚠️ Response not in expected pipe-delimited format")
        else:
            error_text = await response.text()
            sheet["A4"].value = "Error:"
            sheet["B4"].value = f"Status: {response.status}"
            sheet["A5"].value = "Error details:"
            sheet["B5"].value = error_text[:1000]
            print(f"❌ API Error: {response.status}")
        
        print("✅ Custom query process completed")
        
    except Exception as e:
        print(f"❌ DEBUG ERROR: {str(e)}")
        import traceback
        print(traceback.format_exc())
        print("⭐⭐⭐ ENDING get_custom_query with ERROR ⭐⭐⭐")
        return
    
    print("⭐⭐⭐ ENDING get_custom_query SUCCESSFULLY ⭐⭐⭐")

@script
def score_and_deciles(book: xw.Book):
    """Score and create deciles for the specified table using XGBoost model."""
    print("\U0001F4CC Step 1: Reading table name from MASTER sheet...")

    # Read table name from MASTER sheet
    try:
        master_sheet = book.sheets["MASTER"]
        table_name = master_sheet["B16"].value
        if not table_name:
            raise Exception("No table name provided in cell B16 of MASTER sheet")
        print(f"✅ Found table name: {table_name}")
    except Exception as e:
        raise Exception(f"Failed to read table name from MASTER sheet: {str(e)}")

    # Search for table across all sheets
    target_table = None
    target_sheet = None
    
    for sheet in book.sheets:
        try:
            if table_name in sheet.tables:
                target_table = sheet.tables[table_name]
                target_sheet = sheet
                print(f"✅ Found table '{table_name}' in sheet: {sheet.name}")
                break
        except Exception as e:
            continue
    
    if not target_table:
        raise Exception(f"Could not find table '{table_name}' in any sheet. Please make sure the table exists and is named correctly.")
    
    df_orig = target_table.range.options(pd.DataFrame, index=False).value
    df = df_orig.copy()
    print(f"✅ Loaded table into DataFrame with shape: {df.shape}")

    # Step 2: Prepare features and target
    # Drop cust_id and response_tag, keep all other columns as features
    X = df.drop(columns=["cust_id", "response_tag"])
    y = df["response_tag"].astype(int)
    print("🎯 Extracted features and target")

    # Step 3: One-hot encode categorical columns
    categorical_columns = ["job", "marital", "education", "default", "housing", "loan", "contact", 
                         "last_contact_month_of_year", "outcome_of_previous_campaign"]
    X_encoded = pd.get_dummies(X, columns=categorical_columns, drop_first=True)
    print(f"🔢 Encoded features. Shape: {X_encoded.shape}")

    # Step 4: Split
    X_train, X_test, y_train, y_test = train_test_split(
        X_encoded, y, test_size=0.3, random_state=42
    )
    print(f"📊 Train size: {len(X_train)}, Test size: {len(X_test)}")

    # Step 5: Train XGBoost
    model = XGBClassifier(max_depth=1, n_estimators=10, use_label_encoder=False,
                          eval_metric='logloss', verbosity=0)
    model.fit(X_train, y_train)
    print("🌲 Model trained successfully.")

    # Step 6: Score train/test
    train_probs = model.predict_proba(X_train)[:, 1]
    test_probs = model.predict_proba(X_test)[:, 1]
    train_preds = model.predict(X_train)
    test_preds = model.predict(X_test)

    # Calculate confusion matrices
    train_cm = confusion_matrix(y_train, train_preds)
    test_cm = confusion_matrix(y_test, test_preds)

    # Calculate metrics
    train_metrics = {
        'Accuracy': accuracy_score(y_train, train_preds),
        'Precision': precision_score(y_train, train_preds),
        'Recall': recall_score(y_train, train_preds),
        'F1 Score': f1_score(y_train, train_preds)
    }

    test_metrics = {
        'Accuracy': accuracy_score(y_test, test_preds),
        'Precision': precision_score(y_test, test_preds),
        'Recall': recall_score(y_test, test_preds),
        'F1 Score': f1_score(y_test, test_preds)
    }

    # Calculate classification reports
    train_report = classification_report(y_train, train_preds, output_dict=True)
    test_report = classification_report(y_test, test_preds, output_dict=True)

    # Step 7: Gini
    train_gini = 2 * roc_auc_score(y_train, train_probs) - 1
    test_gini = 2 * roc_auc_score(y_test, test_probs) - 1
    print(f"📈 Train Gini: {train_gini:.4f}")
    print(f"📊 Test Gini: {test_gini:.4f}")

    # Step 8: Create Model Evaluation Sheet
    print("📊 Creating Model Evaluation Sheet...")
    eval_sheet_name = f"{table_name}_Model_Eval".upper()
    if eval_sheet_name in [s.name for s in book.sheets]:
        try:
            book.sheets[eval_sheet_name].delete()
            print(f"🧹 Existing '{eval_sheet_name}' sheet deleted.")
        except Exception as e:
            print(f"⚠️ Could not delete existing sheet '{eval_sheet_name}': {e}")

    eval_sht = book.sheets.add(name=eval_sheet_name, after=target_sheet)

    # Add Gini Coefficients at the top
    eval_sht["A1"].value = "Model Performance Metrics"
    try:
        eval_sht["A1"].font.bold = True
        eval_sht["A1"].font.size = 14
    except Exception as e:
        print(f"⚠️ Could not format header: {e}")

    # Create Gini coefficients table
    gini_df = pd.DataFrame({
        'Metric': ['Train Gini', 'Test Gini'],
        'Value': [f"{train_gini:.4f}", f"{test_gini:.4f}"]
    })
    
    try:
        eval_sht["A3"].options(index=False).value = gini_df
        gini_range = eval_sht["A3"].resize(len(gini_df) + 1, len(gini_df.columns))
        eval_sht.tables.add(gini_range)
        print("✅ Formatted Gini coefficients as table")
    except Exception as table_err:
        print(f"⚠️ Could not format Gini coefficients as table: {table_err}")

    # Add Confusion Matrices header - start at row 8 (5 rows gap)
    eval_sht["A8"].value = "Confusion Matrices"
    try:
        eval_sht["A8"].font.bold = True
        eval_sht["A8"].font.size = 14
    except Exception as e:
        print(f"⚠️ Could not format header: {e}")

    # Training Set Confusion Matrix - start at row 10
    eval_sht["A10"].value = "Training Set"
    try:
        eval_sht["A10"].font.bold = True
    except Exception as e:
        print(f"⚠️ Could not format subheader: {e}")

    # Convert confusion matrix to DataFrame with proper column names
    train_cm_df = pd.DataFrame({
        'Actual/Predicted': ['Actual 0', 'Actual 1'],
        'Predicted 0': [train_cm[0][0], train_cm[1][0]],
        'Predicted 1': [train_cm[0][1], train_cm[1][1]]
    })
    
    try:
        eval_sht["A11"].options(index=False).value = train_cm_df
        train_cm_range = eval_sht["A11"].resize(len(train_cm_df) + 1, len(train_cm_df.columns))
        eval_sht.tables.add(train_cm_range, name=f"TrainConfusionMatrix")
        print("✅ Formatted training confusion matrix as table")
    except Exception as table_err:
        print(f"⚠️ Could not format training confusion matrix as table: {table_err}")

    # Test Set Confusion Matrix - start at row 16 (5 rows gap)
    eval_sht["A16"].value = "Test Set"
    try:
        eval_sht["A16"].font.bold = True
    except Exception as e:
        print(f"⚠️ Could not format subheader: {e}")

    # Convert confusion matrix to DataFrame with proper column names
    test_cm_df = pd.DataFrame({
        'Actual/Predicted': ['Actual 0', 'Actual 1'],
        'Predicted 0': [test_cm[0][0], test_cm[1][0]],
        'Predicted 1': [test_cm[0][1], test_cm[1][1]]
    })
    
    try:
        eval_sht["A17"].options(index=False).value = test_cm_df
        test_cm_range = eval_sht["A17"].resize(len(test_cm_df) + 1, len(test_cm_df.columns))
        eval_sht.tables.add(test_cm_range, name=f"TestConfusionMatrix")
        print("✅ Formatted test confusion matrix as table")
    except Exception as table_err:
        print(f"⚠️ Could not format test confusion matrix as table: {table_err}")

    # Add Classification Metrics header - start at row 22 (5 rows gap)
    eval_sht["A22"].value = "Classification Metrics"
    try:
        eval_sht["A22"].font.bold = True
        eval_sht["A22"].font.size = 14
    except Exception as e:
        print(f"⚠️ Could not format header: {e}")

    # Training Set Metrics - start at row 24
    eval_sht["A24"].value = "Training Set"
    try:
        eval_sht["A24"].font.bold = True
    except Exception as e:
        print(f"⚠️ Could not format subheader: {e}")

    # Convert metrics to DataFrame with proper column names
    train_metrics_df = pd.DataFrame({
        'Metric': list(train_metrics.keys()),
        'Value': [f"{v:.4f}" for v in train_metrics.values()]
    })
    
    try:
        eval_sht["A25"].options(index=False).value = train_metrics_df
        train_metrics_range = eval_sht["A25"].resize(len(train_metrics_df) + 1, len(train_metrics_df.columns))
        eval_sht.tables.add(train_metrics_range, name=f"TrainMetrics")
        print("✅ Formatted training metrics as table")
    except Exception as table_err:
        print(f"⚠️ Could not format training metrics as table: {table_err}")

    # Test Set Metrics - start at row 32 (7 rows gap due to larger previous table)
    eval_sht["A32"].value = "Test Set"
    try:
        eval_sht["A32"].font.bold = True
    except Exception as e:
        print(f"⚠️ Could not format subheader: {e}")

    # Convert metrics to DataFrame with proper column names
    test_metrics_df = pd.DataFrame({
        'Metric': list(test_metrics.keys()),
        'Value': [f"{v:.4f}" for v in test_metrics.values()]
    })
    
    try:
        eval_sht["A33"].options(index=False).value = test_metrics_df
        test_metrics_range = eval_sht["A33"].resize(len(test_metrics_df) + 1, len(test_metrics_df.columns))
        eval_sht.tables.add(test_metrics_range, name=f"TestMetrics")
        print("✅ Formatted test metrics as table")
    except Exception as table_err:
        print(f"⚠️ Could not format test metrics as table: {table_err}")

    # Add Detailed Classification Reports header - start at row 40 (7 rows gap)
    eval_sht["A40"].value = "Detailed Classification Reports"
    try:
        eval_sht["A40"].font.bold = True
        eval_sht["A40"].font.size = 14
    except Exception as e:
        print(f"⚠️ Could not format header: {e}")

    # Training Set Report - start at row 42
    eval_sht["A42"].value = "Training Set"
    try:
        eval_sht["A42"].font.bold = True
    except Exception as e:
        print(f"⚠️ Could not format subheader: {e}")

    # Convert classification report to DataFrame with proper formatting
    train_report_df = pd.DataFrame(train_report).transpose().reset_index()
    train_report_df.columns = ['Class'] + list(train_report_df.columns[1:])
    train_report_df = train_report_df.round(4)
    
    try:
        eval_sht["A43"].options(index=False).value = train_report_df
        train_report_range = eval_sht["A43"].resize(len(train_report_df) + 1, len(train_report_df.columns))
        eval_sht.tables.add(train_report_range, name=f"TrainClassificationReport")
        print("✅ Formatted training classification report as table")
    except Exception as table_err:
        print(f"⚠️ Could not format training classification report as table: {table_err}")

    # Test Set Report - start at row 50 (7 rows gap after training report)
    eval_sht["A50"].value = "Test Set"
    try:
        eval_sht["A50"].font.bold = True
    except Exception as e:
        print(f"⚠️ Could not format subheader: {e}")

    # Convert classification report to DataFrame with proper formatting
    test_report_df = pd.DataFrame(test_report).transpose().reset_index()
    test_report_df.columns = ['Class'] + list(test_report_df.columns[1:])
    test_report_df = test_report_df.round(4)
    
    try:
        eval_sht["A51"].options(index=False).value = test_report_df
        test_report_range = eval_sht["A51"].resize(len(test_report_df) + 1, len(test_report_df.columns))
        eval_sht.tables.add(test_report_range, name=f"TestClassificationReport")
        print("✅ Formatted test classification report as table")
    except Exception as table_err:
        print(f"⚠️ Could not format test classification report as table: {table_err}")

    # Format the sheets
    try:
        # Format Gini section
        eval_sht["A1:B4"].color = (240, 240, 240)
        eval_sht["A1"].font.bold = True
        eval_sht["A2"].font.bold = True

        # Format Confusion Matrix section
        eval_sht["A6:B15"].color = (245, 245, 245)
        eval_sht["A6"].font.bold = True
        eval_sht["A7"].font.bold = True
        eval_sht["A12"].font.bold = True

        # Format Classification Metrics section
        eval_sht["A17:B29"].color = (250, 250, 250)
        eval_sht["A17"].font.bold = True
        eval_sht["A18"].font.bold = True
        eval_sht["A24"].font.bold = True

        # Format Classification Reports section
        eval_sht["A30:B45"].color = (255, 255, 255)
        eval_sht["A30"].font.bold = True
        eval_sht["A31"].font.bold = True
        eval_sht["A37"].font.bold = True
    except Exception as e:
        print(f"⚠️ Could not apply formatting: {e}")

    # Step 8: Decile function
    def make_decile_table(probs, actuals):
        df_temp = pd.DataFrame({"prob": probs, "actual": actuals})
        df_temp["decile"] = pd.qcut(df_temp["prob"].rank(method="first", ascending=False), 10, labels=False) + 1
        grouped = df_temp.groupby("decile").agg(
            Obs=("actual", "count"),
            Min_Score=("prob", "min"),
            Max_Score=("prob", "max"),
            Avg_Score=("prob", "mean"),
            Responders=("actual", "sum")
        ).reset_index()
        grouped["Response_Rate(%)"] = round((grouped["Responders"] / grouped["Obs"]) * 100, 2)
        grouped["Cumulative_Responders"] = grouped["Responders"].cumsum()
        grouped["Cumulative_Response_%"] = round((grouped["Cumulative_Responders"] / grouped["Responders"].sum()) * 100, 2)
        return grouped

    train_decile = make_decile_table(train_probs, y_train)
    test_decile = make_decile_table(test_probs, y_test)
    print("📋 Created decile tables.")

    # Step 9: Insert deciles into new sheet
    print("📄 Preparing to insert decile tables into new sheet...")
    sheet_name = f"{table_name}_Score_Deciles".upper()
    existing_sheets = [s.name for s in book.sheets]

    if sheet_name in existing_sheets:
        try:
            book.sheets[sheet_name].delete()
            print(f"🧹 Existing '{sheet_name}' sheet deleted.")
        except Exception as e:
            print(f"⚠️ Could not delete existing sheet '{sheet_name}': {e}")

    new_sht = book.sheets.add(name=sheet_name, after=eval_sht)
    
    # Add and format headers with better error handling
    try:
        # Train Deciles header
        header_cell = new_sht["A1"]
        header_cell.value = "Train Deciles"
        header_range = header_cell.api
        header_range.Font.Bold = True
        header_range.Font.Size = 14
        print("✅ Added and formatted Train Deciles header")
    except Exception as format_err:
        print(f"⚠️ Could not fully format Train Deciles header: {format_err}")
        # Still try to set the value even if formatting fails
        try:
            new_sht["A1"].value = "Train Deciles"
        except:
            print("❌ Could not even set header text")

    new_sht["A2"].options(index=False).value = train_decile

    # Format train deciles as table
    try:
        train_table_range = new_sht["A2"].resize(len(train_decile), len(train_decile.columns))
        new_sht.tables.add(train_table_range)
        print(f"📊 Formatted train deciles as table")
    except Exception as table_err:
        print(f"⚠️ Could not format train deciles as table: {table_err}")

    start_row = train_decile.shape[0] + 4
    
    # Test Deciles header with same robust approach
    try:
        # Test Deciles header
        header_cell = new_sht[f"A{start_row}"]
        header_cell.value = "Test Deciles"
        header_range = header_cell.api
        header_range.Font.Bold = True
        header_range.Font.Size = 14
        print("✅ Added and formatted Test Deciles header")
    except Exception as format_err:
        print(f"⚠️ Could not fully format Test Deciles header: {format_err}")
        # Still try to set the value even if formatting fails
        try:
            new_sht[f"A{start_row}"].value = "Test Deciles"
        except:
            print("❌ Could not even set header text")

    new_sht[f"A{start_row+1}"].options(index=False).value = test_decile

    # Format test deciles as table
    try:
        test_table_range = new_sht[f"A{start_row+1}"].resize(len(test_decile), len(test_decile.columns))
        new_sht.tables.add(test_table_range)
        print(f"📊 Formatted test deciles as table")
    except Exception as table_err:
        print(f"⚠️ Could not format test deciles as table: {table_err}")

    print(f"🗘️ Decile tables inserted into sheet '{sheet_name}'")

    # Step 10: Score full dataset and append as new column
    full_probs = model.predict_proba(X_encoded)[:, 1]
    df_orig["SCORE_PROBABILITY"] = full_probs
    target_table.range.options(index=False).value = df_orig
    print("✅ Appended SCORE_PROBABILITY to original table without changing its structure.")

    # Step 11: Create and insert graphs into Excel
    graph_sheet_name = f"{table_name}_Score_Graphs".upper()
    if graph_sheet_name in existing_sheets:
        try:
            book.sheets[graph_sheet_name].delete()
            print(f"🧹 Existing '{graph_sheet_name}' sheet deleted.")
        except Exception as e:
            print(f"⚠️ Could not delete existing sheet '{graph_sheet_name}': {e}")

    graph_sht = book.sheets.add(name=graph_sheet_name, after=new_sht)

    def plot_and_insert(fig, sheet, top_left_cell, name):
        try:
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, f"{name}.png")
            fig.savefig(temp_path, dpi=150)
            print(f"🖼️ Saved plot '{name}' to {temp_path}")
            anchor_cell = sheet[top_left_cell]
            sheet.pictures.add(temp_path, name=name, update=True, anchor=anchor_cell, format="png")
            print(f"✅ Inserted plot '{name}' at {top_left_cell}")
        except Exception as e:
            print(f"❌ Failed to insert plot '{name}': {e}")
        finally:
            plt.close(fig)

    # ROC Curve
    def plot_roc(y_true, y_prob, label):
        fpr, tpr, _ = roc_curve(y_true, y_prob)
        roc_auc = auc(fpr, tpr)
        plt.plot(fpr, tpr, label=f'{label} (AUC = {roc_auc:.2f})')

    fig1 = plt.figure(figsize=(6, 4))
    plot_roc(y_train, train_probs, "Train")
    plot_roc(y_test, test_probs, "Test")
    plt.plot([0, 1], [0, 1], linestyle='--', color='gray', label='Random')
    plt.title("ROC Curve")
    plt.xlabel("False Positive Rate")
    plt.ylabel("True Positive Rate")
    plt.legend()
    plt.grid(True)
    plot_and_insert(fig1, graph_sht, "A1", name="ROC_Curve")

    # Cumulative Gain Curve
    def cumulative_gain_curve(y_true, y_prob):
        df = pd.DataFrame({'actual': y_true, 'prob': y_prob})
        df = df.sort_values('prob', ascending=False).reset_index(drop=True)
        df['cumulative_responders'] = df['actual'].cumsum()
        df['cumulative_pct_responders'] = df['cumulative_responders'] / df['actual'].sum()
        df['cumulative_pct_customers'] = (df.index + 1) / len(df)
        return df['cumulative_pct_customers'], df['cumulative_pct_responders']

    train_x, train_y = cumulative_gain_curve(y_train, train_probs)
    test_x, test_y = cumulative_gain_curve(y_test, test_probs)

    fig2 = plt.figure(figsize=(6, 4))
    plt.plot(train_x, train_y, label="Train")
    plt.plot(test_x, test_y, label="Test")
    plt.plot([0, 1], [0, 1], linestyle="--", color="gray", label="Random")
    plt.title("Cumulative Gain (Decile) Curve")
    plt.xlabel("Cumulative % of Customers")
    plt.ylabel("Cumulative % of Responders")
    plt.legend()
    plt.grid(True)
    plot_and_insert(fig2, graph_sht, "A20", name="Gain_Curve")

    print(f"📊 Graphs added to sheet '{graph_sheet_name}'.")

    for sht in book.sheets:
        print(f"📄 Sheet found: {sht.name}")

    try:
        book.save()
        print("📅 Workbook saved successfully.")
    except Exception as e:
        print(f"❌ Failed to save workbook: {e}")

@script
def perform_eda(book: xw.Book):
    """Perform comprehensive Exploratory Data Analysis (EDA) on the specified table."""
    print("⭐⭐⭐ STARTING EDA ANALYSIS ⭐⭐⭐")
    
    try:
        # Read table name from MASTER sheet
        try:
            master_sheet = book.sheets["MASTER"]
            table_name = master_sheet["B16"].value
            if not table_name:
                raise Exception("No table name provided in cell B16 of MASTER sheet")
            print(f"✅ Found table name: {table_name}")
        except Exception as e:
            raise Exception(f"Failed to read table name from MASTER sheet: {str(e)}")

        # Search for table across all sheets
        target_table = None
        target_sheet = None
        
        for sheet in book.sheets:
            try:
                if table_name in sheet.tables:
                    target_table = sheet.tables[table_name]
                    target_sheet = sheet
                    print(f"✅ Found table '{table_name}' in sheet: {sheet.name}")
                    break
            except Exception as e:
                continue
        
        if not target_table:
            raise Exception(f"Could not find table '{table_name}' in any sheet. Please make sure the table exists and is named correctly.")
        
        # Load data into DataFrame
        df = target_table.range.options(pd.DataFrame, index=False).value
        print(f"✅ Loaded table into DataFrame with shape: {df.shape}")
        print(f"📊 Original columns: {df.columns.tolist()}")

        # Define categorical variables (same as in scoring model)
        categorical_columns = ["job", "marital", "education", "default", "housing", "loan", "contact", 
                             "last_contact_month_of_year", "outcome_of_previous_campaign"]
        
        # Create numeric and categorical DataFrames
        print(f"📊 Dropping columns: {categorical_columns + ['cust_id', 'response_tag']}")
        numeric_df = df.drop(columns=categorical_columns + ["cust_id", "response_tag"])
        print(f"📊 Numeric columns after dropping: {numeric_df.columns.tolist()}")
        print(f"📊 Numeric DataFrame shape: {numeric_df.shape}")
        
        # Check if we have any numeric columns
        if numeric_df.empty or len(numeric_df.columns) == 0:
            print("⚠️ No numeric columns found after dropping specified columns")
            print("📊 Available columns in original DataFrame:")
            for col in df.columns:
                print(f"  - {col}: {df[col].dtype}")
            raise ValueError("No numeric columns available for analysis. Please check the column names and data types.")
        
        categorical_df = df[categorical_columns]
        print(f"📊 Categorical columns: {categorical_df.columns.tolist()}")
        print(f"📊 Categorical DataFrame shape: {categorical_df.shape}")

        # Create three sheets for different aspects of EDA
        # 1. Tables Sheet
        tables_sheet_name = f"{table_name}_EDA_Tables".upper()
        print(f"📄 Checking for existing sheet: {tables_sheet_name}")
        existing_sheets = [s.name for s in book.sheets]
        print(f"📄 Current sheets in workbook: {existing_sheets}")
        
        if tables_sheet_name in existing_sheets:
            try:
                print(f"🗑️ Attempting to delete existing sheet: {tables_sheet_name}")
                book.sheets[tables_sheet_name].delete()
                print(f"✅ Successfully deleted sheet: {tables_sheet_name}")
            except Exception as e:
                print(f"⚠️ Error deleting sheet '{tables_sheet_name}': {str(e)}")
                print(f"⚠️ Sheet deletion failed, attempting to continue...")
        
        try:
            tables_sht = book.sheets.add(name=tables_sheet_name, after=target_sheet)
            print(f"✅ Created new sheet: {tables_sheet_name}")
        except Exception as e:
            print(f"❌ Failed to create sheet '{tables_sheet_name}': {str(e)}")
            raise

        # 2. Correlation Matrix Sheet
        corr_sheet_name = f"{table_name}_Correlation_Matrix".upper()
        print(f"📄 Checking for existing sheet: {corr_sheet_name}")
        
        if corr_sheet_name in existing_sheets:
            try:
                print(f"🗑️ Attempting to delete existing sheet: {corr_sheet_name}")
                book.sheets[corr_sheet_name].delete()
                print(f"✅ Successfully deleted sheet: {corr_sheet_name}")
            except Exception as e:
                print(f"⚠️ Error deleting sheet '{corr_sheet_name}': {str(e)}")
                print(f"⚠️ Sheet deletion failed, attempting to continue...")
        
        try:
            corr_sht = book.sheets.add(name=corr_sheet_name, after=tables_sht)
            print(f"✅ Created new sheet: {corr_sheet_name}")
        except Exception as e:
            print(f"❌ Failed to create sheet '{corr_sheet_name}': {str(e)}")
            raise

        # 3. Charts Sheet
        charts_sheet_name = f"{table_name}_EDA_Charts".upper()
        print(f"📄 Checking for existing sheet: {charts_sheet_name}")
        
        if charts_sheet_name in existing_sheets:
            try:
                print(f"🗑️ Attempting to delete existing sheet: {charts_sheet_name}")
                book.sheets[charts_sheet_name].delete()
                print(f"✅ Successfully deleted sheet: {charts_sheet_name}")
            except Exception as e:
                print(f"⚠️ Error deleting sheet '{charts_sheet_name}': {str(e)}")
                print(f"⚠️ Sheet deletion failed, attempting to continue...")
        
        try:
            charts_sht = book.sheets.add(name=charts_sheet_name, after=corr_sht)
            print(f"✅ Created new sheet: {charts_sheet_name}")
        except Exception as e:
            print(f"❌ Failed to create sheet '{charts_sheet_name}': {str(e)}")
            raise

        # Step 1: Numeric Variables Analysis
        print("📊 Analyzing numeric variables...")
        
        # Calculate comprehensive statistics for numeric variables
        # Initialize stats_df with numeric columns
        stats_df = pd.DataFrame(index=['count', 'mean', 'std', 'min', '25%', '50%', '75%', 'max',
                                     'skewness', 'kurtosis', 'variance',
                                     '0%', '1%', '5%', '10%', '20%', '30%', '40%', '50%', '60%',
                                     '70%', '80%', '90%', '95%', '99%', '100%',
                                     'missing_count', 'missing_pct'],
                              columns=numeric_df.columns)
        
        # Basic statistics
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
        
        # Write numeric statistics to tables sheet
        tables_sht["A1"].value = "Numeric Variables Analysis"
        try:
            tables_sht["A1"].font.bold = True
            tables_sht["A1"].font.size = 14
        except Exception as e:
            print(f"⚠️ Could not format header: {e}")

        # Write and format numeric statistics table
        try:
            tables_sht["A2"].options(index=False).value = stats_df.reset_index().rename(columns={'index': 'Statistic'})
            stats_range = tables_sht["A2"].resize(len(stats_df) + 1, len(stats_df.columns) + 1)
            tables_sht.tables.add(stats_range)
            print("✅ Formatted numeric statistics as table")
        except Exception as table_err:
            print(f"⚠️ Could not format numeric statistics as table: {table_err}")
        
        # Step 2: Categorical Variables Analysis
        print("📊 Analyzing categorical variables...")
        
        # Calculate value counts and percentages for each categorical variable
        current_row = len(stats_df) + 5  # Add extra space after numeric stats
        
        for col in categorical_columns:
            value_counts = df[col].value_counts()
            value_pcts = df[col].value_counts(normalize=True) * 100
            
            # Create a DataFrame with counts and percentages
            cat_stats = pd.DataFrame({
                'Value': value_counts.index,
                'Count': value_counts.values,
                'Percentage': value_pcts.values.round(2)
            })
            
            # Write category header
            tables_sht[f"A{current_row}"].value = f"{col} Distribution"
            try:
                tables_sht[f"A{current_row}"].font.bold = True
                tables_sht[f"A{current_row}"].font.size = 12
            except Exception as e:
                print(f"⚠️ Could not format category header: {e}")
            
            # Write and format categorical table
            try:
                tables_sht[f"A{current_row + 1}"].options(index=False).value = cat_stats
                cat_range = tables_sht[f"A{current_row + 1}"].resize(len(cat_stats) + 1, len(cat_stats.columns))
                tables_sht.tables.add(cat_range)
                print(f"✅ Formatted {col} distribution as table")
            except Exception as table_err:
                print(f"⚠️ Could not format {col} distribution as table: {table_err}")
            
            current_row += len(cat_stats) + 3  # Add extra space between categories

        # Step 3: Correlation Analysis
        print("📊 Calculating correlation matrix...")
        corr_matrix = numeric_df.corr().round(2)
        corr_sht["A1"].value = "Correlation Matrix"
        try:
            corr_sht["A1"].font.bold = True
            corr_sht["A1"].font.size = 14
        except Exception as e:
            print(f"⚠️ Could not format correlation header: {e}")

        # Convert correlation matrix to DataFrame and format as table
        try:
            corr_df = corr_matrix.reset_index().rename(columns={'index': 'Variable'})
            corr_sht["A2"].options(index=False).value = corr_df
            corr_range = corr_sht["A2"].resize(len(corr_df) + 1, len(corr_df.columns))
            corr_sht.tables.add(corr_range)
            print("✅ Formatted correlation matrix as table")
        except Exception as table_err:
            print(f"⚠️ Could not format correlation matrix as table: {table_err}")

        # Step 4: Visualizations
        print("📊 Creating visualizations...")
        
        def plot_and_insert(fig, sheet, top_left_cell, name):
            try:
                temp_dir = tempfile.gettempdir()
                temp_path = os.path.join(temp_dir, f"{name}.png")
                fig.savefig(temp_path, dpi=150)
                print(f"🖼️ Saved plot '{name}' to {temp_path}")
                anchor_cell = sheet[top_left_cell]
                sheet.pictures.add(temp_path, name=name, update=True, anchor=anchor_cell, format="png")
                print(f"✅ Inserted plot '{name}' at {top_left_cell}")
            except Exception as e:
                print(f"❌ Failed to insert plot '{name}': {e}")
            finally:
                plt.close(fig)

        # 1. Correlation Heatmap
        plt.figure(figsize=(6, 4))  # Match scoring function chart size
        sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', center=0)
        plt.title("Correlation Heatmap")
        plot_and_insert(plt.gcf(), charts_sht, "A1", name="correlation_heatmap")

        # 2. Box Plots for Numeric Variables - Individual plots for each variable
        current_row = 20
        for col in numeric_df.columns:
            plt.figure(figsize=(6, 4))  # Match scoring function chart size
            sns.boxplot(y=numeric_df[col])
            plt.title(f"Box Plot of {col}")
            plt.tight_layout()
            plot_and_insert(plt.gcf(), charts_sht, f"A{current_row}", f"box_{col}")
            current_row += 25

        # 3. Distribution Plots for Numeric Variables
        for col in numeric_df.columns:
            plt.figure(figsize=(6, 4))  # Match scoring function chart size
            sns.histplot(data=numeric_df, x=col, kde=True)
            plt.title(f"Distribution of {col}")
            plt.tight_layout()
            plot_and_insert(plt.gcf(), charts_sht, f"A{current_row}", f"dist_{col}")
            current_row += 25

        # 4. Bar Plots for Categorical Variables
        for col in categorical_columns:
            plt.figure(figsize=(6, 4))  # Match scoring function chart size
            value_counts = df[col].value_counts()
            sns.barplot(x=value_counts.index, y=value_counts.values)
            plt.xticks(rotation=45, ha='right')
            plt.title(f"Distribution of {col}")
            plt.tight_layout()
            plot_and_insert(plt.gcf(), charts_sht, f"A{current_row}", f"bar_{col}")
            current_row += 25

        # Format the sheets
        try:
            # Format tables sheet
            tables_sht["A1"].font.bold = True
            tables_sht["A1"].font.size = 14
            
            # Format correlation matrix sheet
            corr_sht["A1"].font.bold = True
            corr_sht["A1"].font.size = 14
            
            # Format charts sheet
            charts_sht["A1"].font.bold = True
            charts_sht["A1"].font.size = 14
        except Exception as e:
            print(f"⚠️ Could not apply formatting: {e}")

        print("✅ EDA analysis completed successfully!")
        
    except Exception as e:
        print(f"❌ ERROR during EDA: {str(e)}")
        import traceback
        print(traceback.format_exc())
        return
    
    print("⭐⭐⭐ ENDING EDA ANALYSIS ⭐⭐⭐")