import pandas as pd
import numpy as np
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import re
from typing import Dict, List, Tuple, Any
import openpyxl
import datetime # <-- Added for timestamp

class ExcelFuzzyMapper:
    def __init__(self, excel1_path: str, excel2_path: str, mapping_excel_path: str):
        """
        Initialize the mapper with paths to three Excel files.
        
        Args:
            excel1_path: Path to first Excel file (with a1, a2, a3... columns)
            excel2_path: Path to second Excel file (with b1, b2, b3... columns)
            mapping_excel_path: Path to mapping Excel file
        """
        self.excel1_path = excel1_path
        self.excel2_path = excel2_path
        self.mapping_excel_path = mapping_excel_path
        
        # Load the Excel files
        self.df1 = pd.read_excel(excel1_path)
        self.df2 = pd.read_excel(excel2_path)
        self.mapping_df = pd.read_excel(mapping_excel_path)
        
        # Ensure column names are strings
        self.df1.columns = self.df1.columns.astype(str)
        self.df2.columns = self.df2.columns.astype(str)
    
    def parse_mapping_expression(self, expr: str) -> List[str]:
        """
        Parse mapping expressions like 'a10+a12' into ['a10', 'a12']
        
        Args:
            expr: Expression string (e.g., 'a10+a12' or 'a1')
            
        Returns:
            List of column names
        """
        # Remove spaces and split by '+'
        columns = expr.strip().replace(' ', '').split('+')
        return columns
    
    def get_concatenated_value(self, df: pd.DataFrame, columns: List[str], row_idx: int) -> str:
        """
        Get concatenated value from multiple columns for a specific row.
        
        Args:
            df: DataFrame to get values from
            columns: List of column names
            row_idx: Row index
            
        Returns:
            Concatenated string value
        """
        values = []
        for col in columns:
            if col in df.columns:
                val = df.loc[row_idx, col]
                if pd.notna(val):
                    values.append(str(val))
            else:
                print(f"Warning: Column '{col}' not found in DataFrame")
                break
        
        return ' '.join(values)
    
    def fuzzy_match_rows(self, str1: str, str2: str, threshold: int = 80) -> Tuple[bool, int]:
        """
        Perform case-insensitive fuzzy matching between two strings.
        
        Args:
            str1: First string
            str2: Second string
            threshold: Minimum similarity score (0-100)
            
        Returns:
            Tuple of (is_match, similarity_score)
        """
        # Convert to string, handle None values, and convert to lowercase
        str1 = str(str1).lower() if pd.notna(str1) else ''
        str2 = str(str2).lower() if pd.notna(str2) else ''
        
        # Calculate similarity score on lowercase strings
        score = fuzz.ratio(str1, str2)
        
        return score >= threshold, score
    
    def process_mappings(self, threshold: int = 80) -> pd.DataFrame:
        """
        Process all mappings and perform fuzzy matching.
        
        Args:
            threshold: Minimum similarity score for matching (0-100)
            
        Returns:
            DataFrame with fuzzy matching results
        """

        print ("DFs :\n")
        print (f"Excel1: {self.df1.head()}\n")
        print (f"Excel2: {self.df2.head()}\n")
              
        results = []
        
        # Get primary key mapping (assuming first row contains primary key mapping)
        if len(self.mapping_df) > 0:
            primary_key_source = str(self.mapping_df.iloc[0]['source_column'])
            primary_key_target = str(self.mapping_df.iloc[0]['target_column'])
            
            # Parse primary key columns
            pk_source_cols = self.parse_mapping_expression(primary_key_source)
            pk_target_cols = self.parse_mapping_expression(primary_key_target)
            
            print(f"Primary key mapping: {primary_key_source} -> {primary_key_target}")
            
            # Process each row in df1
            for idx1, row1 in self.df1.iterrows():
                # Get primary key value from df1
                pk_value1 = self.get_concatenated_value(self.df1, pk_source_cols, idx1)
                
                # Find matching row in df2 based on primary key
                best_match_idx = None
                best_match_score = 0
                
                for idx2, row2 in self.df2.iterrows():
                    pk_value2 = self.get_concatenated_value(self.df2, pk_target_cols, idx2)
                    # The fuzzy_match_rows function now handles case-insensitivity internally
                    is_match, score = self.fuzzy_match_rows(pk_value1, pk_value2, threshold)
                    
                    if is_match and score > best_match_score:
                        best_match_idx = idx2
                        best_match_score = score
                
                if best_match_idx is not None:
                    # Process all other column mappings for this row pair
                    row_result = {
                        'df1_row_index': idx1,
                        'df2_row_index': best_match_idx,
                        'primary_key_score': best_match_score,
                        'primary_key_value': pk_value1
                    }
                    
                    # Check all other mappings
                    for mapping_idx, mapping_row in self.mapping_df.iterrows():
                        if mapping_idx == 0:  # Skip primary key mapping
                            continue
                        
                        source_expr = str(mapping_row['source_column'])
                        target_expr = str(mapping_row['target_column'])
                        
                        source_cols = self.parse_mapping_expression(source_expr)
                        target_cols = self.parse_mapping_expression(target_expr)
                        
                        value1 = self.get_concatenated_value(self.df1, source_cols, idx1)
                        value2 = self.get_concatenated_value(self.df2, target_cols, best_match_idx)
                        
                        is_match, score = self.fuzzy_match_rows(value1, value2, threshold)
                        
                        row_result[f'mapping_{source_expr}_to_{target_expr}_score'] = score
                        row_result[f'mapping_{source_expr}_to_{target_expr}_match'] = is_match
                        row_result[f'value1_{source_expr}'] = value1
                        row_result[f'value2_{target_expr}'] = value2
                    
                    results.append(row_result)
        
        return pd.DataFrame(results)
    
    def generate_match_report(self, results_df: pd.DataFrame, output_path: str = 'fuzzy_match_report.xlsx'):
        """
        Generate a detailed match report in Excel format.
        
        Args:
            results_df: DataFrame with matching results
            output_path: Path to save the report
        """
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Write main results
            results_df.to_excel(writer, sheet_name='Match Results', index=False)
            
            # Create summary sheet
            summary_data = {
                'Total Rows in Excel1': [len(self.df1)],
                'Total Rows in Excel2': [len(self.df2)],
                'Total Matched Rows': [len(results_df)],
                'Match Rate': [f"{len(results_df)/len(self.df1)*100:.2f}%" if len(self.df1) > 0 else "0.00%"]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Write mapping configuration
            self.mapping_df.to_excel(writer, sheet_name='Mapping Configuration', index=False)
        
        print(f"Match report saved to: {output_path}")

def main():
    """
    Main function to run the fuzzy matching process.
    """
    # Example usage
    # --- IMPORTANT: REPLACE WITH YOUR ACTUAL FILE PATHS ---
    excel1_path = 'final_merged_0_june.xlsx'
    excel2_path = 'result_400_June5.xlsx'
    mapping_path = 'column_compare_crl_trial_09_xx.xlsx'
    
    # Create mapper instance
    mapper = ExcelFuzzyMapper(excel1_path, excel2_path, mapping_path)
    
    # Process mappings with 80% similarity threshold
    results = mapper.process_mappings(threshold=80)
    
    # --- FILENAME WITH TIMESTAMP ---
    # Get current time for the timestamp
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    # Create the dynamic output filename
    output_filename = f"fuzzy_match_results_x_{timestamp}.xlsx"
    
    # Generate report with the new dynamic filename
    mapper.generate_match_report(results, output_filename)
    
    # Display sample results
    print("\nSample Results:")
    print(results.head())
    
    # Display match statistics
    if not results.empty:
        print(f"\nMatch Statistics:")
        print(f"Total matched rows: {len(results)}")
        print(f"Average primary key match score: {results['primary_key_score'].mean():.2f}")

def create_sample_excels():
    """
    Creates sample Excel files to demonstrate how the script works.
    """
    # Create sample Excel 1
    df1_sample = pd.DataFrame({
        'a1': ['ID001', 'ID002', 'ID003', 'ID004'],
        'a2': ['John Doe', 'Jane Smith', 'Bob Johnson', 'Alice Brown'],
        'a3': ['New York', 'Los Angeles', 'Chicago', 'Houston'],
        'a4': ['Engineer', 'Manager', 'Analyst', 'Developer'],
        'a10': ['Dept', 'Dept', 'Dept', 'Dept'],
        'a12': ['IT', 'HR', 'Finance', 'IT']
    })
    
    # Create sample Excel 2
    df2_sample = pd.DataFrame({
        'b1': ['id001', 'id002', 'id003', 'id004'],
        'b2': ['Software', 'Human', 'Financial', 'Software'],
        'b5': ['John D.', 'jane smith', 'Bob J.', 'Alice B.'],
        'b6': ['Dept IT', 'Dept HR', 'DEPT FINANCE', 'Dept IT'],
        'b8': ['Engineer', 'Resources', 'Analyst', 'Developer'],
        'b9': ['ny', 'LA', 'CHI', 'HOU']
    })
    
    # Create sample mapping Excel
    mapping_sample = pd.DataFrame({
        'source_column': ['a1', 'a2', 'a3', 'a10+a12', 'a4'],
        'target_column': ['b1', 'b5', 'b9', 'b6', 'b2+b8'],
        'description': ['Primary Key', 'Name mapping', 'City mapping', 'Department concatenation', 'Role split mapping']
    })
    
    # Save sample files
    df1_sample.to_excel('excel1_sample.xlsx', index=False)
    df2_sample.to_excel('excel2_sample.xlsx', index=False)
    mapping_sample.to_excel('mapping_sample.xlsx', index=False)
    
    print("Sample Excel files (excel1_sample.xlsx, excel2_sample.xlsx, mapping_sample.xlsx) created!")
    print("\nTo use them, update the file paths in the main() function and run the script.")


if __name__ == "__main__":
    # Uncomment the line below to create sample files in your directory
    # create_sample_excels()
    
    # Run the main matching process
    main()
