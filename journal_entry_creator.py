#!/usr/bin/env python3
"""
Journal Entry ID Creator Script

This script reads journal lines from an Excel file and groups them into balanced
journal entries where:
1. Each journal entry has debits = credits (balanced)
2. All journal lines in an entry have the same posted date
3. Journal lines are grouped by combinations of posted date + optional fields

Algorithm:
1. Read all journal lines from Excel
2. Group lines by Posted Date
3. Within each date group, try different combinations of optional fields
4. For each combination, check if the group balances (debits = credits)
5. Assign unique Journal Entry IDs to balanced groups
6. Output results with Journal ID column added
"""

import pandas as pd
import numpy as np
from itertools import combinations, product
from collections import defaultdict
import uuid
import argparse
from pathlib import Path

class JournalEntryCreator:
    def __init__(self):
        self.journal_lines = None
        self.grouped_entries = {}
        self.unassigned_lines = []
        
    def load_data(self, file_path):
        """Load journal lines from Excel file"""
        try:
            # Read the entire file to analyze structure
            df_all = pd.read_excel(file_path, header=None)
            
            if len(df_all) == 0:
                print("No data found in Excel file.")
                return False
            
            # Check the structure - look for the actual field names
            first_row = df_all.iloc[0].values
            second_row = df_all.iloc[1].values if len(df_all) > 1 else None
            
            # Check different possible structures
            if second_row is not None and isinstance(second_row[0], str) and 'Posted Date' in str(second_row[0]):
                print("Detected field names in second row (row 1)")
                template_row = second_row
                df = df_all.iloc[2:].copy()  # Data starts from row 3 (index 2)
            elif isinstance(first_row[0], str) and 'Posted Date' in str(first_row[0]):
                print("Detected template headers in first row")
                template_row = first_row
                df = df_all.iloc[1:].copy()  # Data starts from row 2
            else:
                print("No template headers detected, assuming data starts from first row")
                # Use default column names
                template_row = ['Posted Date', 'Account ID', 'Debit Amount', 'Credit Amount'] + \
                              [f'Optional_{i}' for i in range(len(df_all.columns) - 4)]
                df = df_all.copy()
            
            if len(df) == 0:
                print("No data found in Excel file. Please add journal line data.")
                return False
            
            # Set column names from template
            df.columns = template_row[:len(df.columns)]  # Handle case where data has fewer columns
            
            # Remove completely empty rows
            df = df.dropna(how='all')
            
            # Ensure required columns exist and have data
            required_cols = ['Posted Date', 'Account ID', 'Debit Amount', 'Credit Amount']
            for col in required_cols:
                if col not in df.columns:
                    print(f"Error: Required column '{col}' not found in Excel file")
                    return False
            
            # Convert date column to datetime
            df['Posted Date'] = pd.to_datetime(df['Posted Date'])
            
            # Convert amount columns to numeric, filling NaN with 0
            df['Debit Amount'] = pd.to_numeric(df['Debit Amount'], errors='coerce').fillna(0)
            df['Credit Amount'] = pd.to_numeric(df['Credit Amount'], errors='coerce').fillna(0)
            
            # Validate that each line has either debit OR credit, not both
            both_amounts = (df['Debit Amount'] != 0) & (df['Credit Amount'] != 0)
            if both_amounts.any():
                print("Warning: Some lines have both debit and credit amounts. This may not be standard.")
            
            # Add row index for tracking
            df['_row_index'] = range(len(df))
            
            self.journal_lines = df
            print(f"Loaded {len(df)} journal lines from {file_path}")
            return True
            
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            return False
    
    def get_optional_fields(self):
        """Get list of optional field column names that have data"""
        if self.journal_lines is None:
            return []
        
        # Find columns that are not required fields and have some non-null data
        required_cols = ['Posted Date', 'Account ID', 'Debit Amount', 'Credit Amount', '_row_index']
        optional_cols = []
        
        for col in self.journal_lines.columns:
            if col not in required_cols:
                # Check if column has any non-null, non-empty data
                try:
                    non_empty = self.journal_lines[col].notna() & \
                               (self.journal_lines[col] != '') & \
                               (self.journal_lines[col] != '[INSERT FIELD NAME]')
                    if non_empty.any():
                        optional_cols.append(col)
                except:
                    # Skip columns that cause issues
                    continue
        
        return optional_cols
    
    def check_balance(self, group_df):
        """Check if a group of journal lines is balanced (debits = credits)"""
        total_debits = group_df['Debit Amount'].sum()
        total_credits = group_df['Credit Amount'].sum()
        
        # Use small epsilon for floating point comparison
        return abs(total_debits - total_credits) < 0.01
    
    def generate_grouping_combinations(self, optional_fields, max_fields=3):
        """Generate combinations of fields to try for grouping - prioritize more specific groupings first"""
        combinations_to_try = []
        
        # Start with most specific combinations (more fields = smaller groups)
        # Work backwards from max fields to 1 field, then date-only last
        for r in range(min(len(optional_fields), max_fields), 0, -1):
            for combo in combinations(optional_fields, r):
                combinations_to_try.append(['Posted Date'] + list(combo))
        
        # Add date-only grouping as the last resort
        combinations_to_try.append(['Posted Date'])
        
        return combinations_to_try
    
    def create_journal_entries(self, max_optional_fields=5):
        """Create journal entries by grouping lines that balance"""
        if self.journal_lines is None:
            print("No data loaded. Please load data first.")
            return False
        
        print("Creating journal entries...")
        
        # Get optional fields that have data
        optional_fields = self.get_optional_fields()
        print(f"Found optional fields with data: {optional_fields}")
        
        # Generate combinations of fields to try for grouping
        field_combinations = self.generate_grouping_combinations(optional_fields, max_optional_fields)
        print(f"Will try {len(field_combinations)} different grouping combinations")
        
        # Track which lines have been assigned to journal entries
        assigned_lines = set()
        journal_entry_id = 1
        
        # Try each combination of grouping fields
        for fields in field_combinations:
            print(f"\nTrying grouping by: {fields}")
            
            # Get unassigned lines
            unassigned_df = self.journal_lines[~self.journal_lines['_row_index'].isin(assigned_lines)].copy()
            
            if len(unassigned_df) == 0:
                print("All lines have been assigned to journal entries.")
                break
            
            # Group by the current field combination
            valid_fields = [col for col in fields if col in unassigned_df.columns]
            grouped = unassigned_df.groupby(valid_fields)
            
            balanced_groups = 0
            groups_processed = 0
            
            for group_key, group_df in grouped:
                groups_processed += 1
                if self.check_balance(group_df):
                    # Create journal entry ID
                    je_id = f"JE{journal_entry_id:04d}"
                    
                    # Store the journal entry
                    self.grouped_entries[je_id] = {
                        'lines': group_df.copy(),
                        'grouping_fields': fields,
                        'group_key': group_key,
                        'total_debits': group_df['Debit Amount'].sum(),
                        'total_credits': group_df['Credit Amount'].sum()
                    }
                    
                    # Mark these lines as assigned
                    assigned_lines.update(group_df['_row_index'].tolist())
                    
                    journal_entry_id += 1
                    balanced_groups += 1
            
            print(f"Found {balanced_groups} balanced journal entries from {groups_processed} groups with this grouping")
            
            # If this grouping level found balanced entries, continue with this level
            # to find more specific entries before moving to broader groupings
            if balanced_groups > 0:
                print(f"  -> Successfully created {balanced_groups} specific journal entries")
        
        # Track unassigned lines
        self.unassigned_lines = self.journal_lines[~self.journal_lines['_row_index'].isin(assigned_lines)].copy()
        
        print(f"\nSummary:")
        print(f"Total journal entries created: {len(self.grouped_entries)}")
        print(f"Total lines assigned: {len(assigned_lines)}")
        print(f"Total lines unassigned: {len(self.unassigned_lines)}")
        
        return True
    
    def generate_output(self, input_file_path, output_file_path=None):
        """Generate output Excel file with Journal ID column"""
        if self.journal_lines is None:
            print("No data to output")
            return False
        
        # Create output dataframe
        output_df = self.journal_lines.copy()
        
        # Add Journal ID column
        output_df['Journal ID'] = ''
        
        # Assign Journal IDs
        for je_id, entry_data in self.grouped_entries.items():
            row_indices = entry_data['lines']['_row_index'].tolist()
            output_df.loc[output_df['_row_index'].isin(row_indices), 'Journal ID'] = je_id
        
        # Remove the temporary row index column
        output_df = output_df.drop('_row_index', axis=1)
        
        # Reorder columns to put Journal ID before Posted Date
        cols = list(output_df.columns)
        cols.remove('Journal ID')
        posted_date_idx = cols.index('Posted Date')
        cols.insert(posted_date_idx, 'Journal ID')
        output_df = output_df[cols]
        
        # Generate output file path if not provided
        if output_file_path is None:
            input_path = Path(input_file_path)
            output_file_path = input_path.parent / f"{input_path.stem}_with_journal_ids{input_path.suffix}"
        
        try:
            # Write to Excel
            output_df.to_excel(output_file_path, index=False)
            print(f"Output written to: {output_file_path}")
            
            # Print summary report
            self.print_summary_report()
            
            return True
            
        except Exception as e:
            print(f"Error writing output file: {e}")
            return False
    
    def print_summary_report(self):
        """Print a summary report of journal entries created"""
        print("\n" + "="*60)
        print("JOURNAL ENTRY SUMMARY REPORT")
        print("="*60)
        
        for je_id, entry_data in sorted(self.grouped_entries.items()):
            lines = entry_data['lines']
            print(f"\n{je_id}:")
            print(f"  Date: {lines['Posted Date'].iloc[0].strftime('%Y-%m-%d')}")
            print(f"  Lines: {len(lines)}")
            print(f"  Total Debits: ${entry_data['total_debits']:,.2f}")
            print(f"  Total Credits: ${entry_data['total_credits']:,.2f}")
            print(f"  Grouped by: {entry_data['grouping_fields']}")
            if len(entry_data['grouping_fields']) > 1:
                print(f"  Group values: {entry_data['group_key']}")
        
        if len(self.unassigned_lines) > 0:
            print(f"\nUNASSIGNED LINES ({len(self.unassigned_lines)}):")
            for _, line in self.unassigned_lines.iterrows():
                print(f"  Date: {line['Posted Date'].strftime('%Y-%m-%d')}, "
                      f"Account: {line['Account ID']}, "
                      f"Debit: ${line['Debit Amount']:.2f}, "
                      f"Credit: ${line['Credit Amount']:.2f}")

def main():
    parser = argparse.ArgumentParser(description='Create balanced journal entries from Excel file')
    parser.add_argument('input_file', help='Path to input Excel file')
    parser.add_argument('-o', '--output', help='Path to output Excel file (optional)')
    parser.add_argument('--max-fields', type=int, default=5, 
                       help='Maximum number of optional fields to use for grouping (default: 5)')
    
    args = parser.parse_args()
    
    # Create journal entry creator
    creator = JournalEntryCreator()
    
    # Load data
    if not creator.load_data(args.input_file):
        return 1
    
    # Create journal entries
    if not creator.create_journal_entries(args.max_fields):
        return 1
    
    # Generate output
    if not creator.generate_output(args.input_file, args.output):
        return 1
    
    return 0

if __name__ == "__main__":
    exit(main())
