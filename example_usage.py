#!/usr/bin/env python3
"""
Example usage of the Journal Entry ID Creator
"""
import pandas as pd
from datetime import datetime
from journal_entry_creator import JournalEntryCreator

def create_example_data():
    """Create a simple example Excel file for demonstration"""
    
    # Create sample data
    data = [
        # Template row
        ['Posted Date', 'Account ID', 'Debit Amount', 'Credit Amount', 'Description', 'Reference'],
        
        # Journal Entry 1: Simple sale
        [datetime(2024, 1, 15), '1000', 1000.00, 0, 'Product Sales', 'INV-001'],
        [datetime(2024, 1, 15), '4000', 0, 1000.00, 'Product Sales', 'INV-001'],
        
        # Journal Entry 2: Expense
        [datetime(2024, 1, 16), '6000', 500.00, 0, 'Office Rent', 'RENT-001'],
        [datetime(2024, 1, 16), '1000', 0, 500.00, 'Office Rent', 'RENT-001'],
        
        # Journal Entry 3: Multi-line entry
        [datetime(2024, 1, 17), '6100', 200.00, 0, 'Travel Expense', 'TRV-001'],
        [datetime(2024, 1, 17), '6200', 150.00, 0, 'Meal Expense', 'TRV-001'],
        [datetime(2024, 1, 17), '1000', 0, 350.00, 'Travel Reimbursement', 'TRV-001'],
        
        # Unbalanced line (will remain unassigned)
        [datetime(2024, 1, 18), '2000', 100.00, 0, 'Unbalanced Entry', 'UNB-001'],
    ]
    
    # Create DataFrame and save to Excel
    df = pd.DataFrame(data)
    df.to_excel('example_journal_data.xlsx', index=False, header=False)
    
    print("Created example_journal_data.xlsx")
    return 'example_journal_data.xlsx'

def run_example():
    """Run the journal entry creator on example data"""
    
    # Create example data
    input_file = create_example_data()
    
    # Create and run journal entry creator
    creator = JournalEntryCreator()
    
    print(f"\nProcessing {input_file}...")
    
    # Load data
    if creator.load_data(input_file):
        # Create journal entries
        if creator.create_journal_entries():
            # Generate output
            output_file = 'example_output_with_journal_ids.xlsx'
            if creator.generate_output(input_file, output_file):
                print(f"\n✅ Success! Check {output_file} for results.")
            else:
                print("❌ Failed to generate output file")
        else:
            print("❌ Failed to create journal entries")
    else:
        print("❌ Failed to load input file")

if __name__ == "__main__":
    run_example()
