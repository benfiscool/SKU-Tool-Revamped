#!/usr/bin/env python3
"""
Test script to verify backup file validation improvements
"""
import os
import json
import tempfile

def test_backup_validation():
    """Test the enhanced backup validation logic"""
    
    # Create a temporary directory structure
    with tempfile.TemporaryDirectory() as temp_dir:
        data_dir = os.path.join(temp_dir, 'data')
        os.makedirs(data_dir)
        
        # Test cases for different file conditions
        test_files = {
            'valid_sku_database_2025-07-16_15-40-22.json': {'test': 'data', 'items': [1, 2, 3]},
            'empty_sku_database_2025-07-15_10-30-15.json': {},
            'corrupted_sku_database_2025-07-14_09-20-10.json': None,  # Will create empty file
            'valid_cost_db_2025-07-16_15-40-22.json': {'costs': {'item1': 10.50}},
            'empty_cost_db_2025-07-15_10-30-15.json': {},
            'sku_database_temp.json': {'temp': 'file'},  # Should be excluded
            'cost_db_temp.json': {'temp': 'cost'},       # Should be excluded
        }
        
        # Create test files
        for filename, content in test_files.items():
            file_path = os.path.join(data_dir, filename)
            if content is None:
                # Create empty file for corruption test
                open(file_path, 'w').close()
            else:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(content, f)
        
        print("Created test files:")
        for filename in sorted(os.listdir(data_dir)):
            file_path = os.path.join(data_dir, filename)
            size = os.path.getsize(file_path)
            print(f"  {filename}: {size} bytes")
        
        # Test the validation logic
        print("\nTesting validation logic:")
        print(f"Files in directory: {sorted(os.listdir(data_dir))}")
        
        # Test SKU database validation
        sku_files = []
        for filename in os.listdir(data_dir):
            print(f"  Checking file: {filename}")
            if filename.startswith('sku_database_') and filename.endswith('.json'):
                print(f"    Matches SKU pattern")
                if '_temp' in filename or filename.endswith('_temp.json'):
                    print(f"    Excluded temp file: {filename}")
                    continue
                try:
                    timestamp_part = filename.replace('sku_database_', '').replace('.json', '')
                    if timestamp_part and timestamp_part != 'temp':
                        sku_files.append((filename, timestamp_part))
                        print(f"    Found SKU file: {filename} (timestamp: {timestamp_part})")
                except:
                    print(f"    Failed to parse timestamp for: {filename}")
                    continue
        
        sku_files.sort(key=lambda x: x[1], reverse=True)
        latest_sku_db = None
        
        for sku_filename, _ in sku_files:
            test_path = os.path.join(data_dir, sku_filename)
            print(f"  Testing SKU file: {sku_filename}")
            try:
                # Check if file exists and has content
                if not os.path.exists(test_path):
                    print(f"    ❌ File doesn't exist")
                    continue
                if os.path.getsize(test_path) == 0:
                    print(f"    ❌ File is empty (0 bytes)")
                    continue
                
                with open(test_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # Verify the data is not empty
                    if data:
                        print(f"    ✅ Valid data found")
                        latest_sku_db = test_path
                        break
                    else:
                        print(f"    ❌ File contains empty data")
            except json.JSONDecodeError as e:
                print(f"    ❌ JSON decode error: {e}")
                continue
            except UnicodeDecodeError as e:
                print(f"    ❌ Unicode decode error: {e}")
                continue
            except OSError as e:
                print(f"    ❌ OS error: {e}")
                continue
        
        # Test cost database validation
        cost_files = []
        for filename in os.listdir(data_dir):
            if filename.startswith('cost_db_') and filename.endswith('.json'):
                if '_temp' in filename or filename.endswith('_temp.json'):
                    print(f"  Excluded temp cost file: {filename}")
                    continue
                try:
                    timestamp_part = filename.replace('cost_db_', '').replace('.json', '')
                    if timestamp_part and timestamp_part != 'temp':
                        cost_files.append((filename, timestamp_part))
                        print(f"  Found cost file: {filename} (timestamp: {timestamp_part})")
                except:
                    print(f"  Failed to parse timestamp for: {filename}")
                    continue
        
        cost_files.sort(key=lambda x: x[1], reverse=True)
        cost_db_path = None
        
        for cost_filename, _ in cost_files:
            test_path = os.path.join(data_dir, cost_filename)
            print(f"  Testing cost file: {cost_filename}")
            try:
                # Check if file exists and has content
                if not os.path.exists(test_path):
                    print(f"    ❌ File doesn't exist")
                    continue
                if os.path.getsize(test_path) == 0:
                    print(f"    ❌ File is empty (0 bytes)")
                    continue
                
                with open(test_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # Verify the data is not empty
                    if data:
                        print(f"    ✅ Valid data found")
                        cost_db_path = test_path
                        break
                    else:
                        print(f"    ❌ File contains empty data")
            except json.JSONDecodeError as e:
                print(f"    ❌ JSON decode error: {e}")
                continue
            except UnicodeDecodeError as e:
                print(f"    ❌ Unicode decode error: {e}")
                continue
            except OSError as e:
                print(f"    ❌ OS error: {e}")
                continue
        
        print(f"\nResults:")
        print(f"  Selected SKU database: {os.path.basename(latest_sku_db) if latest_sku_db else 'None'}")
        print(f"  Selected cost database: {os.path.basename(cost_db_path) if cost_db_path else 'None'}")
        
        # Verify our logic works correctly
        if latest_sku_db and 'valid_sku_database_2025-07-16_15-40-22.json' in latest_sku_db:
            print("  ✅ SKU validation working correctly - selected valid file")
        else:
            print("  ❌ SKU validation failed - selected wrong file or no file")
            
        if cost_db_path and 'valid_cost_db_2025-07-16_15-40-22.json' in cost_db_path:
            print("  ✅ Cost validation working correctly - selected valid file")
        else:
            print("  ❌ Cost validation failed - selected wrong file or no file")

if __name__ == '__main__':
    test_backup_validation()
