"""
Test script for risk register standardization model
Tests against the 3 training examples
"""

import sys
import os
from pathlib import Path
import pandas as pd
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent))

from model import RiskRegisterStandardizer

def compare_outputs(generated_file, expected_file):
    """Compare generated output with expected output"""
    
    print(f"\nComparing:")
    print(f"  Generated: {generated_file}")
    print(f"  Expected:  {expected_file}")
    
    # Read both files
    df_gen = pd.read_excel(generated_file, sheet_name='Simplified Register')
    df_exp = pd.read_excel(expected_file, sheet_name='Simplified Register')
    
    print(f"\n  Generated shape: {df_gen.shape}")
    print(f"  Expected shape:  {df_exp.shape}")
    
    # Compare columns
    gen_cols = set(df_gen.columns)
    exp_cols = set(df_exp.columns)
    
    missing_cols = exp_cols - gen_cols
    extra_cols = gen_cols - exp_cols
    
    if missing_cols:
        print(f"  ⚠ Missing columns: {missing_cols}")
    if extra_cols:
        print(f"  ⚠ Extra columns: {extra_cols}")
    
    if not missing_cols and not extra_cols:
        print(f"  ✓ Column match!")
    
    # Check for mandatory fields
    mandatory = [
        'Risk ID', 'Risk Description', 'Project Stage', 'Project Category',
        'Risk Owner', 'Mitigating Action', 'Likelihood (1-10)', 
        'Impact (1-10)', 'Risk Priority (low, med, high)'
    ]
    
    missing_mandatory = [col for col in mandatory if col not in df_gen.columns]
    if missing_mandatory:
        print(f"  ✗ Missing mandatory columns: {missing_mandatory}")
    else:
        print(f"  ✓ All mandatory columns present")
    
    # Show sample comparison
    print(f"\n  Sample row comparison:")
    print(f"    Generated row 1: {df_gen.iloc[0].to_dict() if len(df_gen) > 0 else 'No data'}")
    print(f"    Expected row 1:  {df_exp.iloc[0].to_dict() if len(df_exp) > 0 else 'No data'}")
    
    return {
        'columns_match': not missing_cols and not extra_cols,
        'has_mandatory': not missing_mandatory,
        'row_count_match': df_gen.shape[0] == df_exp.shape[0]
    }


def main():
    """Run tests on training data"""
    
    # Check for API key
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        print("Error: ANTHROPIC_API_KEY environment variable not set")
        print("Set it with: export ANTHROPIC_API_KEY='your-key-here'")
        sys.exit(1)
    
    print("="*80)
    print("RISK REGISTER STANDARDIZATION MODEL - TESTING")
    print("="*80)
    
    # Set up paths
    uploads_dir = Path('/mnt/user-data/uploads')
    test_output_dir = Path('/home/claude/test_outputs')
    test_output_dir.mkdir(exist_ok=True)
    
    # Define test cases (training data)
    test_cases = [
        {
            'name': 'Test 1: IVC DOE',
            'input': uploads_dir / '1__IVC_DOE_R2__Input_.xlsx',
            'expected': uploads_dir / '1__IVC_DOE__Final_.xlsx',
            'output': test_output_dir / '1__IVC_DOE__Test_Output.xlsx'
        },
        {
            'name': 'Test 2: City of York Council',
            'input': uploads_dir / '2__City_of_York_Council__Input_.xlsx',
            'expected': uploads_dir / '2__City_of_York_Council__Final_.xlsx',
            'output': test_output_dir / '2__City_of_York_Council__Test_Output.xlsx'
        },
        {
            'name': 'Test 3: Digital Security IT',
            'input': uploads_dir / '3__Digital_Security_IT_Sample_Register__Input_.xlsx',
            'expected': uploads_dir / '3__Digital_Security_IT_Sample_Register__Final_.xlsx',
            'output': test_output_dir / '3__Digital_Security_IT__Test_Output.xlsx'
        }
    ]
    
    # Initialize standardizer
    standardizer = RiskRegisterStandardizer(api_key=api_key)
    
    results = []
    
    # Run tests
    for test in test_cases:
        print(f"\n{'='*80}")
        print(f"{test['name']}")
        print(f"{'='*80}")
        
        try:
            # Process file
            print(f"Processing: {test['input'].name}")
            standardizer.process_file(str(test['input']), str(test['output']))
            
            # Compare outputs
            comparison = compare_outputs(test['output'], test['expected'])
            
            results.append({
                'test': test['name'],
                'status': 'PASS' if all(comparison.values()) else 'PARTIAL',
                'details': comparison
            })
            
        except Exception as e:
            print(f"✗ ERROR: {e}")
            import traceback
            traceback.print_exc()
            results.append({
                'test': test['name'],
                'status': 'FAIL',
                'error': str(e)
            })
    
    # Summary
    print(f"\n{'='*80}")
    print("TEST SUMMARY")
    print(f"{'='*80}")
    
    for result in results:
        status_symbol = '✓' if result['status'] == 'PASS' else ('⚠' if result['status'] == 'PARTIAL' else '✗')
        print(f"{status_symbol} {result['test']}: {result['status']}")
        if 'details' in result:
            details = result['details']
            print(f"    Columns match: {details['columns_match']}")
            print(f"    Has mandatory: {details['has_mandatory']}")
            print(f"    Row count match: {details['row_count_match']}")
    
    print(f"\nTest outputs saved to: {test_output_dir}")


if __name__ == '__main__':
    main()
