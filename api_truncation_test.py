import os
import time
from datetime import datetime

def read_large_file(file_path):
    """Read a large source code file."""
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.read()

def write_log(message):
    """Write to log file with timestamp."""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open('truncation_test.log', 'a') as f:
        f.write(f'[{timestamp}] {message}\n')

def check_for_truncation(content):
    """
    Check if content appears to be truncated.
    Look for common indicators like missing closing braces,
    incomplete class definitions, etc.
    """
    # Count opening and closing braces
    open_braces = content.count('{')
    close_braces = content.count('}')
    
    # Check for common indicators of truncation
    truncation_indicators = [
        open_braces != close_braces,  # Mismatched braces
        content.strip().endswith('{'),  # Ends with open brace
        'namespace' in content and '}' not in content,  # Namespace without closing
        'class' in content and '}' not in content,  # Class without closing
        content.count('/*') != content.count('*/'),  # Unclosed comments
    ]
    
    return any(truncation_indicators)

def test_file_for_truncation(file_path, iterations=10, delay=5):
    """
    Test a file for truncation by repeatedly reading and analyzing it.
    
    Args:
        file_path: Path to the file to test
        iterations: Number of times to test
        delay: Delay between tests in seconds
    """
    write_log(f'Starting truncation test for {file_path}')
    write_log(f'File size: {os.path.getsize(file_path)} bytes')
    
    original_content = read_large_file(file_path)
    original_size = len(original_content)
    write_log(f'Original content length: {original_size} characters')
    
    for i in range(iterations):
        write_log(f'\nIteration {i+1}/{iterations}')
        
        try:
            # Read the file again
            current_content = read_large_file(file_path)
            current_size = len(current_content)
            
            # Check size difference
            size_diff = original_size - current_size
            write_log(f'Current content length: {current_size} characters')
            write_log(f'Size difference: {size_diff} characters')
            
            # Check for truncation
            if check_for_truncation(current_content):
                write_log('TRUNCATION DETECTED!')
                write_log('Truncation indicators found in content')
                
                # Log the last 500 characters to see where it cuts off
                write_log('Last 500 characters of truncated content:')
                write_log(current_content[-500:])
            else:
                write_log('No truncation detected')
            
            # Add a delay between iterations
            if i < iterations - 1:
                write_log(f'Waiting {delay} seconds before next iteration...')
                time.sleep(delay)
                
        except Exception as e:
            write_log(f'ERROR: {str(e)}')
    
    write_log('Truncation test completed')

def main():
    # List of files to test
    files_to_test = [
        'Services/PpsDeviceCommunicationService.cs',
        'Services/VisaCommunicationService.cs',
        'Pps Tests/Pps3150Afx/Pps3150AfxTestExecution.cs'
    ]
    
    # Create or clear log file
    with open('truncation_test.log', 'w') as f:
        f.write(f'Truncation Test Started at {datetime.now()}\n\n')
    
    # Test each file
    for file_path in files_to_test:
        if os.path.exists(file_path):
            write_log(f'\n{"="*50}')
            write_log(f'Testing file: {file_path}')
            write_log(f'{"="*50}\n')
            
            test_file_for_truncation(
                file_path,
                iterations=10,  # Test 10 times
                delay=5        # 5 second delay between tests
            )
        else:
            write_log(f'File not found: {file_path}')
    
    write_log('\nAll tests completed')

if __name__ == '__main__':
    main()
