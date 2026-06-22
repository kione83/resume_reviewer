"""
LinkedIn Employment Status Checker
Helps check student employment status by reading LinkedIn profiles
"""

import time
import csv
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

def setup_browser():
    """Set up Chrome browser with user profile"""
    print("🔧 Setting up browser...")
    print("⚠️  IMPORTANT: You must be logged into LinkedIn in Chrome!")
    print("   If you're not logged in, the script will fail.\n")
    
    options = webdriver.ChromeOptions()
    # Use your existing Chrome profile so you're already logged in
    # Uncomment and modify the path below if needed:
    # options.add_argument("user-data-dir=/path/to/your/chrome/profile")
    
    driver = webdriver.Chrome(options=options)
    return driver

def extract_employment_info(driver, linkedin_url):
    """Extract current employment from a LinkedIn profile"""
    try:
        print(f"   📍 Visiting: {linkedin_url}")
        driver.get(linkedin_url)
        
        # Wait for page to load
        time.sleep(3)
        
        # Try to find the experience section
        try:
            # Look for the current position (usually first in experience)
            # LinkedIn's structure changes frequently, so we try multiple selectors
            
            # Try method 1: Find experience section
            experience_section = driver.find_element(By.ID, "experience")
            parent = experience_section.find_element(By.XPATH, "..")
            
            # Find the first position listed
            position_elements = parent.find_elements(By.CSS_SELECTOR, "[class*='experience-item']")
            
            if position_elements:
                first_position = position_elements[0]
                
                # Try to extract title and company
                try:
                    title = first_position.find_element(By.CSS_SELECTOR, "[class*='profile-section-card__title']").text
                    company = first_position.find_element(By.CSS_SELECTOR, "[class*='profile-section-card__subtitle']").text
                    
                    print(f"   ✅ Found: {title} at {company}")
                    return {
                        'title': title.strip(),
                        'company': company.strip(),
                        'status': 'Employed'
                    }
                except:
                    pass
            
            # Method 2: Try alternative selectors
            try:
                title_elem = driver.find_element(By.CSS_SELECTOR, "div.text-body-medium")
                company_elem = driver.find_element(By.CSS_SELECTOR, "span.text-body-small")
                
                title = title_elem.text
                company = company_elem.text
                
                print(f"   ✅ Found: {title} at {company}")
                return {
                    'title': title.strip(),
                    'company': company.strip(),
                    'status': 'Employed'
                }
            except:
                pass
                
        except NoSuchElementException:
            print(f"   ⚠️  Could not find experience section")
            return {
                'title': 'Not found',
                'company': 'Not found',
                'status': 'Unknown'
            }
            
    except Exception as e:
        print(f"   ❌ Error: {str(e)}")
        return {
            'title': 'Error',
            'company': 'Error',
            'status': 'Error'
        }
    
    print(f"   ⚠️  Could not extract employment info")
    return {
        'title': 'Not found',
        'company': 'Not found', 
        'status': 'Unknown'
    }

def process_csv(input_file, output_file):
    """Process CSV with LinkedIn URLs and add employment info"""
    print("\n" + "="*70)
    print("LINKEDIN EMPLOYMENT CHECKER")
    print("="*70 + "\n")
    
    # Read input CSV
    try:
        df = pd.read_csv(input_file)
        print(f"✅ Loaded {len(df)} records from {input_file}\n")
    except Exception as e:
        print(f"❌ Could not read {input_file}: {str(e)}")
        return
    
    # Check for LinkedIn column
    linkedin_col = None
    for col in df.columns:
        if 'linkedin' in col.lower():
            linkedin_col = col
            break
    
    if not linkedin_col:
        print("❌ No LinkedIn column found in CSV")
        print(f"   Available columns: {', '.join(df.columns)}")
        return
    
    print(f"📋 Using column: '{linkedin_col}'\n")
    
    # Set up browser
    driver = setup_browser()
    
    # Add new columns for employment data
    df['Job Title'] = ''
    df['Company'] = ''
    df['Employment Status'] = ''
    
    try:
        # Process each row
        for idx, row in df.iterrows():
            linkedin_url = row[linkedin_col]
            
            if pd.isna(linkedin_url) or not linkedin_url:
                print(f"⏭️  Row {idx+1}: No LinkedIn URL")
                df.at[idx, 'Employment Status'] = 'No URL'
                continue
            
            print(f"\n👤 Row {idx+1}/{len(df)}: {row.get('first_name', '')} {row.get('last_name', '')}")
            
            # Extract employment info
            employment = extract_employment_info(driver, linkedin_url)
            
            # Update dataframe
            df.at[idx, 'Job Title'] = employment['title']
            df.at[idx, 'Company'] = employment['company']
            df.at[idx, 'Employment Status'] = employment['status']
            
            # Be polite to LinkedIn - wait between requests
            time.sleep(2)
        
        # Save results
        df.to_csv(output_file, index=False)
        print("\n" + "="*70)
        print(f"✅ Results saved to: {output_file}")
        print("="*70)
        
    finally:
        driver.quit()
        print("\n🔒 Browser closed")

def main():
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python linkedin_checker.py <input_csv> [output_csv]")
        print("\nExample: python linkedin_checker.py students.csv students_with_employment.csv")
        print("\nYour CSV must have a column with 'linkedin' in the name containing LinkedIn profile URLs")
        return
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else input_file.replace('.csv', '_employment.csv')
    
    # Check if Selenium is installed
    try:
        import selenium
    except ImportError:
        print("❌ Selenium not installed!")
        print("\nInstall it with: pip install selenium")
        print("\nYou also need Chrome and ChromeDriver installed.")
        return
    
    process_csv(input_file, output_file)

if __name__ == "__main__":
    main()
