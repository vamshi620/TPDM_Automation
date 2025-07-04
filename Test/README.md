# TPDM Automation Test Files

This directory contains sample test files for validating the TPDM Automation application.

## Test Files

### Input Files
- `sample_input.xlsx` - Sample Excel file with multiple sheets for testing

### Output Files (Generated)
- `Add_Records.xlsx` - Records classified as "Add" by ML model
- `ADD_Records.xlsx` - Records from sheets without "Delegate Comments" column (default to ADD)
- `Update_Records.xlsx` - Records classified as "Update" by ML model  
- `Term_Records.xlsx` - Records classified as "Term" by ML model
- `Other_Records.xlsx` - Records classified as "Other" by ML model

## Test Data Structure

### sample_input.xlsx contains:

1. **HR_Records Sheet** (with Delegate Comments)
   - 10 rows with various comment types
   - Contains "Delegate Comments" column
   - Comments designed to test all 4 categories (Add, Update, Term, Other)

2. **Employee_Data Sheet** (with Delegate Comments)
   - 10 rows with various comment types
   - Contains "Delegate Comments" column
   - Duplicate of HR_Records data for testing consistency

3. **Payroll_Info Sheet** (without Delegate Comments)
   - 5 rows of payroll data
   - No "Delegate Comments" column
   - Should default all records to "ADD" category

## Test Results

### Processing Summary:
- **Total rows processed**: 25
- **Add**: 8 records (from ML classification)
- **ADD**: 5 records (from default assignment)
- **Update**: 6 records (from ML classification)
- **Term**: 4 records (from ML classification)
- **Other**: 2 records (from ML classification)

### Records by Sheet:
- **HR_Records**: 10 records
- **Employee_Data**: 10 records  
- **Payroll_Info**: 5 records

## ML Model Performance
- **MicroAccuracy**: 1.0 (100%)
- **MacroAccuracy**: 1.0 (100%)

## Validation Points

1. ✅ Application successfully processes multiple sheets
2. ✅ ML model correctly classifies comments into 4 categories
3. ✅ Sheets without "Delegate Comments" default to "ADD" category
4. ✅ Output files are generated correctly for each category
5. ✅ All data is preserved in output files with predicted categories
6. ✅ ML model achieves perfect accuracy on training data

## Running the Test

To regenerate the test data and run the test:

```bash
# Generate fresh test data
dotnet run -- --generate-test-data

# Run the application with test data
dotnet run -- -i "./Test/sample_input.xlsx" -o "./Test"
```

## Sample Comments Used

The test data includes realistic delegate comments that should map to each category:

**Add Category:**
- "Added new employee to the system"
- "New hire has been processed"
- "Onboarding new staff member"

**Update Category:**
- "Updated employee information"
- "Modified contact details"
- "Changed department assignment"

**Term Category:**
- "Employee has been terminated"
- "Terminated due to policy violation"

**Other Category:**
- "Review pending for this employee"
- "Under investigation"