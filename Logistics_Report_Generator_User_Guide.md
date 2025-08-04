# Logistics Report Generator - User Guide

## What is the Logistics Report Generator?

The Logistics Report Generator is a web-based tool that helps you turn messy Excel files into professional, standardized reports. Think of it as a smart assistant that takes your raw logistics data and transforms it into clean, organized reports that you can use for business purposes.

### What Does It Do?

- **Takes Excel files** with logistics data (like packing lists, inventory sheets)
- **Cleans up the data** automatically
- **Creates professional reports** using a standard template
- **Checks for errors** and missing information
- **Lets you download** the finished reports

### Why Use It?

- **Save Time**: No more manual copying and pasting data
- **Reduce Errors**: Automatic validation catches mistakes
- **Look Professional**: All reports have consistent formatting
- **Handle Complex Files**: Works with files that have multiple tables
- **Easy to Use**: Just drag and drop your files

---

## Getting Started

### What You Need

Before you start, make sure you have:
- A computer with internet access
- Excel files (.xlsx or .xlsm format) that you want to process
- The application installed and running

### Quick Setup (5 Minutes)

1. **Install the Application**
   - Follow the installation instructions provided by your IT team
   - This is usually a one-time setup

2. **Start the Application**
   - Open the application in your web browser
   - You should see the main screen with a file upload area

3. **You're Ready!**
   - The application is now ready to process your files

---

## How to Use the Application

### Step 1: Upload Your File

1. **Find your Excel file** on your computer
2. **Drag and drop** the file into the upload area, OR
3. **Click "Choose File"** and select your file from the file browser

**What files work?**
- Excel files (.xlsx or .xlsm)
- Files that contain logistics data (packing lists, inventory sheets)
- Files with tables that have headers like "CASE NOS", "COLOR", "SIZE", etc.

### Step 2: Review the Data

After uploading, the application will:
- **Show you what data it found** in your file
- **Display any errors or warnings** it found
- **Let you preview** the data before processing

**What to look for:**
- Check that all your important data was found
- Review any error messages
- Make sure the data looks correct

### Step 3: Generate Reports

1. **Click "Generate Reports"** button
2. **Wait for processing** (this usually takes a few seconds)
3. **Review the results** when complete

### Step 4: Download Your Reports

You can download your reports in several ways:
- **Individual reports**: Download each report separately
- **Combined file**: Download all reports in one Excel file
- **Preview first**: Look at the reports before downloading

---

## Understanding Your Data

### What Data Does the Application Look For?

The application recognizes these common data fields:

| Field Name | What It Is | Required? |
|------------|------------|-----------|
| CASE NOS | Carton or case numbers | Yes |
| SA4 PO NO# | Purchase order number | Yes |
| SAP STYLE NO | Product style identifier | Yes |
| STYLE NAME # | Product name or model | Yes |
| COLOR | Product color | Yes |
| Size | Product size (S, M, L, etc.) | Yes |
| Total QTY | Total quantity | Yes |
| CARTON | Carton information | Yes |
| QTY / CARTON | Quantity per carton | Yes |

### Size Categories

The application understands these standard sizes:
- **OS** (One Size)
- **XS** (Extra Small)
- **S** (Small)
- **M** (Medium)
- **L** (Large)
- **XL** (Extra Large)
- **XXL** (Double Extra Large)

---

## What Happens to Your Data

### Data Processing Steps

1. **File Reading**: The application reads your Excel file
2. **Table Detection**: It finds tables in your file (looks for "CASE NOS" headers)
3. **Data Extraction**: It pulls out the important information
4. **Data Cleaning**: It fixes common issues like missing carton numbers
5. **Validation**: It checks for errors and missing data
6. **Report Generation**: It creates professional reports using a template
7. **Formatting**: It applies consistent styling and layout

### Data Transformations

**Carton Number Filling**: If some rows are missing carton numbers, the application fills them in from the row above.

**Size Breakdown**: The application separates quantities by size (S, M, L, etc.) and calculates totals.

**Color Summaries**: It groups data by color and creates summary totals.

**Weight Calculations**: It calculates total weights and volumes (CBM) where possible.

---

## Understanding the Reports

### What's in Your Generated Reports?

Each report contains:

1. **Header Section**
   - Company information
   - Purchase order details
   - Report date and reference numbers

2. **Main Data Table**
   - All your product information
   - Quantities by size
   - Carton and weight information

3. **Summary Section**
   - Total quantities
   - Total weights
   - Overall statistics

4. **Color Breakdown**
   - Summary by color
   - Size breakdowns
   - Totals for each color

### Report Formatting

- **Professional appearance** with borders and styling
- **Consistent layout** across all reports
- **Print-ready format** for business use
- **Clear organization** of information

---

## Common Scenarios

### Scenario 1: Processing a Packing List

**What you have**: An Excel file from a supplier with packing list data
**What you get**: Professional reports ready for your inventory system

**Steps**:
1. Upload the supplier's Excel file
2. Review the extracted data
3. Generate reports
4. Download and use in your business

### Scenario 2: Multiple Orders in One File

**What you have**: One Excel file with data for multiple purchase orders
**What you get**: Separate reports for each order

**Steps**:
1. Upload the file
2. The application automatically detects multiple orders
3. Generate reports (creates one report per order)
4. Download individual reports or combined file

### Scenario 3: Data Validation

**What you have**: Excel file that might have errors or missing data
**What you get**: Clean data with error reports

**Steps**:
1. Upload the file
2. Review validation results
3. Fix any issues in your original file if needed
4. Re-upload and generate reports

---

## Troubleshooting Common Issues

### "File Won't Upload"

**Possible causes**:
- File is too large (try breaking it into smaller files)
- Wrong file format (use .xlsx or .xlsm files only)
- File is corrupted

**Solutions**:
- Check file size (should be under 50MB)
- Make sure it's an Excel file
- Try opening the file in Excel first to check for corruption

### "Data Not Found"

**Possible causes**:
- File doesn't have the expected headers
- Data is in a different format than expected
- Tables are not clearly defined

**Solutions**:
- Make sure your file has headers like "CASE NOS", "COLOR", etc.
- Check that data is organized in tables
- Contact support if your file format is different

### "Reports Look Wrong"

**Possible causes**:
- Data was not extracted correctly
- Template doesn't match your needs
- Processing errors occurred

**Solutions**:
- Review the data preview before generating reports
- Check the validation results for errors
- Contact support for template adjustments

### "Application Won't Start"

**Possible causes**:
- Installation issues
- Port conflicts
- Missing dependencies

**Solutions**:
- Contact your IT team for installation help
- Make sure no other applications are using the same ports
- Check that all required software is installed

---

## Best Practices

### Preparing Your Files

1. **Use consistent headers**: Make sure your Excel files have clear, consistent column headers
2. **Organize data in tables**: Keep your data in organized tables rather than scattered across the sheet
3. **Check for errors**: Review your data for obvious errors before uploading
4. **Use standard formats**: Stick to common date and number formats

### Working with the Application

1. **Start small**: Test with a small file first to understand the process
2. **Review results**: Always check the data preview before generating reports
3. **Save backups**: Keep copies of your original files
4. **Validate output**: Review generated reports for accuracy

### File Organization

1. **Clear naming**: Use descriptive names for your files
2. **Consistent structure**: Keep similar files organized in the same way
3. **Regular processing**: Process files regularly rather than in large batches
4. **Archive processed files**: Keep track of which files have been processed

---

## Getting Help

### When to Contact Support

Contact your IT team or support if you experience:
- **Technical errors** that prevent the application from working
- **Data processing issues** that can't be resolved through troubleshooting
- **Template customization** needs
- **New file format** requirements
- **Performance problems** with large files

### What Information to Provide

When contacting support, include:
- **Description of the problem** (what you were trying to do)
- **Error messages** (copy the exact text)
- **File information** (size, format, sample data)
- **Steps you took** (what you did before the problem occurred)
- **Expected vs. actual results** (what you expected vs. what happened)

### Self-Help Resources

- **Check this guide** for common solutions
- **Review error messages** carefully for clues
- **Try with a different file** to isolate the problem
- **Check file format** and data structure

---

## Tips and Tricks

### Efficiency Tips

1. **Batch processing**: Process multiple similar files together
2. **Template consistency**: Use the same data format across all files
3. **Regular validation**: Check data quality before processing
4. **Organized workflow**: Establish a consistent process for file handling

### Quality Assurance

1. **Double-check data**: Always review extracted data before generating reports
2. **Validate totals**: Make sure summary totals match your expectations
3. **Compare outputs**: Check that generated reports match your source data
4. **Keep records**: Maintain logs of processed files and any issues

### Time-Saving Strategies

1. **Standardize inputs**: Use consistent file formats and structures
2. **Automate routine tasks**: Set up regular processing schedules
3. **Use templates**: Create standard file templates for your data
4. **Train team members**: Ensure everyone follows the same process

---

## Conclusion

The Logistics Report Generator is designed to make your work easier and more efficient. By automating the process of creating professional reports from raw data, it saves you time and reduces errors.

### Key Benefits

- **Save hours** of manual data processing
- **Reduce errors** through automated validation
- **Create professional reports** consistently
- **Handle complex data** easily
- **Improve data quality** through validation

### Getting the Most Out of It

- **Learn the features**: Take time to understand what the application can do
- **Use it regularly**: The more you use it, the more efficient you'll become
- **Provide feedback**: Let your team know about any issues or improvements needed
- **Stay updated**: Keep up with any new features or improvements

Remember, this tool is here to help you work smarter, not harder. Don't hesitate to ask for help when you need it, and always review your results to ensure they meet your business needs.

---

*This user guide is designed to help you get the most out of the Logistics Report Generator. For technical support or advanced features, please contact your IT team or system administrator.* 