## SQL --> Excel

### What

If you have a `sql` query, and need to farm it out into 1,000+ excel workbooks (for whatever reason..), this file will do it for you. Just specify the path of the output director (for the 1,000+ files), and run the macro

### How

The code stores the contents of the `sql` query in memory, and then iteratively goes through and creates a new workbook when it sees a new merchant