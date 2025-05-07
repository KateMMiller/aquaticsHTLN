# Source functions script
source("./scripts/utils.R")

#++++ Parameters to update/check every time you try to append: ++++
path = "./data/" # Update to path for data to be exported to
spreadsheet = "OZAR_CURRT15_2024_Complete.xlsx" # Update to match spreadsheet name to append
#spreadsheet = "OZAR CURRT15.xlsx" # has wrong EventIDs, good one to test that bug handling catches it properly.
db_filename = "./data/HTLNInvert3.9.0 - Copy.accdb" # Update to path and name of the database to append to

# Compile records to append to tblSquares
square_currt15 <- compile_squares(filepath = path, filename = spreadsheet, worksheet = "Squares", save = TRUE)
square_currt15
nrow(square_currt15) # print number of rows that will be appended = 9

# Compile records to append to tblCount by trimming Count tab
count_currt15 <- compile_counts(filepath = path, filename = spreadsheet, worksheet = "Count", save = TRUE)
# There are 2 duplicates. Code adds them together, but you should check that is the correct action to take.
  # - Riffle 2, Rep L, taxon PLEUEL
  # - Riffle 3, Rep M, taxon PLEUEL
head(count_currt15)
nrow(count_currt15) #237

# Append tbl_Squares records
append_squares(db_path =  db_filename, square = square_currt15)

# Append tbl_Count
append_counts(db_path = db_filename, count = count_currt15)

