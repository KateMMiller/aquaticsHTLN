# Source functions script
source("./scripts/utils.R")

#++++ Parameters to update/check every time you try to append: ++++
path = "./data/" # Update to path for data to be exported to
spreadsheet = "OZAR_JACKM01_2024_Complete.xlsx"
db_filename = "./data/HTLNInvert3.9.0 - Copy.accdb" # Update to path and name of the database to append to

# Compile records to append to tblSquares
square_jackm01 <- compile_squares(filepath = path, filename = spreadsheet, worksheet = "Squares", save = TRUE)
square_jackm01
nrow(square_jackm01) # 9
str(square_jackm01)

# Compile records to append to tblCount by trimming Count tab
count_jackm01 <- compile_counts(filepath = path, filename = spreadsheet, worksheet = "Count", save = TRUE)
# There's one duplicate. Code adds them together, but you should check that is the correct action to take.
# - Riffle 2, Rep R, taxon CAENCA
head(count_jackm01)
nrow(count_jackm01) # 159 rows;

# Append tbl_Squares records
append_squares(db_path =  db_filename, square = square_jackm01)

# Append tbl_Count
append_counts(db_path = db_filename, count = count_jackm01)
# 159 appended for OZAR_JACM01_2024
