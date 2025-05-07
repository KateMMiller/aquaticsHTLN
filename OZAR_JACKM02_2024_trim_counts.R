# Source functions script
source("./scripts/utils.R")

#++++ Parameters to update/check every time you try to append: ++++
path = "./data/" # Update to path for data to be exported to
spreadsheet = "OZAR_JACKM02_2024_Complete.xlsx"
#spreadsheet = "./temp/OZAR_JACKM02_2024_Complete.xlsx" # missing a percent sampled in squares and PLAOGY is misspelled as PLANGY, so fails- good to test on.

db_filename = "./data/HTLNInvert3.9.0 - Copy.accdb" # Update to path and name of the database to append to

# Compile records to append to tblSquares
square_jackm02 <- compile_squares(filepath = path, filename = spreadsheet, worksheet = "Squares", save = TRUE)
square_jackm02
nrow(square_jackm02) # 9

# Compile records to append to tblCount by trimming Count tab
count_jackm02 <- compile_counts(filepath = path, filename = spreadsheet, worksheet = "Count", save = TRUE)
head(count_jackm02)
nrow(count_jackm02) # 206 rows;

# Append tbl_Squares records
append_squares(db_path =  db_filename, square = square_jackm02)

# Append tbl_Count
append_counts(db_path = db_filename, count = count_jackm02)
# 206
