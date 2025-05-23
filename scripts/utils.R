#--------------------------------------------------------------------------
# Function for cleaning Count tab of aquatic invertebrate taxon counts
#--------------------------------------------------------------------------

# filepath is the quoted file location of the spreadsheet. Defaults to a data folder, if not specified.
# filename is the quoted name of the spreadsheet complete with the .xlsx ending
# worksheet is the quoted name of the tab in the spreadsheet with taxon counts. Defaults to "Count" if not specified
# save is logical. TRUE (default) will save a csv of the output to the specified filepath with the same filename and
  # a date stamp. FALSE does not save output to file.
# db_path = quoted path and file name of the database to append to
# square_df = data frame to append to tbl_Squares which was compiled using compile_squares() function
# count_df = data frame to append to tbl_Count which was compiled using compile_counts() function

#---- Function that compiles squares records to append to tblSquares ----
compile_squares <- function(filepath = "./data/", filename = NA, worksheet = "Squares", save = FALSE){
  #---- Bug/error handling ----
  # add / to end of filepath if doesn't exist
  if(!grepl("/$", filepath)){filepath <- paste0(filepath, "/")}
  # check that a filename was specified
  if(is.na(filename)){stop("Must specify the spreadsheet name as filename.")}
  # check that specified filepath exists
  if(!file.exists(filepath)){stop("Speficied filepath does not exist.")}
  # check that specified filename exists
  if(!file.exists(paste0(filepath, filename))){
    stop("Speficied filepath and filename combination does not exist.")}

  # check for readxl package to be installed. If not installed, asks user to install in console.
  if(!requireNamespace("readxl", quietly = TRUE)){
    stop("Package 'readxl' needed for this function to work. Please install it.", call. = FALSE)
  }

  #---- Prepare data for tbl_Squares ----
  tblSquares <- tryCatch(readxl::read_xlsx(path = paste0(filepath, filename), sheet = worksheet) |> data.frame(),
                         error = function(e){
                           stop(paste0("Unable to open ", paste0(filepath, filename), ". Make sure spreadsheet is not currently open in Excel."))
                         })
  dets <- as.POSIXct(round(Sys.time(), units = "sec"), format = "%Y-%m%-%d %H:%M:%S", tz = "UTC")
  
  dets_frm <- as.POSIXct(dets, format = "%Y-%m-%d %H:%M", tz = "UTC")#tz = Sys.timezone())
  tblSquares$DETimeStamp <- dets_frm
  
  # Check for logical dates
  curr_year <- format(Sys.Date(), "%Y")
  season_check <- tblSquares |> dplyr::mutate(DEyear = format(DETimeStamp, "%Y")) |> 
    dplyr::filter(DEyear < Season | Season > curr_year)
  
  error_tally <- data.frame(error = "Season entry greater than data entry timestamp or in the future", num_records = nrow(season_check))
  
  if(nrow(season_check) > 0){
    assign("season_check", season_check, envir = .GlobalEnv)
    warning(paste0(
      "There's at least one 'Season' entry that is either greater than the data entry time stamp or is in the future.", 
      "\n", "Run View(season_check) for more details."))}
  
  # Check that there are no missing required fields
  squares_blank <- tblSquares[
    !complete.cases(tblSquares[,c("LocationID", "EventID", "Season", "RiffleNo", "Replicate", "PerSample", "DETimeStamp")]),]

  error_tally <- rbind(error_tally, 
                       data.frame(error = "Blanks in required Squares table", num_records = nrow(squares_blank)))
  
  if(nrow(squares_blank) > 0){
    assign("squares_blank", squares_blank, envir = .GlobalEnv)
    warning(paste0("At least one required value is missing (see above). Cannot append until record is fixed. ",
                "\n", "Run View(squares_blank) for more details"))}

  #---- Write to file if save = T ----
  new_file = sub(".xlsx", "", filename)
  date_stamp = format(Sys.Date(), format = "%Y%m%d")

  if(save == TRUE){write.csv(tblSquares,
                             paste0(filepath, new_file, "_Squares_", date_stamp, ".csv"),
                             row.names = F)
  }

  errors <- error_tally |> dplyr::filter(num_records > 0)
  
  if(sum(error_tally$num_records) > 0){
    stop("The following errors were detected in the Squares tab preventing data from being appended to the database. ", "\n", 
         paste0("\t", "-", errors$error, " (n=", errors$num_records, ")", collapse = "\n"), "\n", 
         "  See warnings below for more details", "\n\n")}
  #---- Return final dataset ----
  return(data.frame(tblSquares))
}



#---- Function that compiles count records to append to tblCount ----
compile_counts <- function(filepath = "./data/", filename = NA, worksheet = "Count", save = FALSE){

  #---- Bug/error handling ----
  # add / to end of filepath if doesn't exist
  if(!grepl("/$", filepath)){filepath <- paste0(filepath, "/")}
  # check that a filename was specified
  if(is.na(filename)){stop("Must specify the spreadsheet name as filename.")}
  # check that specified filepath exists
  if(!file.exists(filepath)){stop("Speficied filepath does not exist.")}
  # check that specified filename exists
  if(!file.exists(paste0(filepath, filename))){
    stop("Speficied filepath and filename combination does not exist.")}

  # check for readxl package to be installed. If not installed, asks user to install in console.
  if(!requireNamespace("readxl", quietly = TRUE)){
    stop("Package 'readxl' needed for this function to work. Please install it.", call. = FALSE)
  }

  # check for dplyr package to be installed. If not installed, asks user to install in console.
  if(!requireNamespace("dplyr", quietly = TRUE)){
    stop("Package 'dplyr' needed for this function to work. Please install it.", call. = FALSE)
  }

  #---- Prepare data for tbl_Count ----
  # import spreadsheet, only including the specified tab
  count_orig <- tryCatch(readxl::read_xlsx(path = paste0(filepath, filename), sheet = worksheet,
                                           col_types = c("text", "text", "numeric", "numeric", "text", 
                                                         "text", "text", "numeric", "text", "text"),
                                           na = c("")) |> data.frame(),
                         error = function(e){
                         stop(paste0("Unable to open ", paste0(filepath, filename), 
                                     ". Make sure spreadsheet is not currently open in Excel."))
                         })
  
  #count_orig$RepCount <- gsub("", NA_real_, count_orig$RepCount)
  count_orig$LargeRare <- gsub("", NA_character_, count_orig$LargeRare)
  count_orig$RefCollection <- gsub("", NA_character_, count_orig$RefCollection)
  
  taxon_check <- count_orig |> dplyr::filter(!is.na(RepCount)) |> dplyr::filter(is.na(TaxonCode))
  
  error_tally <- data.frame(error = "Non-zero RepCount missing a TaxonCode", num_records = nrow(taxon_check))
  
  if(nrow(taxon_check) > 0){
    assign("taxon_check", taxon_check, envir = .GlobalEnv)
    warning(paste0("There is at least one row with a non-zero count that has a blank TaxonCode.",
                "\n", "Run View(taxon_check) for more details."))}
  
  lr_check <- count_orig |> dplyr::filter(LargeRare == 0)
  if(nrow(lr_check) > 0){
    assign("lr_check", lr_check, envir = .GlobalEnv)
    warning(paste0("There is at least one row a LargeRare recorded as 0 instead of -1 or blank.",
                   "\n", "Run View(lr_check) for more details."))}
  
  rc_check <- count_orig |> dplyr::filter(RefCollection == 0)
  if(nrow(rc_check) > 0){
    assign("rc_check", rc_check, envir = .GlobalEnv)
    warning(paste0("There is at least one row with a non-zero count that has a blank TaxonCode.",
                   "\n", "Run View(rc_check) for more details."))}
  
  error_tally <- rbind(error_tally,
                       data.frame(error = "LargeRare = 0", num_records = nrow(lr_check)),
                       data.frame(error = "RefCollection = 0", num_records = nrow(rc_check)))
  # Setting numerics to integers to match database defs
  count_orig$Season <- as.integer(count_orig$Season)
  count_orig$RiffleNo <- as.integer(count_orig$RiffleNo)
  count_orig$RepCount <- suppressWarnings(as.integer(count_orig$RepCount)) # expect NAs introduced by coercion

  # drop all rows where TaxonCode is blank
  count_trim <- count_orig[!is.na(count_orig$TaxonCode),]

  # Add 2 required columns in tbl_Count
  ## Add DETimeStamp
  dets <- as.POSIXct(round(Sys.time(), units = "sec"), format = "%Y-%m%-%d %H:%M:%S", tz = "UTC")
  dets_frm <- as.POSIXct(dets, format = "%Y-%m-%d %H:%M", tz = "UTC")#tz = Sys.timezone())
  count_trim$DETimeStamp <- dets_frm

  # Check for logical dates
  curr_year <- format(Sys.Date(), "%Y")
  season_check <- count_trim |> dplyr::mutate(DEyear = format(DETimeStamp, "%Y")) |> 
    dplyr::filter(DEyear < Season | Season > curr_year)
  
  error_tally <- rbind(error_tally, 
                  data.frame(error = "Season entry greater than data entry timestamp or in the future", 
                             num_records = nrow(season_check)))
  
  if(nrow(season_check) > 0){
    assign("season_check", season_check, envir = .GlobalEnv)
    warning(paste0(
      "There's at least one 'Season' entry that is either greater than the data entry time stamp or is in the future.", 
      "Run View(season_check) data frame for more details."))}
  
  ## Add PrevTaxonCode if it doesn't already exist in the data
  if(!all(names(count_trim) %in% c("PrevTaxonCode"))){count_trim$PrevTaxonCode <- NA_character_}

  # Change blanks to FALSE for LargeRare and RefCollection and TRUE for 1
  count_trim$LargeRare <- ifelse(is.na(count_trim$LargeRare), FALSE, TRUE)
  count_trim$RefCollection <- ifelse(is.na(count_trim$RefCollection), FALSE, TRUE)

  # Check for duplicate species to combine, as database won't allow duplicate species within same riffle and replicate
  check_cols = c("LocationID", "EventID", "Season", "RiffleNo", "Replicate", "TaxonCode")
  spp_dup <- count_trim[,check_cols][duplicated(count_trim[,check_cols]), ]
  
  error_tally <- rbind(error_tally,
                       data.frame(error = "Duplicate species detected", num_records = nrow(spp_dup)))
  
  if(nrow(spp_dup) > 0){
    assign("spp_dup", spp_dup, envir = .GlobalEnv)
    warning(paste0("There are duplicate taxon codes within an individual riffle and replicate.",
                   "\n","Please check that duplicates are not due to an error before appending to the database. ",
                   "\n","Duplicates were summed in returned data frame.", 
                   "\n","Run view(spp_dup) for more details."))}

  count_trim2 <-
  if(nrow(spp_dup) > 0){
    count_trim |>
      dplyr::group_by(LocationID, EventID, Season, RiffleNo, Replicate, TaxonCode, LargeRare,
                                  Note, RefCollection, DETimeStamp, PrevTaxonCode) |>
      dplyr::summarize(RepCount = sum(RepCount), .groups = 'drop')
  } else {count_trim}

  # Change column order to match tbl_Count in database
  count_final <- count_trim2[,c("LocationID", "EventID", "Season", "RiffleNo", "Replicate", "TaxonCode",
                               "PrevTaxonCode", "LargeRare", "RepCount", "Note",  "DETimeStamp",
                               "RefCollection")]

  #---- QC results before saving ----
  # Check that number of records by taxa are the same between original and trimmed
  count_orig$RefCollection[is.na(count_orig$RefCollection)] <- 0
  count_orig$Replicate[is.na(count_orig$Replicate)] <- 0
  count_orig$LargeRare[is.na(count_orig$LargeRare)] <- 0
  count_orig$RepCount[is.na(count_orig$RepCount)] <- 0
  
  count_trim2$RefCollection[is.na(count_trim2$RefCollection)] <- 0
  count_trim2$Replicate[is.na(count_trim2$Replicate)] <- 0
  count_trim2$LargeRare[is.na(count_trim2$LargeRare)] <- 0
  count_trim2$RepCount[is.na(count_trim2$RepCount)] <- 0
  
  cnt_sum_or <- aggregate(RepCount ~ EventID + TaxonCode + RiffleNo + Replicate + RepCount,
                          data = count_orig, FUN = function(x){sum(x)})
  cnt_sum_tr <- aggregate(RepCount ~ EventID + TaxonCode + RiffleNo + Replicate + RepCount,
                          data = count_trim2, FUN = function(x){sum(x)})
  cnt_comb <- merge(cnt_sum_or, cnt_sum_tr, by = c("EventID", "TaxonCode", "RiffleNo", "Replicate"),
                  all.x = T, all.y = T, suffixes = c("_orig", "_trim"))
  cnt_check <- cnt_comb[cnt_comb$RepCount_orig != cnt_comb$RepCount_trim,]

  # Error checking on trim
  if(nrow(cnt_check) > 0){
    assign("data_check", cnt_check, envir = .GlobalEnv)
    warning(paste0("Trimming white space resulted in an error in the following taxa: ",
           paste0(cnt_check$TaxonCode, collapse = ", "),
           ". Run View(data_check) for more details."))}
  
  error_tally <- rbind(error_tally,
                       data.frame(error = "Issues resulting from trimming white space", num_records = nrow(cnt_check)))

  # Check that RepCount isn't blank
  miss_count <- count_trim2[is.na(count_trim2$RepCount),]
  if(nrow(miss_count) > 0){
    assign("miss_count", miss_count, envir = .GlobalEnv)
    warning(paste0("There are ", nrow(miss_count), " records with a missing RepCount. ", 
                "Counts can't be appended until blank counts are resolved",
                "Run View(miss_count) for more details."))}

  error_tally <- rbind(error_tally,
                       data.frame(error = "TaxonCodes with missing RepCount", num_records = nrow(miss_count)))
  
  #---- Write to file if save = T ----
  new_file = sub(".xlsx", "", filename)
  date_stamp = format(Sys.Date(), format = "%Y%m%d")

  if(save == TRUE){write.csv(count_final,
                             paste0(filepath, new_file, "_", "Count", "_", date_stamp, ".csv"),
                             row.names = F)}

  errors <- error_tally |> dplyr::filter(num_records > 0)
  errors2 <- errors |> dplyr::filter(!error %in% "Duplicate species detected")   # drop spp dups from stop 

  if(sum(errors2$num_records) > 0){
    stop("The following errors were detected in the Count tab preventing data from being appended to the database. ", "\n", 
         paste0("\t", "-", errors$error, " (n=", errors$num_records, ")", collapse = "\n"), "\n", 
         "  See warnings below for more details", "\n\n")}
  
  #---- Return final dataset ----
  return(data.frame(count_final))
}



#---- Function that appends square records to database ----
append_squares <- function(db_path =  "./data/HTLNInvert3.9.0 - Copy.accdb", square_df = NA){

  # check for dplyr package to be installed. If not installed, asks user to install in console.
  if(!requireNamespace("dplyr", quietly = TRUE)){
    stop("Package 'dplyr' needed for this function to work. Please install it.", call. = FALSE)
  }

  # check for odbc package to be installed. If not installed, asks user to install in console.
  if(!requireNamespace("odbc", quietly = TRUE)){
    stop("Package 'odbc' needed for this function to work. Please install it.", call. = FALSE)
  }

  # check for DBI package to be installed. If not installed, asks user to install in console.
  if(!requireNamespace("DBI", quietly = TRUE)){
    stop("Package 'DBI' needed for this function to work. Please install it.", call. = FALSE)
  }

  options(odbc.batch_rows = 1) # Set option to allow for row append more than one row at a time

  # Try to connect to the database. If the connection doesn't work, you'll get the error message in the stop()
  tryCatch(
    db <- DBI::dbConnect(drv = odbc::odbc(),
                         .connection_string =
                         paste0("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=", db_path)),
    error = function(e){
      stop(paste0("Unable to connect to specified database. Check that you have the correct db_path"))})

  # Check that I can read the tables, and get highest autonumbers from tbl_Squares
  # and will use to identify the rows that appended later in a anti-join.
  tbl_square_orig <- DBI::dbReadTable(db, "tbl_Squares") # read existing table into R
  recIDmax <- max(tbl_square_orig$RecID, na.rm = T) # highest autonumber in existing db table
  square_df$RecID <- recIDmax + as.numeric(row.names(square_df)) # Adds autonumber
  # Reorder columns
  square_df <- square_df[,c("RecID", "LocationID", "EventID", "Season",
                            "RiffleNo", "Replicate", "PerSample", "DETimeStamp")]

  # Change numeric to integer to match db table defs. Doing this here instead of compile_squares b/c of RecID.
  square_df$RecID <- as.integer(square_df$RecID)
  square_df$PerSample <- as.integer(square_df$PerSample)
  square_df$RiffleNo <- as.integer(square_df$RiffleNo)
  square_df$Season <- as.integer(square_df$Season)

  #options(digits.secs=0)

  #DBI::dbWriteTable(db, "tbl_Squares", square_df, append = T)

  # Append rows to tbl_Squares
  tryCatch(
  DBI::dbWriteTable(db, "tbl_Squares", square_df, append = T), # append new rows
  error = function(e){
    DBI::dbDisconnect(db)
    if(grepl("You cannot add or change a record", e)){
      print(paste0("The Square data frame you are trying to append does not yet have matching records in tbl_SamplingEvents.",
      " Check that you've entered the tbl_SamplingEvents and tbl_SamplingPeriod records, and that the EventIDs match between",
      " those generated in the Excel spreadsheet and the tbl_SamplingEvents and tbl_SamplingPeriod in the database"))
      } else print(e)}
  )

  # Check that append worked
  tbl_square_app <- DBI::dbReadTable(db, "tbl_Squares") # pull new appended table down
  orig_rows <- nrow(tbl_square_orig) # num rows in tbl_squares before the append
  app_rows <- nrow(tbl_square_app) # num rows in tbl_squares after append
  diff_rows <- app_rows - orig_rows # calc # of rows appended
  app_df_rows <- nrow(square_df) # rows in df to append
  missing_rows <- app_df_rows - diff_rows

  # Use anti-join to determine any differences in rows/columns
  join_cols <- names(square_df) # columns to check

  # Find rows that were appended in the database
  new_db_rows <- dplyr::anti_join(tbl_square_app, tbl_square_orig, by = join_cols)

  # Find rows that should have been appended, but were not.
  rows_not_matched <- dplyr::anti_join(new_db_rows, square_df, by = join_cols)

  # setdiff(tbl_square_app, tbl_square_orig) base R version of anti_join
  # setdiff(new_db_rows, square_df) base R version of anti_join

  if(nrow(rows_not_matched) > 0){
    # delete appended rows if validation returns unmatched or missing rows
    new_recs <- square_df$RecID
    del_qry <- sprintf("DELETE from tbl_Squares where RecID IN (%s)",
                       paste0(as.integer(new_recs), collapse = ", "))
    DBI::dbSendQuery(db, del_qry)
    DBI::dbDisconnect(db)

    # add dataframe of missing or mismatched values to global environment to help with troubleshooting
    assign("squares_unmatched", rows_not_matched, envir = .GlobalEnv)

    stop(paste0("Not all squares rows were appended. Missing ", missing_rows, " from dataset.",
    "The missing records are in a data frame called *squares_unmatched* in your global environment."))

    } else if(nrow(rows_not_matched) == 0){
    cat(paste0("Success! Appended ", diff_rows,
               " records to tbl_Squares in the specified database and were validated as identical to the compiled squares data frame in R."))
    }

  # Close database connection
  DBI::dbDisconnect(db)
}



#---- Function that appends square records to database ----
append_counts <- function(db_path =  "./data/HTLNInvert3.9.0 - Copy.accdb", count_df = NA){
  # check for dplyr package to be installed. If not installed, asks user to install in console.
  if(!requireNamespace("dplyr", quietly = TRUE)){
    stop("Package 'dplyr' needed for this function to work. Please install it.", call. = FALSE)
  }

  # check for odbc package to be installed. If not installed, asks user to install in console.
  if(!requireNamespace("odbc", quietly = TRUE)){
    stop("Package 'odbc' needed for this function to work. Please install it.", call. = FALSE)
  }

  # check for DBI package to be installed. If not installed, asks user to install in console.
  if(!requireNamespace("DBI", quietly = TRUE)){
    stop("Package 'DBI' needed for this function to work. Please install it.", call. = FALSE)
  }

  options(odbc.batch_rows = 1) # Set option to allow for row append more than one row at a time

  # Try to connect to the database. If the connection doesn't work, you'll get the error message in the stop()
  tryCatch(
    db <- DBI::dbConnect(drv = odbc::odbc(),
                         .connection_string =
                           paste0("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=", db_path)),
    error = function(e){
      stop(paste0("Unable to connect to specified database. Check that you have the correct db_path"))})

  # Check that I can read the tables. Will use the DETimeStamp to identify the rows that appended later in a anti-join.
  tbl_count_orig <- DBI::dbReadTable(db, "tbl_Count") # read existing table into R

  # Read in taxa lookup table to find any TaxonCodes not on the list
  tlu_Taxa <- DBI::dbReadTable(db, "tlu_Taxa")
  # Check for taxa in count_df missing from tlu_Taxa
  miss_spp <- setdiff(count_df$TaxonCode, tlu_Taxa$TaxonCode)

  if(length(miss_spp) > 0){
    DBI::dbDisconnect(db)
    stop(paste0("The following TaxonCodes in the Count spreadsheet are not in the database tlu_Taxa table. "),
        "\n\t", paste(miss_spp, collapse = ", "))
    }

  # Append rows to tbl_Squares
  tryCatch(
    DBI::dbWriteTable(db, "tbl_Count", count_df, append = T), # append new rows
    error = function(e){
      DBI::dbDisconnect(db)
      if(grepl("You cannot add or change a record", e)){
        stop(paste0("The Count data frame you are trying to append does not yet have matching records in tbl_Squares.",
                     " Check that you've appended the tbl_Squares records first, and that you've entered the ",
                     "tbl_SamplingEvents and tbl_SamplingPeriod records, and that the EventIDs match between",
                     " those generated in the Excel spreadsheet and the tbl_Squares, tbl_SamplingEvents and ",
                     "tbl_SamplingPeriod in the database."))
      } else {stop(print(e))}
      }
  )

  # Check that append worked
  tbl_count_app <- DBI::dbReadTable(db, "tbl_Count") # pull new appended table down
  orig_rows <- nrow(tbl_count_orig) # num rows in tbl_squares before the append
  app_rows <- nrow(tbl_count_app) # num rows in tbl_squares after append
  diff_rows <- app_rows - orig_rows # calc # of rows appended
  app_df_rows <- nrow(count_df) # rows in df to append
  missing_rows <- app_df_rows - diff_rows

  # Use anti-join to determine any differences in rows/columns
  join_cols <- names(count_df) # columns to check

  # Find rows that were appended in the database
  new_db_rows <- dplyr::anti_join(tbl_count_app, tbl_count_orig, by = join_cols)

  # Find rows that should have been appended, but were not.
  rows_not_matched <- dplyr::anti_join(new_db_rows, count_df, by = join_cols)

  if(nrow(rows_not_matched) > 0){
    # delete appended rows if validation returns unmatched or missing rows
    new_recs <- count_df$DETimeStamp
    del_qry <- sprintf("DELETE from tbl_Count where DETimeStamp IN (%s)",
                       paste0(as.integer(new_recs), collapse = ", "))
    DBI::dbSendQuery(db, del_qry)
    DBI::dbDisconnect(db)

    # add dataframe of missing or mismatched values to global environment to help with troubleshooting
    assign("counts_unmatched", rows_not_matched, envir = .GlobalEnv)

    stop(paste0("Not all counts rows were appended. Missing ", missing_rows, " from dataset.",
                "The missing records are in a data frame called *counts_unmatched* in your global environment."))

  } else if(nrow(rows_not_matched) == 0){
    cat(paste0("Success! Appended ", diff_rows,
               " records to tbl_Count in the specified database and were validated as identical to the compiled count data frame in R."))
  }


  # Close database connection
  DBI::dbDisconnect(db)
}
