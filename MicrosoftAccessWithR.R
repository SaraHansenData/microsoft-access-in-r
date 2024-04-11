# Working with Microsoft Access data in R

## Sara Hansen, hanse2s
## Modified April 11, 2024

library(tidyverse) # to write nice code
library(odbc) # to check drivers
library(RODBC) # to connect to Access

# Flat files (txt, csv) or Excel workbooks can be useful for analysis
# However, many times we want a relational database to minimize redundancy
# and reduce storage space of our files

# One of the most user-friendly relational database systems is Microsoft Access

# This script will cover importing and exporting flat files and Access data
# using a small subset of the Prairie Fen Biodiversity Project Database
# available via GBIF: https://www.gbif.org/dataset/72c4d3c6-5b8d-49f5-bfbe-febd53849588

################################################################################

# Step 1: Read in the flat file
dat <- read.csv("data/prairie-fen-data-flat.csv", header = TRUE, 
                fileEncoding = "UTF-8-BOM")

# Step 2: Explore data
glimpse(dat)
# We have three main identifiers: occurrenceID, eventID, and locationID
# This tells us we have data that are arranged based on some physical hierarchy
# In this case, we have prairie fens (locations), with sampling points (events),
# and species presences at sampling points (occurrence)

# This is a good time for a relational database!

# Step 3: Set up a relational database in R
# We'll set up three distinct files that minimize redundancy
location <- dat %>%
  select(locationID, locality, countryCode, stateProvince, county) %>%
  distinct(locationID, .keep_all = TRUE)
# Now, rather than repeating all this location information for every record,
# we have just one record for each location in the database

event <- dat %>%
  select(eventID, eventDate, year, month, day, decimalLatitude, decimalLongitude,
         locationID) %>%
  distinct(eventID, .keep_all = TRUE) %>%
  # we also want a standard date format
  mutate(eventDate = as.character(as.Date(eventDate, format = "%m/%d/%Y")))
# Now we have one record for each sampling event
# Why did we include locationID in the event table?
# In the location table, locationID is a primary key which is a unique identifier
# In the event table, eventID is the primary key and locationID is a foreign key
# A foreign key is like a pointer to another table, letting us know how to link records

occurrence <- dat %>%
  select(occurrenceID, scientificName, recordedBy, catalogNumber, recordNumber,
         occurrenceRemarks, eventID)
# Note that the occurrence table has the same number of rows as the flat file,
# but we significantly reduced the number of columns and computational cost
# eventID is the foreign key -> occurrence happen in events, which happen in locations

# What is the primary key? There seem to be a few options
occurrence %>% distinct(occurrenceID) %>% count() #325
occurrence %>% distinct(catalogNumber) %>% count() #325
occurrence %>% distinct(recordNumber) %>% count()  #126

# occurrenceID and catalogNumber are both unique identifiers, so either can be a primary key
# recordNumber has fewer distinct values than records, so it is not a unique identifier
# For now, we'll actually drop occurrenceID because it is long and takes a lot of space
occurrence <- occurrence %>% select(-occurrenceID)

# Now we have a relational database of locations, events, and occurrences!
# We went from a data frame (dat) that takes about 167,000 bytes
# to three smaller data frames that collectively take about 114,000 bytes
# Imagine how much we could reduce storage space of massive data sets!

################################################################################

# Now we are ready to export these data to an MS Access database for safekeeping

# Step 4: Check Access drivers
odbc::odbcListDrivers()
# If the "Microsoft Access Driver" isn't listed, download and install it to your computer

# Step 5: Create a blank Access database in Access
# Then connect to it here by providing the file path
# A blank Access database is already available in the data folder for this project
database <- "data/prairie-fen-database.accdb"
connection <- odbcConnectAccess2007(database)

# Step 6: Write the database out into the blank Access database
# Below function checks whether a table already exists and deletes it if yes,
# then saves the table to Access
# If the string length exceeds 255, the type is automatically set to "LONGTEXT",
# otherwise it is "VARCHAR(255)" which is the default
writeOut <- function(connection, x) {
  
  if(deparse(substitute(x)) %in% sqlTables(connection)$TABLE_NAME) {
    sqlDrop(connection, deparse(substitute(x))) 
  }
  
  varTypeVector = purrr::map(x, ~max(stringr::str_length(.x))) %>%
    as.data.frame() %>%
    pivot_longer(cols = everything(),
                 names_to = "column", values_to = "max_length") %>%
    mutate(columnType = case_when(max_length < 255 ~ "VARCHAR(255)",
                                  max_length >= 255 ~ "LONGTEXT")) %>%
    pull(columnType)
  
  names(varTypeVector) <- colnames(x)
  
  sqlSave(connection, x, tablename = deparse(substitute(x)), 
          rownames = FALSE, varTypes = varTypeVector)
  # note that Access drivers don't support primary key creation
}

# Alternatively, you could specify column types individually
# But keeping columns as character format is often preferable to avoid automatic edits

#sqlSave(connection, event, varTypes = 
          #c(eventID = "VARCHAR(255)", eventDate = "DATE",
            #year = "INT", month = "INT", day = "INT",
            #decimalLatitude = "FLOAT", decimalLongitude = "FLOAT",
            #locationID = "VARCHAR(255)"))

# Write out each table to the Access database
writeOut(connection, location)
writeOut(connection, event)
writeOut(connection, occurrence)

################################################################################

# You now have an Access database! Check it out in Access

################################################################################

# What if we need to work with the Access data in R again?
# We can connect to the database and query from it
# The connection is made in the same way
database <- "data/prairie-fen-database.accdb"
connection <- odbcConnectAccess2007(database)
# We don't need to redo this while we're still in R,
# but we do need to re-establish the connection each time we open R

# We can easily fetch each table
location2 <- sqlFetch(connection, "location")
event2 <- sqlFetch(connection, "event")
occurrence2 <- sqlFetch(connection, "occurrence")

# Note each of this is exactly the same as the original version
# All we did was write it out into Access and then read it back in

# We can do whatever we want with the data inside R, and it won't affect the actual database
# For example, we can join all the data back together into one file again
dat2 <- occurrence2 %>%
  full_join(event2, by = "eventID") %>%
  full_join(location2, by = "locationID")

# We won't always want to fetch the entirety of a table,
# especially when we're working with huge databases

# In those cases, it's better to craft SQL queries to extract smaller portions of the data
# Note that we are querying the Access database itself, not our R objects

# For example, we might want only occurrences in the quadrat "BVF11"
sqlQuery(connection,
         query = "SELECT * FROM occurrence WHERE eventID = 'BVF11'")
# SELECT * means we want all the columns, but we can also choose a small subset

# For example, what species were observed in quadrats "BVF11" and "BVF12"?
sqlQuery(connection,
         query = "SELECT DISTINCT(scientificName) FROM occurrence WHERE eventID IN ('BVF11', 'BVF12')")
# In this case, we added DISTINCT() because we only want to know the species that  occurred
# and we don't want any duplicates

# We can also aggregate data in our queries
# For example, how many times was each species observed?
# We'll order it with the most frequent species on top
sqlQuery(connection,
         query = "SELECT scientificName, COUNT(*) AS count
                  FROM occurrence GROUP BY scientificName ORDER BY COUNT(*) DESC")
# We aliased the count column as "count" and made sure to group by scientific name so we'd get the count of each species

# Sometimes we need information from multiple tables at once
# For example, how many times was Dasiphora fruticosa observed in Little Appleton Lake Fen?
sqlQuery(connection,
         query = "SELECT COUNT(*) as count FROM occurrence 
                  INNER JOIN event ON occurrence.eventID = event.eventID 
                  WHERE occurrence.scientificName = 'Dasiphora fruticosa (L.) Rydb.' 
                    AND locationID = 'LAL'")

################################################################################

# Sometimes when we start exploring our data further, we realize we need to make some changes
# We can update the Access database from R
# Note that in the following examples, we will be editing Access data in R,
# unlike in the previous queries where we were pulling from Access without changing the database

# We might decide to add another location to the location table
# First, which columns are in the location table? We need to make sure we enter data in the correct order
sqlColumns(connection, "location")

# dplyr::rows_insert comes in handy
sqlDrop(connection, "location")
sqlSave(connection,
        dat = location %>% 
          rows_insert(data.frame(
            locationID = "WLQ", locality = "Whelan Lake Fen", 
            countryCode = "US", stateProvince = "Michigan", 
            county = "Washtenaw")),
        rownames = FALSE)

# We can query the location table to see it got updated
sqlQuery(connection,
         query = "SELECT * FROM location")

# Why did we have to drop the location table first?
# The Access drivers don't support primary keys,
# so we don't have a way to tell Access we don't want to duplicate the same records
# If we didn't drop the table and added append = TRUE to sqlSave(),
# we would get duplicate rows of Bridge Valley Fen and Little Appleton Lake Fen

# Sometimes we just need to update a single value
# For example, let's say Betula pumila, which is bog birch,
# in LAL9 should actually be Betula papyrifera, which is paper birch

# We can update the record in the occurrence table,
# but because we don't have a primary key Access actually won't let us update
# So we're going to basically replace the occurrence table
sqlDrop(connection, "occurrence")
sqlSave(connection,
        dat = occurrence %>% 
          mutate(scientificName = case_when(
            scientificName == "Betula pumila L." &
              catalogNumber == "LAL9-4" ~ "Betula papyrifera Marshall",
            TRUE ~ scientificName)))

# We can query to make sure the change worked
sqlQuery(connection,
         query = "SELECT scientificName FROM occurrence WHERE catalogNumber = 'LAL9-4'")

################################################################################

# Often we want to save a query to an external file
# Let's build a query, assign it to an object, then write it out as a flat file

# Where are all the locations of prairie sedge (Carex prairea)?
prairie_sedge_points <- sqlQuery(connection,
                                 query = "SELECT event.eventID, decimalLatitude, decimalLongitude
                                          FROM event
                                          INNER JOIN occurrence
                                          ON event.eventID = occurrence.eventID
                                          WHERE occurrence.scientificName = 'Carex prairea Dewey'")

# We'll write it out as a text file for mapping somewhere else
write.table(prairie_sedge_points, "data/prairie_sedge_points.txt",
            row.names = FALSE, col.names = TRUE, sep = "\t", quote = FALSE,
            fileEncoding = "UTF-8", append = FALSE)

################################################################################

# We started with a flat file and ended with a flat file, and did a lot in between!


