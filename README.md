# Microsoft Access in R

## Background
Microsoft Access is a useful relational database storage and management software used by many scientists and data professionals. While Access offers a user-friendly interface, sometimes it's better to interact with Access through a programming language to optimize reproducibility and readability. A programming language like R offers many ways to interact with your data that Access does not and is usually the best choice for analyses and visualizations.

This project describes how to go between flat files, R data frames, and Access databases. The target audience is scientists and anyone with R experience and a desire to increase their database management skills. We'll use a combination of the tidyverse and Structured Query Language (SQL) syntax in R to import and export data in various formats, query the Access database, and make changes to the Access database.

## Data
The example data provided are a small snapshot (with minimal edits) of the Prairie Fen Biodiversity Project Database available via GBIF: https://www.gbif.org/dataset/72c4d3c6-5b8d-49f5-bfbe-febd53849588. The Prairie Fen Biodiversity Project is an ongoing effort to study Michigan prairie fens, which are incredibly diverse wetland habitats. I encourage you to learn more about prairie fen diversity through the linked dataset. Please cite as indicated on the GBIF page. 

## Funding
The work that inspired the current project was completed by Central Michigan University and Michigan State University Extension, Michigan Natural Features Inventory with support from the United States Fish and Wildlife Service (USFWS) Great Lakes Restoration Initiative grants #F19AC00653, #F20AC10391, and #F21AC02286. Any opinions, findings, and conclusions or recommendations expressed in this material are those of the author(s) and do not necessarily reflect the views of the USFWS.
