#
# Per avviare AppOst selezionare il menu
# Code > Run Region > Run All
#
# oppure premere i tasti:
# "CTRL+ALT+R" (Windows)
# "OPTION+CMD+R" (Mac)
#

if(!require(devtools)) install.packages("devtools")
devtools::install_github("giovabubi/appost")
library(appost)
appost()
