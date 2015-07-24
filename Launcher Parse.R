library(dplyr)
library(lubridate)

LauncherParser <- function(start = today() %m-% months(1), end = today())
{
    names <- vector()
    for(i in 1:10) {names <- c(names,paste0("V",i))}
    
    getLogs <- function(folder) {
        logs <- list.files(folder, pattern="\\.log$", full.names=T)
        allLogs <- data.frame()
        for(i in 1:length(logs)) {
            df <- read.csv(logs[i], header=F, strip.white=T, col.names=names, colClasses=rep("character", 10))
            allLogs <- bind_rows(allLogs, df)
        }
        allLogs <- mutate(allLogs, V1 = ymd_hms(V1))
        return(allLogs)
    }
    
    launches <- bind_rows(getLogs("C:\\Users\\mpfammatter\\AppData\\Roaming\\WHA Revit Launcher\\Log"),
        getLogs("C:\\Users\\mpfammatter\\AppData\\Roaming\\Revit Launcher Dev\\Log"),
        getLogs("\\\\wa-127\\c$\\Users\\mpfammatter\\AppData\\Roaming\\WHA Revit Launcher\\Log"),
        getLogs("\\\\wa-127\\c$\\Users\\mpfammatter\\AppData\\Roaming\\Revit Launcher Dev\\Log")) %>% 
        filter(V1 >= as.character(start) & V1 <= as.character(end)) %>%
        filter(V2 == "Launcher") %>%
        filter(V3 == "Detach" | V3 == "Standard" | V3 == "detach") %>%
        arrange(V1)
    return(launches)
}