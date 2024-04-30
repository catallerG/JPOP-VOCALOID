library(humdrumR)
library(dplyr)
library(purrr)
library(openxlsx)
library(stringr)

process_krn_file <- function(file) {
  # Read the Humdrum file
  humdrum_data <- readHumdrum(file)
  humdrum_kern <- humdrum_data[[,c('**kern')]]
  
  n <- spines(humdrum_kern)$Spines
  
  token <- with(humdrum_kern, midi(Token), subset = Spine == n)
  dur <- with(humdrum_kern, recip(Token) |> duration(), subset = Spine == n)
  
  #sing higher
  ptoken_filtered <- na.omit(token)
  pdur_filtered <- dur[!is.na(token)]
  
  weighted_average_pitch <- sum(ptoken_filtered*pdur_filtered)/sum(pdur_filtered)
  highest_pitch <- max(ptoken_filtered)
  median_pitch <- median(ptoken_filtered)
  
  #use less rest
  rdur_filtered <- dur[is.na(token)]
  in_melody_rest <- sum(rdur_filtered[rdur_filtered != 1]) #忽略掉所有全休止符的休止符时长合
  
  #average_interval
  intervals <- with(humdrum_kern, kern(Token) |> mint(), subset = Spine == n)
  
  conversion <- list(
    "P1" = 0, "+m2" = 1, "-m2" = -1, "+M2" = 2, "-M2" = -2,
    "+m3" = 3, "-m3" = -3, "+M3" = 4, "-M3" = -4,
    "P4" = 5, "-P4" = -5, "+A4" = 6, "-A4" = -6,
    "+P5" = 7, "-P5" = -7, "+m6" = 8, "-m6" = -8,
    "+M6" = 9, "-M6" = -9, "+m7" = 10, "-m7" = -10,
    "+M7" = 11, "-M7" = -11, "+P8" = 12, "-P8" = -12,
    "+m9" = 13, "-m9" = -13, "+M9" = 14, "-M9" = -14,
    "+m10" = 15, "-m10" = -15, "+M10" = 16, "-M10" = -16,
    "+P11" = 17, "-P11" = -17, "+A11" = 18, "-A11" = -18,
    "+P12" = 19, "-P12" = -19, "+m13" = 20, "-m13" = -20,
    "+M13" = 21, "-M13" = -21, "+m14" = 22, "-m14" = -22,
    "+M14" = 23, "-M14" = -23, "+P15" = 24, "-P15" = -24
  )
  
  numeric_intervals <- unlist(conversion[intervals])
  filtered_numeric_intervals <- numeric_intervals[!is.na(numeric_intervals)]
  average_interval <- mean(abs(filtered_numeric_intervals))
  
  #proportion of skips
  abs_interval_2 <- ifelse(abs(filtered_numeric_intervals) > 2, 1, 0)
  pos <- sum(abs_interval_2)/length(filtered_numeric_intervals)
  
  #density of notes
  f_notes <- within(humdrum_kern$Token, subset=Token %!~% 'r' & Spine==n, Token)
  f_notetab <- with(f_notes, length(Token), by=Bar)
  note_density <- mean(f_notetab)
  
  #bpm - box plot
  bpm <- as.numeric(stringr::str_extract(basename(file), "(?<=BPM=)\\d{2,3}"))
  
  #all the bar plot could be better changed to box plot
  
  #output
  return(list(
    "Title" = tools::file_path_sans_ext(basename(file)),
    "Average Pitch" = weighted_average_pitch,
    "Highest Pitch" = highest_pitch,
    "Median Pitch" = median_pitch,
    "Rest" = in_melody_rest,
    "Positive Max Interval" = max(filtered_numeric_intervals),
    "Minus Max Interval" = min(filtered_numeric_intervals),
    "Average Interval" = average_interval,
    "Proportion of Skips" = pos,
    "Density" = note_density,
    "BPM" = bpm,
    "Year" = basename(dirname(file)),
    "Gender" = ifelse(grepl("FM", file), "Female", "Male")
  ))
}

search_directory <- "/Users/johnsmith/Desktop/7100/J-pop"
#search_directory <- "/Users/johnsmith/Desktop/7100/VOCALOID"
#search_directory <- "/Users/johnsmith/Desktop/7100/fix"

list_krn_files <- function(path) {
  krn_files <- list.files(path, pattern = "\\.krn$", full.names = TRUE, recursive = TRUE)
  return(krn_files)
}

krn_files <- list_krn_files(search_directory)

safe_process_krn_file <- safely(process_krn_file)
safe_results <- map(sapply(krn_files, normalizePath), safe_process_krn_file)
results <- bind_rows(map(safe_results, "result"))
#results <- bind_rows(map(krn_files, safe_process_krn_file))

#output_file <- "/Users/johnsmith/Desktop/7100/output_vocaloid.xlsx"
output_file <- "/Users/johnsmith/Desktop/7100/output_j-pop.xlsx"
write.xlsx(results, output_file)


errors <- map(results, "error")
