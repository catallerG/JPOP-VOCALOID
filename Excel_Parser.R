library(readxl)
library(dplyr)
library(openxlsx)

# 读取输入文件
input_file1 <- "/Users/johnsmith/Desktop/7100/output_j-pop.xlsx"
input_file2 <- "/Users/johnsmith/Desktop/7100/output_vocaloid.xlsx"

data1 <- read_xlsx(input_file1)
data2 <- read_xlsx(input_file2)

# 定义计算各种指标的函数
calc_metrics <- function(data) {
  avg_pitch <- mean(data$`Average Pitch`)
  avg_median_pitch <- mean(data$`Median Pitch`)
  avg_highest_pitch <- mean(data$`Highest Pitch`)
  
  avg_pitch_list <- data$`Average Pitch`
  median_pitch_list <- data$`Median Pitch`
  highest_pitch_list <- data$`Highest Pitch`
  
  avg_pitch_list_male <- data %>%
    filter(Gender == "Male") %>%
    pull(`Average Pitch`)
  
  avg_pitch_list_female <- data %>%
    filter(Gender == "Female") %>%
    pull(`Average Pitch`)
  
  avg_pitch_by_year <- data %>% 
    group_by(Year) %>%
    summarise(AvgPitch = mean(`Average Pitch`))
  
  max_abs_interval <- rowMeans(cbind(abs(data$`Positive Max Interval`), abs(data$`Minus Max Interval`)))
  avg_max_abs_interval <- mean(max_abs_interval)
  
  avg_interval <- mean(data$`Average Interval`)
  
  max_interval_list <- cbind(abs(data$`Positive Max Interval`), abs(data$`Minus Max Interval`))
  
  avg_bpm <- mean(data$BPM)
  
  bpm_list <- data$BPM
  
  avg_density_by_year <- data %>%
    group_by(Year) %>%
    summarise(AvgDensity = mean(Density))
  
  avg_bpm_by_year <- data %>%
    group_by(Year) %>%
    summarise(AvgBPM = mean(BPM))
  
  pos_list <- data$`Proportion of Skips`
  
  npm <- data$Density
  
  med_pitch_list_female <- data %>%
    filter(Gender == "Female") %>%
    pull(`Median Pitch`)
  
  highest_pitch_list_female <- data %>%
    filter(Gender == "Female") %>%
    pull(`Highest Pitch`)
  
  return(list(avg_pitch, avg_median_pitch, avg_highest_pitch, avg_pitch_by_year, avg_max_abs_interval, avg_interval, avg_bpm, avg_density_by_year, avg_bpm_by_year,
              avg_pitch_list, median_pitch_list, highest_pitch_list, avg_pitch_list_male, avg_pitch_list_female, max_interval_list, bpm_list, pos_list, npm, med_pitch_list_female, highest_pitch_list_female))
}

# 计算两个输入文件的指标
metrics1 <- calc_metrics(data1)
metrics2 <- calc_metrics(data2)

# box plot
avg_pitch_list_jpop <- metrics1[10]
avg_pitch_list_vocaloid <- metrics2[10]
boxplot(list(avg_pitch_list_jpop[[1]], avg_pitch_list_vocaloid[[1]]), names = c("J-pop", "VOCALOID"), main = "Average Pitch", ylab = "MIDI")

med_pitch_list_jpop <- metrics1[11]
med_pitch_list_vocaloid <- metrics2[11]
boxplot(list(med_pitch_list_jpop[[1]], med_pitch_list_vocaloid[[1]]), names = c("J-pop", "VOCALOID"), main = "Median Pitch", ylab = "MIDI")

highest_pitch_list_jpop <- metrics1[12]
highest_pitch_list_vocaloid <- metrics2[12]
boxplot(list(highest_pitch_list_jpop[[1]], highest_pitch_list_vocaloid[[1]]), names = c("J-pop", "VOCALOID"), main = "Highest Pitch", ylab = "MIDI")

avg_pitch_list_male_jpop <- metrics1[13]
avg_pitch_list_male_vocaloid <- metrics2[13]
boxplot(list(avg_pitch_list_male_jpop[[1]], avg_pitch_list_male_vocaloid[[1]]), names = c("J-pop", "VOCALOID"), main = "Average Pitch(Male)", ylab = "MIDI")

avg_pitch_list_female_jpop <- metrics1[14]
avg_pitch_list_female_vocaloid <- metrics2[14]
boxplot(list(avg_pitch_list_female_jpop[[1]], avg_pitch_list_female_vocaloid[[1]]), names = c("J-pop", "VOCALOID"), main = "Average Pitch(Female)", ylab = "MIDI")

#max_interval_list, bpm_list, pos_list
max_interval_list_jpop <- metrics1[15]
max_interval_list_vocaloid <- metrics2[15]
boxplot(list(max_interval_list_jpop[[1]], max_interval_list_vocaloid[[1]]), names = c("J-pop", "VOCALOID"), main = "Max Absolute Intervals", ylab = "semitones")

pos_list_jpop <- metrics1[17]
pos_list_vocaloid <- metrics2[17]
boxplot(list(pos_list_jpop[[1]], pos_list_vocaloid[[1]]), names = c("J-pop", "VOCALOID"), main = "Proportion of Skips", ylab = "porportion")


bpm_list_jpop <- metrics1[16]
bpm_list_vocaloid <- metrics2[16]
boxplot(list(bpm_list_jpop[[1]], bpm_list_vocaloid[[1]]), names = c("J-pop", "VOCALOID"), main = "BPM", ylab = "BPM")

npm_jpop <- metrics1[18]
npm_vocaloid <- metrics2[18]
boxplot(list(npm_jpop[[1]], npm_vocaloid[[1]]), names = c("J-pop", "VOCALOID"), main = "Notes per Measure", ylab = "Notes")

med_pitch_list_female_jpop <- metrics1[19]
med_pitch_list_female_vocaloid <- metrics2[19]
boxplot(list(med_pitch_list_female_jpop[[1]], med_pitch_list_female_vocaloid[[1]]), names = c("J-pop", "VOCALOID"), main = "Median Pitch(Female)", ylab = "MIDI")

highest_pitch_list_female_jpop <- metrics1[20]
highest_pitch_list_female_vocaloid <- metrics2[20]
boxplot(list(highest_pitch_list_female_jpop[[1]], highest_pitch_list_female_vocaloid[[1]]), names = c("J-pop", "VOCALOID"), main = "Highest Pitch(Female)", ylab = "MIDI")

# 创建新的.xlsx文件
output_file <- "/Users/johnsmith/Desktop/output.xlsx"
wb <- createWorkbook()

# 添加汇总结果到一个新的工作表
result_data <- tibble(Metrics = c("Avg_Pitch", "Avg_Median_Pitch", "Avg_Highest_Pitch", "Avg_Pitch_by_Year", "Avg_Max_Abs_Interval", "Avg_Interval", "Avg_BPM", "Avg_Density_by_Year", "Avg_BPM_by_Year"),
                      File1 = c(sapply(metrics1[1:7], function(x) if (is.data.frame(x)) "data.frame" else x), "data.frame", "data.frame"),
                      File2 = c(sapply(metrics2[1:7], function(x) if (is.data.frame(x)) "data.frame" else x), "data.frame", "data.frame"))

addWorksheet(wb, "Summary")
writeData(wb, "Summary", result_data)

# 将按年分组的结果追加到输出文件的其他工作表中
addWorksheet(wb, "File1_Avg_Pitch_by_Year")
writeData(wb, "File1_Avg_Pitch_by_Year", metrics1[[4]])

addWorksheet(wb, "File2_Avg_Pitch_by_Year")
writeData(wb, "File2_Avg_Pitch_by_Year", metrics2[[4]])

addWorksheet(wb, "File1_Avg_Density_by_Year")
writeData(wb, "File1_Avg_Density_by_Year", metrics1[[8]])

addWorksheet(wb, "File2_Avg_Density_by_Year")
writeData(wb, "File2_Avg_Density_by_Year", metrics2[[8]])

addWorksheet(wb, "File1_Avg_BPM_by_Year")
writeData(wb, "File1_Avg_BPM_by_Year", metrics1[[9]])

addWorksheet(wb, "File2_Avg_BPM_by_Year")
writeData(wb, "File2_Avg_BPM_by_Year", metrics2[[9]])

# 保存.xlsx文件
saveWorkbook(wb, output_file, overwrite = TRUE)
