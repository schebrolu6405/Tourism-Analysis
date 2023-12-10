install.packages("readxl","gridExtra")
setwd("C:/Users/sushe/Documents/BDA_Fall_2023/Sem1/BDA_592/Project/Datasets/Cleaned Datasets")

sheets <- readxl::excel_sheets("Tour2016-21AU22.xlsx")
for (sheet in sheets) {
  df <- readxl::read_excel("Tour2016-21AU22.xlsx", sheet = sheet)
  print(df)
}
#################sheet1##########################
data <- readxl::read_excel("Tour2016-21AU22.xlsx", sheet = 1)

#variation
library(ggplot2)
library(gridExtra)
plot_hist <- ggplot(data, aes(x=`Air transportation`)) + geom_histogram(bins = 30, fill = "yellow", color = "black", alpha = 0.7) +
  labs(title = "Histogram of Air Transportation", x = "Air Transportation", y = "Frequency")

plot_box <- ggplot(data, aes(x = factor(Year), y = `Rail transportation`)) +
geom_boxplot(fill = "skyblue", color = "black", alpha = 0.7) +
labs(title = "Box Plot of Rail Transportation Over Years", x = "Year", y = "Rail Transportation")

plot_scatter <- ggplot(data, aes(x = Year, y = `Water transportation`)) +
  geom_point(color = "red") +
  labs(title = "Line Plot of Water Transportation Over Years", x = "Year", y = "Water Transportation")

grid.arrange(plot_hist, plot_box, plot_scatter, ncol=3)

##################sheet2###############################
data2 <- readxl::read_excel("Tour2016-21AU22.xlsx", sheet = 2)

# Assuming 'data2' is your data frame
library(dplyr)
library(tidyr)

selected_data <- data2 %>% select(Year, Imports, `Exports of goods and services`, `Domestic production at producers' prices`,`Government expenditures`,`Private expenditures`)

wide_data <- selected_data %>%
  pivot_wider(names_from = Year, values_from = c(Imports, `Exports of goods and services`,`Domestic production at producers' prices`,`Government expenditures`,`Private expenditures`))

# Calculate the correlation matrix for the numeric columns
correlation_matrix <- cor(selected_data[, c("Imports", "Exports of goods and services","Domestic production at producers' prices","Government expenditures","Private expenditures")])

# Print the correlation matrix
print(correlation_matrix)

# Create a data frame for heatmap
correlation_df <- data.frame(
  Var1 = rep(colnames(correlation_matrix), each = ncol(correlation_matrix)),
  Var2 = rep(colnames(correlation_matrix), times = ncol(correlation_matrix)),
  value = as.vector(correlation_matrix)
)

# Plot the heatmap
ggplot(data = correlation_df,
       aes(x = Var1, y = Var2, fill = value)) +
  geom_tile() +
  scale_fill_gradient(low = "white", high = "blue") +
  labs(title = "Correlation Heatmap",
       x = "Variable",
       y = "Variable")

##################sheet3######################
data3 <- readxl::read_excel("Tour2016-21AU22.xlsx", sheet = 3)
# t-Test between Resident and Nonresidents 
print(sum(is.na(data3$`Resident households`)))
print(sum(is.na(data3$Nonresidents)))
data3$Resident_households <- as.numeric(data3$`Resident households`)
data3$Nonresidents <- as.numeric(data3$Nonresidents)
data3 <- na.omit(data3)
group1 <- as.numeric(data3$`Resident households`)
group2 <- as.numeric(data3$Nonresidents)
t_test_result <- t.test(group1, group2)
print(t_test_result)
response_variable <- c(data3$Business, data3$Government)
group_variable <- rep(c("Group1", "Group2"), each = nrow(data3)) 
# Perform Analysis of Variance
anova_result <- aov(response_variable ~ group_variable)
print(summary(anova_result))
data3v <- data.frame(data3$`Resident households`, data3$Nonresidents, data3$Business, data3$Government)
data_long <- gather(data3v, key = "Variable", value = "Value")
# Create a horizontal bar plot
ggplot(data_long, aes(x = Value, y = factor(Variable), fill = Variable)) +
  geom_bar(stat = "identity", position = "dodge") +
  labs(title = "Bar Plot for various types of visitors", x = "Value", y = "Variable") +
  theme_minimal() +
  theme(legend.position = "bottom")

####################Sheet4##################################
data4 <- readxl::read_excel("Tour2016-21AU22.xlsx", sheet = 5)
contingency_table1 <- table(data4$Industry, data4$`Tourism intermediate consumption`)
contingency_table2 <- table(data4$Industry, data4$`Intermediate consumption`)
# Perform chi-square test
chi_square_result1 <- chisq.test(contingency_table1)
chi_square_result2 <- chisq.test(contingency_table2)
print(chi_square_result1)
print(chi_square_result2)
data4 <- na.omit(data4)
# Grouped bar plot
data4_sorted <- data4[order(-data4$`Tourism intermediate consumption`), ]
# Select the top 15 rows
data4_top15 <- head(data4_sorted, 15)
ggplot(data4_top15, aes(x = `Intermediate consumption`, y = `Tourism intermediate consumption`, color = Industry)) +
  geom_point(size = 3) +
  labs(title = "Scatter Plot of Consumption by Industry", x = "Intermediate Consumption", y = "Tourism Intermediate Consumption", color = "Industry")

##########################Sheet5##########################
data5 <- readxl::read_excel("Tour2016-21AU22.xlsx", sheet = 7)
data5_sorted <- data5[order(-data5$`Total employment (thousands of employees)`), ]

# Select the top 15 industries
top15_data <- head(data5_sorted, 15)
data_cluster <- top15_data[, c("Total employment (thousands of employees)", "Compensation (millions of dollars)", "Tourism industry ratio")]
data_cluster_standardized <- scale(data_cluster)

#Hierarchical clustering
dist_matrix <- dist(data_cluster_standardized)
hclust_result <- hclust(dist_matrix)

#Plot dendrogram
dendrogram <- as.dendrogram(hclust_result)
attr(dendrogram, "labels_colors") <- c("red", "green", "blue")  # Customize label colors if needed
plot(dendrogram, main = "Hierarchical Clustering Dendrogram for Top 15 Industries", horiz = TRUE)

############sheet6#############
data6 <- readxl::read_excel("Tour2016-21AU22.xlsx", sheet = 8)
numeric_data <- data6[, sapply(data6, is.numeric)]
summary_statistics <- aggregate(. ~ Year, data = numeric_data, function(x) c(Mean = mean(x), Median = median(x), Min = min(x), Max = max(x)))
print(summary_statistics)

#############sheet7#####################
data7 <- readxl::read_excel("Tour2016-21AU22.xlsx", sheet = 9)
# Fit linear regression model
model <- lm(data7$`Direct output (Millions of dollars)` ~ data7$`Real output (Millions of chained (2012) dollars)` + data7$`Chain-type price index`, data = data7)

summary(model)

plot(data7$`Direct output (Millions of dollars)`, data7$`Real output (Millions of chained (2012) dollars)`, main = "Multivariable Analysis", xlab = "Direct Output", ylab = "Real Output")
abline(model, col = "red")
