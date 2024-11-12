
library(FactoMineR)
library(SensoMineR)
library(factoextra)
library(readxl)
library(tidyverse)
library(caret)
library(dplyr)
library(tidyr)
library(writexl)
library(rstatix)
library(readxl)
library(ggplot2)
library(ggpubr)
library(car)
library(carData)
library(tibble)
library(flextable)
library(ExPosition) 



data <- read_excel("C:/Users/LeTuanPhuc/Desktop/SPISE/template/Bangthuatngu.xlsx", 
                   sheet = "CA")
data <- as.data.frame(data)
data <- data[c(1:3),c(1:16)]
rownames(data) <- data$...1
data <- data[,-1]

Factoshiny::Factoshiny(data)

table <- aggregate(data[,c(4:ncol(data))], list(product = data[,2]), sum, na.rm =TRUE)
rownames(table) <- table[,1]
ca.ex <- epCA(data, graph = FALSE)
epGraphs(ca.ex)

## Trích xuất thông tin hàng cột để vẽ data
factor <- rbind(ca.ex$ExPosition.Data$fi, ca.ex$ExPosition.Data$fj)
contributions <- rbind(ca.ex$ExPosition.Data$ci, ca.ex$ExPosition.Data$cj)

#Vẽ dữ liệu lên 2 trục thành phần
prettyPlot(factor, 
           contributionCircles = TRUE,
           contributions=contributions,
           dev.new=F) #required so doesn't print list


res.CA<-CA(data,col.sup=c(15),graph=FALSE)
plot.CA(res.CA,cex=1.3,cex.main=1.3,cex.axis=1.3,col.col.sup='#18F051')



data <- data %>% 
    janitor::clean_names()


data <- replace(data, is.na(data), 0)

colnames(data)[1] <- ("nguoi_thu")
colnames(data)[2] <- ("san_pham")

data$nguoi_thu <- as.factor(data$nguoi_thu)
data$san_pham <- as.factor(data$san_pham)


#### Đặt lại tên sản phẩm thay cho mã hóa mẫu
data$san_pham <- gsub("628", "Trà ô long Teaplus", data$san_pham)
data$san_pham <- gsub("580", "Trà C2", data$san_pham)
data$san_pham <- gsub("309", "Trà xanh không độ", data$san_pham)
data$san_pham <- gsub("710", "CT 1", data$san_pham)
data$san_pham <- gsub("233", "CT 2", data$san_pham)


##### Test Liking
data_liking <- read_excel("C:/Users/LeTuanPhuc/Desktop/SPISE/template/Bangthuatngu.xlsx", 
                   sheet = "anova")
mean(data_liking$OKF)
sd(data_liking$OKF)


data_linking2 <- gather(data_liking, key = "product", value = "liking")

data_liking <- data[,c(1,2,47)]
data_liking <- data_liking[data_liking$san_pham != "id",]
### ANOVA
res.aov <- aov(liking ~ product, data = data_linking2)
summary(res.aov)
TukeyHSD(res.aov, ordered=TRUE)
ggbarplot(data_linking2, x = "product", y = "liking", add = "mean_sd",
          title="Điểm ưa thích chung",xlab = "Mã sản phẩm")

#############Q test###########
# Define the list of variables

data <- read_excel("D:/Project/SENSORY/SENSORY_LASUCO_2024/2024_04_19/CATA-Rút gọn thuật ngữ trà-LSC-2024_04_19.xlsx", sheet = "rutgon1")
data <- data %>% 
    janitor::clean_names()
data <- as.data.frame(data)
data <- replace(data, is.na(data), 0)
colnames(data)[1] <- ("nguoi_thu")
colnames(data)[2] <- ("san_pham")

data$nguoi_thu <- as.factor(data$nguoi_thu)
data$san_pham <- as.factor(data$san_pham)
#### Đặt lại tên sản phẩm thay cho mã hóa mẫu
data$san_pham <- gsub("628", "Trà ô long Teaplus", data$san_pham)
data$san_pham <- gsub("580", "Trà C2", data$san_pham)
data$san_pham <- gsub("309", "Trà xanh không độ", data$san_pham)
data$san_pham <- gsub("710", "CT 1", data$san_pham)
data$san_pham <- gsub("233", "CT 2", data$san_pham)

variables <- colnames(data[1,-c(1,2,3)])

# Create an empty list to store results
results <- list()

# Loop through each variable and perform Cochran's Q test
for (var in variables) {
    # Create the formula for Cochran's Q test
    formula <- as.formula(paste(var, '~ san_pham | nguoi_thu'))
    # Perform Cochran's Q test
    test_result <- cochran_qtest(formula, data = data)
    # Store the results in the list
    results[[var]] <- test_result
}
############Lưu kết quả thành dạng bảng##########
# Create an empty data frame to store the results
results_table <- data.frame(
    variable = character(),
    n = numeric(),
    statistic = numeric(),
    df = numeric(),
    p = numeric(),
    method = character(),
    stringsAsFactors = FALSE
)

# Loop through each variable and extract relevant information
for (var in variables) {
    temp_result <- results[[var]]
    result_row <- c(
        variable = var,
        n = temp_result$n,
        statistic = temp_result$statistic,
        df = temp_result$df,
        p = temp_result$p,
        method = temp_result$method
    )
    results_table <- rbind(results_table, result_row)
}

# Assign column names
colnames(results_table) <- c("Variable", "Sample_Size", "Statistic", "Degrees_of_Freedom", "p_value", "Test_Method")
# Print the results table
print(results_table)


# CA
table <- aggregate(data[,c(4:ncol(data))], list(product = data[,2]), sum, na.rm =TRUE)
rownames(table) <- table[,1]
ca.ex <- epCA(table[,-c(1)], graph = FALSE)
epGraphs(ca.ex)

## Trích xuất thông tin hàng cột để vẽ data
factor <- rbind(ca.ex$ExPosition.Data$fi, ca.ex$ExPosition.Data$fj)
contributions <- rbind(ca.ex$ExPosition.Data$ci, ca.ex$ExPosition.Data$cj)

#Vẽ dữ liệu lên 2 trục thành phần
prettyPlot(factor, 
           contributionCircles = TRUE,
           contributions=contributions,
           xlab='Component 1',
           ylab='Component 2',
           main='Component plot',
           dev.new=F) #required so doesn't print list


#Tính toán khoảng cách MDS của các tương quan thuật ngữ
mds <- cmdscale(1-cor(data[,-c(1,2,3)], use='complete.obs'))
plot(mds, type = "n", asp =1)
text(mds, rownames(mds))

## Tinhas toán penaty lift
pLift(data[,c(4)], data[,3])

#penalty-lift
library('cata')

pLift_results <- list()
# Vòng lặp qua các cột từ 3 đến 28 của bảng dữ liệu 'data'
for (i in 4:26) {
    # Tạo data2 cho cột hiện tại
    data2 <- data[, c(1, 2, i)]
    # Chuyển đổi data2
    data2 <- data2 %>%
        pivot_wider(names_from = san_pham, values_from = 3)
    data2 <- data2[,-1]  # Loại bỏ cột đầu tiên nếu không cần
    data2 <- as.matrix(data2)
    
    # Giả định data_like không thay đổi qua mỗi lần lặp
    data_like <- data[, c(1, 2, 3)]  # Cột 29 là data_like
    data_like <- data_like %>%
        pivot_wider(names_from = san_pham, values_from = 3)
    data_like <- data_like[,-1]
    data_like <- as.matrix(data_like)
    # Tính pLift
    pLift_value <- pLift(data2, data_like)
    # Lưu kết quả pLift với tên là tên cột thứ ba trong data2
    pLift_results[[names(data)[i]]] <- pLift_value
}

# Chuyển danh sách kết quả thành data frame để in ra dễ dàng
pLift_results_df <- data.frame(pLift_results)
print(pLift_results_df)


data5<- t(pLift_results_df)
data5<- as.data.frame(data5)
rownames(data5)
data5$names <- rownames(data5)
df4 <- data5[order(abs(data5$V1), decreasing = TRUE), ]

alpha_scale <- scales::rescale(abs(df4$V1)) + 0.05
# Vẽ ngang
ggplot(df4, aes(x = V1, y = reorder(names,V1))) +
    geom_bar(stat = "identity", fill = alpha("red", alpha_scale)) +
    labs(x = "Penalty-Lift Values", y = "Attribute")



### Xử lý cải tiến sản phẩm

data100 <- table

ideal.ref <- data100[data100$product == "id", 2:ncol(data100)]
CT_1.ref <- data100[data100$product == "CT 1", 2:ncol(data100)]
CT_2.ref <- data100[data100$product == "CT 2", 2:ncol(data100)]


deviation <- (CT_2.ref/30 - ideal.ref/30)
col.dev <- as.numeric(sign(deviation))
col.dev <- as.character(recode(col.dev,"1='black'; else='grey60'"))
pos <- barplot(t(deviation),xlim=c(0,32),ylim=c(-1,1), ylab="Deviation from Ideal",beside=TRUE,space=0.5,col=col.dev)
text(pos-0.4,-0.8,colnames(deviation),srt=45,cex=0.9)






# Lấy tên cột
column_names <- names(data)
# Sử dụng biểu thức chính quy để trích xuất các từ trong []
extracted_names <- gsub(".*\\[([^\\]]+)\\].*", "\\1", column_names)
