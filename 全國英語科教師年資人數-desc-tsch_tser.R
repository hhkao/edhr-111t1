rm(list=ls())

#套件名稱
packages <- c("DBI", "odbc", "magrittr", "dplyr", "rJava", "xlsx", "RStata", "readxl", "mailR", "stringr", "haven", "openxlsx", "tidyr", "gmodels")

# 安裝尚未安裝的套件
installed_packages <- packages %in% rownames(installed.packages())
if (any(installed_packages == FALSE)) {
  install.packages(packages[!installed_packages])
}

# 載入所需套件
lapply(packages, require, character.only = TRUE)

#原始資料
personnel <- readxl :: read_excel("C:/edhr-111t1/1101公立學校及1102私立學校_人事資料整合.xlsx")

#只抓專任和代理、高中部及中學部、服務身分別為教師
personnel_2 <- personnel %>%
  subset(sertype == "教師" & (emptype == "專任" | emptype == "代理") & (empunit == "高中部日間部" | empunit == "高中部進修部" | empunit == "中學部" | empunit == "中學部進修部"))

#標記正確的聘任類別
personnel_2$err_flag <- 1
personnel_2$err_flag <- if_else(grepl("英", personnel_2$emsub), 0, personnel_2$err_flag)
personnel_2$err_flag <- if_else(grepl("國（初）中", personnel_2$emsub), 1, personnel_2$err_flag)
personnel_2$err_flag <- if_else(grepl("國中", personnel_2$emsub), 1, personnel_2$err_flag)
personnel_2$err_flag <- if_else(grepl("雙語AP", personnel_2$emsub), 1, personnel_2$err_flag)

#
personnel_2 <- personnel_2 %>%
  subset(err_flag == 0)

personnel_2$survey_year <- 2022
personnel_2$onbodaty <- ""
personnel_2$onbodatm <- ""
personnel_2$onbodatd <- ""
personnel_2$onbodatd <- ""

personnel_2$onbodaty <- if_else(nchar(personnel_2$onbodat) == 6, substr(personnel_2$onbodat, 1, 2), personnel_2$onbodaty)
personnel_2$onbodatm <- if_else(nchar(personnel_2$onbodat) == 6, substr(personnel_2$onbodat, 3, 4), personnel_2$onbodatm)
personnel_2$onbodatd <- if_else(nchar(personnel_2$onbodat) == 6, substr(personnel_2$onbodat, 5, 6), personnel_2$onbodatd)
personnel_2$onbodaty <- if_else(nchar(personnel_2$onbodat) == 7, substr(personnel_2$onbodat, 1, 3), personnel_2$onbodaty)
personnel_2$onbodatm <- if_else(nchar(personnel_2$onbodat) == 7, substr(personnel_2$onbodat, 4, 5), personnel_2$onbodatm)
personnel_2$onbodatd <- if_else(nchar(personnel_2$onbodat) == 7, substr(personnel_2$onbodat, 6, 7), personnel_2$onbodatd)

personnel_2$onbodaty <- as.numeric(personnel_2$onbodaty)
personnel_2$onbodatm <- as.numeric(personnel_2$onbodatm)
personnel_2$onbodatd <- as.numeric(personnel_2$onbodatd)

#本校服務年資
personnel_2$tser <- 0
personnel_2$tser <- if_else(personnel_2$survey_year %% 4 != 0, ((personnel_2$survey_year-1911) + 7/12 + 31/365) - (personnel_2$onbodaty + (personnel_2$onbodatm/12) + (personnel_2$onbodatd/365)), personnel_2$tser)
personnel_2$tser <- if_else(personnel_2$survey_year %% 4 == 0, ((personnel_2$survey_year-1911) + 7/12 + 31/366) - (personnel_2$onbodaty + (personnel_2$onbodatm/12) + (personnel_2$onbodatd/366)), personnel_2$tser)

#本次本校任職需扣除之年資
personnel_2$desey <- substr(personnel_2$desedym, 1, 2) %>% as.numeric()
personnel_2$desem <- substr(personnel_2$desedym, 3, 4) %>% as.numeric()

personnel_2$dese <- (personnel_2$desey + (personnel_2$desem / 12))

#本校服務年資-本校任職需扣除資年資 才是實際在本校的服務年資
personnel_2$tser <- personnel_2$tser - personnel_2$dese

#避免掉年資小於零的情況（因本校到職日期+本次本校任職需扣除之年資可能為8/1的情況）
personnel_2$tser <- if_else(personnel_2$tser < 0, 0, personnel_2$tser)

#本校到職前學校服務總年資
personnel_2$beoby <- substr(personnel_2$beobdym, 1, 2) %>% as.numeric
personnel_2$beobm <- substr(personnel_2$beobdym, 3, 4) %>% as.numeric

personnel_2$beob <- (personnel_2$beoby + (personnel_2$beobm / 12))

#學校教學工作總年資
personnel_2$tsch <- personnel_2$tser + personnel_2$beob
#以上與flag24相同
  
personnel_2$tser4 <- case_when(
  (personnel_2$tser >=  0 & personnel_2$tser <=  5) ~ "5年(含)以下",
  (personnel_2$tser  >  5 & personnel_2$tser <= 10) ~ "5-10(含)年",
  (personnel_2$tser  > 10 & personnel_2$tser <= 15) ~ "10-15(含)年",
  (personnel_2$tser  > 15) ~ "15年以上"
)

personnel_2$tser4 <- personnel_2$tser4 %>% 
  factor(levels = c("5年(含)以下", "5-10(含)年", "10-15(含)年", "15年以上"))


personnel_2$tsch4 <- case_when(
  (personnel_2$tsch >=  0 & personnel_2$tsch <=  5) ~ "5年(含)以下",
  (personnel_2$tsch  >  5 & personnel_2$tsch <= 10) ~ "5-10(含)年",
  (personnel_2$tsch  > 10 & personnel_2$tsch <= 15) ~ "10-15(含)年",
  (personnel_2$tsch  > 15) ~ "15年以上"
)
personnel_2$tsch4 <- personnel_2$tsch4 %>% 
  factor(levels = c("5年(含)以下", "5-10(含)年", "10-15(含)年", "15年以上"))

personnel_2$emptype <- personnel_2$emptype %>% 
  factor(levels = c("專任", "代理"))

table(personnel_2$tser4, personnel_2$emptype)
CrossTable(personnel_2$emptype, personnel_2$tser4, digits=1, max.width = 5, expected=FALSE, prop.r=TRUE, prop.c=FALSE, prop.t=FALSE, prop.chisq=FALSE,
           format=c("SPSS"))

table(personnel_2$tsch4, personnel_2$emptype)
CrossTable(personnel_2$tsch4, personnel_2$emptype, digits=1, max.width = 5, expected=FALSE, prop.r=TRUE, prop.c=FALSE, prop.t=FALSE, prop.chisq=FALSE,
           format=c("SPSS"))

openxlsx :: write.xlsx(personnel_2, file = paste("C:/edhr-111t1/", "全國英語科教師年資人數_R", ".xlsx", sep = ""), rowNames = FALSE, overwrite = TRUE)


#0830

#tser4本校服務年資，專任
emptype1 <- personnel_2 %>% subset(emptype == "專任")
emptype1_table <- table(emptype1$organization_id, emptype1$tser4)
openxlsx :: write.xlsx(emptype1_table, file = paste("C:/edhr-111t1/", "tser4_emptype1_table", ".xlsx", sep = ""), rowNames = FALSE, overwrite = TRUE)

#tser4本校服務年資，代理
emptype2 <- personnel_2 %>% subset(emptype == "代理")
emptype2_table <- table(emptype2$organization_id, emptype2$tser4)
openxlsx :: write.xlsx(emptype2_table, file = paste("C:/edhr-111t1/", "tser4_emptype2_table", ".xlsx", sep = ""), rowNames = FALSE, overwrite = TRUE)

#tser4本校服務年資，專任及代理
emptype1_2 <- personnel_2
emptype1_2_table <- table(emptype1_2$organization_id, emptype1_2$tser4)
openxlsx :: write.xlsx(emptype1_2_table, file = paste("C:/edhr-111t1/", "tser4_emptype1_2_table", ".xlsx", sep = ""), rowNames = FALSE, overwrite = TRUE)



#tsch4本校服務年資，專任
emptype1 <- personnel_2 %>% subset(emptype == "專任")
emptype1_table <- table(emptype1$organization_id, emptype1$tsch4)
openxlsx :: write.xlsx(emptype1_table, file = paste("C:/edhr-111t1/", "tsch4_emptype1_table", ".xlsx", sep = ""), rowNames = FALSE, overwrite = TRUE)

#tsch4本校服務年資，代理
emptype2 <- personnel_2 %>% subset(emptype == "代理")
emptype2_table <- table(emptype2$organization_id, emptype2$tsch4)
openxlsx :: write.xlsx(emptype2_table, file = paste("C:/edhr-111t1/", "tsch4_emptype2_table", ".xlsx", sep = ""), rowNames = FALSE, overwrite = TRUE)

#tsch4本校服務年資，專任及代理
emptype1_2 <- personnel_2
emptype1_2_table <- table(emptype1_2$organization_id, emptype1_2$tsch4)
openxlsx :: write.xlsx(emptype1_2_table, file = paste("C:/edhr-111t1/", "tsch4_emptype1_2_table", ".xlsx", sep = ""), rowNames = FALSE, overwrite = TRUE)
