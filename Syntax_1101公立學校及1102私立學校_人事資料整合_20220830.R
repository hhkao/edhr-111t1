rm(list=ls())

#套件名稱
packages <- c("DBI", "odbc", "magrittr", "dplyr", "rJava", "xlsx", "RStata", "readxl", "mailR", "stringr", "haven", "openxlsx")

# 安裝尚未安裝的套件
installed_packages <- packages %in% rownames(installed.packages())
if (any(installed_packages == FALSE)) {
  install.packages(packages[!installed_packages])
}

# 載入所需套件
lapply(packages, require, character.only = TRUE)

#資料讀取#
edhr <- dbConnect(odbc::odbc(), "CHER01-EDHR-NEW", timeout = 10)


# 整合20校試辦中的公立學校、所有公立、私立教員資料表及職員(工)資料表

##### 1101 20校試辦 教員資料表#####
#請輸入本次填報設定檔標題(字串需與標題完全相符，否則會找不到)
title <- "110學年度上學期高級中等學校教育人力資源資料庫（20校人事及教務）"

department <- "人事室"

#讀取審核同意之學校名單
list_agree <- dbGetQuery(edhr, 
                         paste("
SELECT DISTINCT b.id AS organization_id , 1 AS agree
FROM [plat5_edhr].[dbo].[teacher_fillers] a 
LEFT JOIN 
(SELECT a.reporter_id, c.id
FROM [plat5_edhr].[dbo].[teacher_fillers] a LEFT JOIN [plat5_edhr].[dbo].[teacher_reporters] b ON a.reporter_id = b.id
LEFT JOIN [plat5_edhr].[dbo].[organization_details] c ON b.organization_id = c.organization_id
) b ON a.reporter_id = b.reporter_id
WHERE a.agree = 1 AND department_id IN (SELECT id FROM [plat5_edhr].[dbo].[teacher_departments]
                                        WHERE report_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_reports]
                                                            WHERE title = '", title, "'))", sep = "")
) %>%
  distinct(organization_id, .keep_all = TRUE)

#讀取教員資料表名稱
teacher_tablename <- dbGetQuery(edhr, 
                                paste("
SELECT [name] FROM [plat5_edhr].[dbo].[row_tables] 
	where sheet_id = (SELECT [id] FROM [plat5_edhr].[dbo].[row_sheets] 
						          where file_id = (SELECT field_component_id FROM [plat5_edhr].[dbo].[teacher_datasets] 
											                  WHERE title = '教員資料表' AND department_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_departments] 
																						                                              WHERE title = '", department, "' AND  report_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_reports] 
																												                                                                                      WHERE title = '", title, "'))))", sep = "")
) %>% as.character()

#讀取教員資料表
teacher <- dbGetQuery(edhr, 
                      paste("SELECT * FROM [rows].[dbo].[", teacher_tablename, "] WHERE deleted_at IS NULL", sep = "")
) %>%
  subset(select = -c(id, created_at, deleted_at, updated_by, created_by, deleted_by))

#欄位名稱更改為設定的欄位代號
col_names <- dbGetQuery(edhr, "SELECT id, name, title FROM [plat5_edhr].[dbo].[row_columns]")
col_names$id <- paste("C", col_names$id, sep = "")
for (i in 2 : dim(teacher)[2]) #從2開始是因為第一的欄位是update_at
{
  colnames(teacher)[i] <- col_names$name[grep(paste(colnames(teacher)[i], "$", sep = ""), col_names$id)]
}
#格式調整
teacher$gender <- formatC(teacher$gender, dig = 0, wid = 1, format = "f", flag = "0")
teacher$birthdate <- formatC(teacher$birthdate, dig = 0, wid = 7, format = "f", flag = "0")
teacher$onbodat <- formatC(teacher$onbodat, dig = 0, wid = 7, format = "f", flag = "0")
teacher$desedym <- formatC(teacher$desedym, dig = 0, wid = 4, format = "f", flag = "0")
teacher$beobdym <- formatC(teacher$beobdym, dig = 0, wid = 4, format = "f", flag = "0")
teacher$organization_id <- formatC(teacher$organization_id, dig = 0, wid = 6, format = "f", flag = "0")

#只留下審核通過之名單 只留下公立
teacher <- merge(x = teacher, y = list_agree, by = "organization_id", all.x = TRUE) %>%
  subset(agree == 1) %>%
  subset(select = -c(updated_at, agree)) %>%
  subset(substr(teacher$organization_id, 3, 3) == "0" | substr(teacher$organization_id, 3, 3) == "3" | substr(teacher$organization_id, 3, 3) == "4")

teacher20_1101 <- teacher_1101 %>%
  mutate(dta_teacher = "教員資料表")

##### 1101公立學校 職員(工)資料表#####
#讀取職員(工)資料表名稱
staff_tablename <- dbGetQuery(edhr, 
                              paste("
SELECT [name] FROM [plat5_edhr].[dbo].[row_tables] 
	where sheet_id = (SELECT [id] FROM [plat5_edhr].[dbo].[row_sheets] 
						          where file_id = (SELECT field_component_id FROM [plat5_edhr].[dbo].[teacher_datasets] 
											                   WHERE title = '職員(工)資料表' AND department_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_departments] 
																							                                                 WHERE title = '", department, "' AND  report_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_reports] 
																												                                                            WHERE title = '", title, "'))))", sep = "")
) %>% as.character()

#讀取職員(工)資料表
staff <- dbGetQuery(edhr, 
                    paste("SELECT * FROM [rows].[dbo].[", staff_tablename, "] WHERE deleted_at IS NULL", sep = "")
) %>%
  subset(select = -c(id, created_at, deleted_at, updated_by, created_by, deleted_by))
#欄位名稱更改為設定的欄位代號
for (i in 2 : dim(staff)[2]) #從2開始是因為第一的欄位是update_at
{
  colnames(staff)[i] <- col_names$name[grep(paste(colnames(staff)[i], "$", sep = ""), col_names$id)]
}

#格式調整
staff$gender <- formatC(staff$gender, dig = 0, wid = 1, format = "f", flag = "0")
staff$birthdate <- formatC(staff$birthdate, dig = 0, wid = 7, format = "f", flag = "0")
staff$onbodat <- formatC(staff$onbodat, dig = 0, wid = 7, format = "f", flag = "0")
staff$desedym <- formatC(staff$desedym, dig = 0, wid = 4, format = "f", flag = "0")
staff$beobdym <- formatC(staff$beobdym, dig = 0, wid = 4, format = "f", flag = "0")
staff$organization_id <- formatC(staff$organization_id, dig = 0, wid = 6, format = "f", flag = "0")

#只留下審核通過之名單
staff <- merge(x = staff, y = list_agree, by = "organization_id", all.x = TRUE) %>%
  subset(agree == 1) %>%
  subset(select = -c(updated_at, agree)) %>%
  subset(substr(teacher$organization_id, 3, 3) == "0" | substr(teacher$organization_id, 3, 3) == "3" | substr(teacher$organization_id, 3, 3) == "4")

staff20_1102 <- staff %>%
  mutate(dta_teacher = "職員(工)資料表")

##### 1101公立學校 教員資料表#####
#請輸入本次填報設定檔標題(字串需與標題完全相符，否則會找不到)
title <- "110學年度上學期高級中等學校教育人力資源資料庫（公立學校人事）"

department <- "人事室"

#讀取審核同意之學校名單
list_agree <- dbGetQuery(edhr, 
                         paste("
SELECT DISTINCT b.id AS organization_id , 1 AS agree
FROM [plat5_edhr].[dbo].[teacher_fillers] a 
LEFT JOIN 
(SELECT a.reporter_id, c.id
FROM [plat5_edhr].[dbo].[teacher_fillers] a LEFT JOIN [plat5_edhr].[dbo].[teacher_reporters] b ON a.reporter_id = b.id
LEFT JOIN [plat5_edhr].[dbo].[organization_details] c ON b.organization_id = c.organization_id
) b ON a.reporter_id = b.reporter_id
WHERE a.agree = 1 AND department_id IN (SELECT id FROM [plat5_edhr].[dbo].[teacher_departments]
                                        WHERE report_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_reports]
                                                            WHERE title = '", title, "'))", sep = "")
) %>%
  distinct(organization_id, .keep_all = TRUE)

#讀取教員資料表名稱
teacher_tablename <- dbGetQuery(edhr, 
                                paste("
SELECT [name] FROM [plat5_edhr].[dbo].[row_tables] 
	where sheet_id = (SELECT [id] FROM [plat5_edhr].[dbo].[row_sheets] 
						          where file_id = (SELECT field_component_id FROM [plat5_edhr].[dbo].[teacher_datasets] 
											                  WHERE title = '教員資料表' AND department_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_departments] 
																						                                              WHERE title = '", department, "' AND  report_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_reports] 
																												                                                                                      WHERE title = '", title, "'))))", sep = "")
) %>% as.character()

#讀取教員資料表
teacher <- dbGetQuery(edhr, 
                      paste("SELECT * FROM [rows].[dbo].[", teacher_tablename, "] WHERE deleted_at IS NULL", sep = "")
) %>%
  subset(select = -c(id, created_at, deleted_at, updated_by, created_by, deleted_by))

#欄位名稱更改為設定的欄位代號
col_names <- dbGetQuery(edhr, "SELECT id, name, title FROM [plat5_edhr].[dbo].[row_columns]")
col_names$id <- paste("C", col_names$id, sep = "")
for (i in 2 : dim(teacher)[2]) #從2開始是因為第一的欄位是update_at
{
  colnames(teacher)[i] <- col_names$name[grep(paste(colnames(teacher)[i], "$", sep = ""), col_names$id)]
}
#格式調整
teacher$gender <- formatC(teacher$gender, dig = 0, wid = 1, format = "f", flag = "0")
teacher$birthdate <- formatC(teacher$birthdate, dig = 0, wid = 7, format = "f", flag = "0")
teacher$onbodat <- formatC(teacher$onbodat, dig = 0, wid = 7, format = "f", flag = "0")
teacher$desedym <- formatC(teacher$desedym, dig = 0, wid = 4, format = "f", flag = "0")
teacher$beobdym <- formatC(teacher$beobdym, dig = 0, wid = 4, format = "f", flag = "0")
teacher$organization_id <- formatC(teacher$organization_id, dig = 0, wid = 6, format = "f", flag = "0")

#只留下審核通過之名單
teacher <- merge(x = teacher, y = list_agree, by = "organization_id", all.x = TRUE) %>%
  subset(agree == 1) %>%
  subset(select = -c(updated_at, agree))

teacher_1101 <- teacher %>%
  mutate(dta_teacher = "教員資料表")

##### 1101公立學校 職員(工)資料表#####
#讀取職員(工)資料表名稱
staff_tablename <- dbGetQuery(edhr, 
                              paste("
SELECT [name] FROM [plat5_edhr].[dbo].[row_tables] 
	where sheet_id = (SELECT [id] FROM [plat5_edhr].[dbo].[row_sheets] 
						          where file_id = (SELECT field_component_id FROM [plat5_edhr].[dbo].[teacher_datasets] 
											                   WHERE title = '職員(工)資料表' AND department_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_departments] 
																							                                                 WHERE title = '", department, "' AND  report_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_reports] 
																												                                                            WHERE title = '", title, "'))))", sep = "")
) %>% as.character()

#讀取職員(工)資料表
staff <- dbGetQuery(edhr, 
                    paste("SELECT * FROM [rows].[dbo].[", staff_tablename, "] WHERE deleted_at IS NULL", sep = "")
) %>%
  subset(select = -c(id, created_at, deleted_at, updated_by, created_by, deleted_by))
#欄位名稱更改為設定的欄位代號
for (i in 2 : dim(staff)[2]) #從2開始是因為第一的欄位是update_at
{
  colnames(staff)[i] <- col_names$name[grep(paste(colnames(staff)[i], "$", sep = ""), col_names$id)]
}

#格式調整
staff$gender <- formatC(staff$gender, dig = 0, wid = 1, format = "f", flag = "0")
staff$birthdate <- formatC(staff$birthdate, dig = 0, wid = 7, format = "f", flag = "0")
staff$onbodat <- formatC(staff$onbodat, dig = 0, wid = 7, format = "f", flag = "0")
staff$desedym <- formatC(staff$desedym, dig = 0, wid = 4, format = "f", flag = "0")
staff$beobdym <- formatC(staff$beobdym, dig = 0, wid = 4, format = "f", flag = "0")
staff$organization_id <- formatC(staff$organization_id, dig = 0, wid = 6, format = "f", flag = "0")

#只留下審核通過之名單
staff <- merge(x = staff, y = list_agree, by = "organization_id", all.x = TRUE) %>%
  subset(agree == 1) %>%
  subset(select = -c(updated_at, agree))

staff_1102 <- staff %>%
  mutate(dta_teacher = "職員(工)資料表")

##### 1102私立學校 教員資料表#####
#請輸入本次填報設定檔標題(字串需與標題完全相符，否則會找不到)
title <- "110學年度下學期高級中等學校教育人力資源資料庫（私立學校人事）"

department <- "人事室"

#讀取審核同意之學校名單
list_agree <- dbGetQuery(edhr, 
                         paste("
SELECT DISTINCT b.id AS organization_id , 1 AS agree
FROM [plat5_edhr].[dbo].[teacher_fillers] a 
LEFT JOIN 
(SELECT a.reporter_id, c.id
FROM [plat5_edhr].[dbo].[teacher_fillers] a LEFT JOIN [plat5_edhr].[dbo].[teacher_reporters] b ON a.reporter_id = b.id
LEFT JOIN [plat5_edhr].[dbo].[organization_details] c ON b.organization_id = c.organization_id
) b ON a.reporter_id = b.reporter_id
WHERE a.agree = 1 AND department_id IN (SELECT id FROM [plat5_edhr].[dbo].[teacher_departments]
                                        WHERE report_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_reports]
                                                            WHERE title = '", title, "'))", sep = "")
) %>%
  distinct(organization_id, .keep_all = TRUE)

#讀取教員資料表名稱
teacher_tablename <- dbGetQuery(edhr, 
                                paste("
SELECT [name] FROM [plat5_edhr].[dbo].[row_tables] 
	where sheet_id = (SELECT [id] FROM [plat5_edhr].[dbo].[row_sheets] 
						          where file_id = (SELECT field_component_id FROM [plat5_edhr].[dbo].[teacher_datasets] 
											                  WHERE title = '教員資料表' AND department_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_departments] 
																						                                              WHERE title = '", department, "' AND  report_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_reports] 
																												                                                                                      WHERE title = '", title, "'))))", sep = "")
) %>% as.character()

#讀取教員資料表
teacher <- dbGetQuery(edhr, 
                      paste("SELECT * FROM [rows].[dbo].[", teacher_tablename, "] WHERE deleted_at IS NULL", sep = "")
) %>%
  subset(select = -c(id, created_at, deleted_at, updated_by, created_by, deleted_by))

#欄位名稱更改為設定的欄位代號
col_names <- dbGetQuery(edhr, "SELECT id, name, title FROM [plat5_edhr].[dbo].[row_columns]")
col_names$id <- paste("C", col_names$id, sep = "")
for (i in 2 : dim(teacher)[2]) #從2開始是因為第一的欄位是update_at
{
  colnames(teacher)[i] <- col_names$name[grep(paste(colnames(teacher)[i], "$", sep = ""), col_names$id)]
}
#格式調整
teacher$gender <- formatC(teacher$gender, dig = 0, wid = 1, format = "f", flag = "0")
teacher$birthdate <- formatC(teacher$birthdate, dig = 0, wid = 7, format = "f", flag = "0")
teacher$onbodat <- formatC(teacher$onbodat, dig = 0, wid = 7, format = "f", flag = "0")
teacher$desedym <- formatC(teacher$desedym, dig = 0, wid = 4, format = "f", flag = "0")
teacher$beobdym <- formatC(teacher$beobdym, dig = 0, wid = 4, format = "f", flag = "0")
teacher$organization_id <- formatC(teacher$organization_id, dig = 0, wid = 6, format = "f", flag = "0")

#只留下審核通過之名單
teacher <- merge(x = teacher, y = list_agree, by = "organization_id", all.x = TRUE) %>%
  subset(agree == 1) %>%
  subset(select = -c(updated_at, agree))

teacher_1102 <- teacher %>%
  mutate(dta_teacher = "教員資料表")

##### 1102私立學校 職員(工)資料表#####
#讀取職員(工)資料表名稱
staff_tablename <- dbGetQuery(edhr, 
                              paste("
SELECT [name] FROM [plat5_edhr].[dbo].[row_tables] 
	where sheet_id = (SELECT [id] FROM [plat5_edhr].[dbo].[row_sheets] 
						          where file_id = (SELECT field_component_id FROM [plat5_edhr].[dbo].[teacher_datasets] 
											                   WHERE title = '職員(工)資料表' AND department_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_departments] 
																							                                                 WHERE title = '", department, "' AND  report_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_reports] 
																												                                                            WHERE title = '", title, "'))))", sep = "")
) %>% as.character()

#讀取職員(工)資料表
staff <- dbGetQuery(edhr, 
                    paste("SELECT * FROM [rows].[dbo].[", staff_tablename, "] WHERE deleted_at IS NULL", sep = "")
) %>%
  subset(select = -c(id, created_at, deleted_at, updated_by, created_by, deleted_by))
#欄位名稱更改為設定的欄位代號
for (i in 2 : dim(staff)[2]) #從2開始是因為第一的欄位是update_at
{
  colnames(staff)[i] <- col_names$name[grep(paste(colnames(staff)[i], "$", sep = ""), col_names$id)]
}

#格式調整
staff$gender <- formatC(staff$gender, dig = 0, wid = 1, format = "f", flag = "0")
staff$birthdate <- formatC(staff$birthdate, dig = 0, wid = 7, format = "f", flag = "0")
staff$onbodat <- formatC(staff$onbodat, dig = 0, wid = 7, format = "f", flag = "0")
staff$desedym <- formatC(staff$desedym, dig = 0, wid = 4, format = "f", flag = "0")
staff$beobdym <- formatC(staff$beobdym, dig = 0, wid = 4, format = "f", flag = "0")
staff$organization_id <- formatC(staff$organization_id, dig = 0, wid = 6, format = "f", flag = "0")

#只留下審核通過之名單
staff <- merge(x = staff, y = list_agree, by = "organization_id", all.x = TRUE) %>%
  subset(agree == 1) %>%
  subset(select = -c(updated_at, agree))

staff_1102 <- staff %>%
  mutate(dta_teacher = "職員(工)資料表")

#####合併#####
personnel_1101_1102 <- bind_rows(teacher_1101, staff_1101, teacher_1102, staff_1102)
openxlsx :: write.xlsx(personnel_1101_1102, file = paste("C:/R/", "1101公立學校及1102私立學校_人事資料整合", ".xlsx", sep = ""), rowNames = FALSE, overwrite = TRUE)
