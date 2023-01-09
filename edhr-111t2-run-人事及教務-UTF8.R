# 若自動化檢誤執行有問題，要先把下面兩個檔案刪掉，再重跑
# pre_correct_list.xlsx
# pre_list_agree.xlsx

rm(list=ls())

#套件名稱
packages <- c("DBI", "odbc", "magrittr", "dplyr", "rJava", "xlsx", "RStata", "readxl", "mailR", "stringr", "haven", "openxlsx", "tidyr", "maditr")

# 安裝尚未安裝的套件
installed_packages <- packages %in% rownames(installed.packages())
if (any(installed_packages == FALSE)) {
  install.packages(packages[!installed_packages])
}

# 載入所需套件
lapply(packages, require, character.only = TRUE)

# 匯入學校資料檔 -------------------------------------------------------------------
# input data
# 分頁名稱為系統指定。

#資料讀取#
edhr <- dbConnect(odbc::odbc(), "CHER01-EDHR-NEW", timeout = 10)

#請輸入本次填報設定檔標題(字串需與標題完全相符，否則會找不到)
title <- "111學年度上學期高級中等學校教育人力資源資料庫（公立學校人事及教務）"

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
  distinct(organization_id, .keep_all = TRUE) %>%
  subset(organization_id == "421302" | organization_id == "381302" | organization_id == "331404" | organization_id == "421303" | organization_id == "411301" | organization_id == "401302" | organization_id == "351402" | organization_id == "381305" | organization_id == "411401" | organization_id == "421301" | organization_id == "361B09" | organization_id == "331403" | organization_id == "331402" | organization_id == "341302" | organization_id == "411302" | organization_id == "401303" | organization_id == "411303" | organization_id == "401301" | organization_id == "381306" | organization_id == "381304" | organization_id == "381303" | organization_id == "331301" | organization_id == "341402" | organization_id == "361401" | organization_id == "311401" | organization_id == "331304" | organization_id == "351B09" | organization_id == "331302" | organization_id == "361301" | organization_id == "351301" | organization_id == "421404" | organization_id == "321399" | organization_id == "381301" | organization_id == "581301" | organization_id == "121318" | organization_id == "581401" | organization_id == "121417" | organization_id == "521303" | organization_id == "581302" | organization_id == "581402" | organization_id == "521301" | organization_id == "551303" | organization_id == "121307" | organization_id == "121410" | organization_id == "121302" | organization_id == "121415" | organization_id == "551402" | organization_id == "521401" | organization_id == "551301" | organization_id == "121413" | organization_id == "121405" | organization_id == "121320" | organization_id == "121306" | organization_id == "101303" | organization_id == "101304" | organization_id == "101406" | organization_id == "101405" | organization_id == "141406" | organization_id == "141307" | organization_id == "141301" | organization_id == "181305" | organization_id == "181307" | organization_id == "181308" | organization_id == "181306" | organization_id == "201314" | organization_id == "201313" | organization_id == "201304" | organization_id == "201408" | organization_id == "201309" | organization_id == "201312" | organization_id == "201310")


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

#管區
teacher <- teacher %>%
  #subset(organization_id == "421302" | organization_id == "381302" | organization_id == "331404" | organization_id == "421303" | organization_id == "411301" | organization_id == "401302" | organization_id == "351402" | organization_id == "381305" | organization_id == "411401" | organization_id == "421301" | organization_id == "361B09" | organization_id == "331403" | organization_id == "331402" | organization_id == "341302" | organization_id == "411302" | organization_id == "401303" | organization_id == "411303" | organization_id == "401301" | organization_id == "381306" | organization_id == "381304" | organization_id == "381303" | organization_id == "331301" | organization_id == "341402" | organization_id == "361401" | organization_id == "311401" | organization_id == "331304" | organization_id == "351B09" | organization_id == "331302" | organization_id == "361301" | organization_id == "351301" | organization_id == "421404" | organization_id == "321399" | organization_id == "381301" | organization_id == "581301" | organization_id == "121318" | organization_id == "581401" | organization_id == "121417" | organization_id == "521303" | organization_id == "581302" | organization_id == "581402" | organization_id == "521301" | organization_id == "551303" | organization_id == "121307" | organization_id == "121410" | organization_id == "121302" | organization_id == "121415" | organization_id == "551402" | organization_id == "521401" | organization_id == "551301" | organization_id == "121413" | organization_id == "121405" | organization_id == "121320" | organization_id == "121306" | organization_id == "101303" | organization_id == "101304" | organization_id == "101406" | organization_id == "101405" | organization_id == "141406" | organization_id == "141307" | organization_id == "141301" | organization_id == "181305" | organization_id == "181307" | organization_id == "181308" | organization_id == "181306" | organization_id == "201314" | organization_id == "201313" | organization_id == "201304" | organization_id == "201408" | organization_id == "201309" | organization_id == "201312" | organization_id == "201310")
  
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

#管區
staff <- staff %>%
  #subset(organization_id == "421302" | organization_id == "381302" | organization_id == "331404" | organization_id == "421303" | organization_id == "411301" | organization_id == "401302" | organization_id == "351402" | organization_id == "381305" | organization_id == "411401" | organization_id == "421301" | organization_id == "361B09" | organization_id == "331403" | organization_id == "331402" | organization_id == "341302" | organization_id == "411302" | organization_id == "401303" | organization_id == "411303" | organization_id == "401301" | organization_id == "381306" | organization_id == "381304" | organization_id == "381303" | organization_id == "331301" | organization_id == "341402" | organization_id == "361401" | organization_id == "311401" | organization_id == "331304" | organization_id == "351B09" | organization_id == "331302" | organization_id == "361301" | organization_id == "351301" | organization_id == "421404" | organization_id == "321399" | organization_id == "381301" | organization_id == "581301" | organization_id == "121318" | organization_id == "581401" | organization_id == "121417" | organization_id == "521303" | organization_id == "581302" | organization_id == "581402" | organization_id == "521301" | organization_id == "551303" | organization_id == "121307" | organization_id == "121410" | organization_id == "121302" | organization_id == "121415" | organization_id == "551402" | organization_id == "521401" | organization_id == "551301" | organization_id == "121413" | organization_id == "121405" | organization_id == "121320" | organization_id == "121306" | organization_id == "101303" | organization_id == "101304" | organization_id == "101406" | organization_id == "101405" | organization_id == "141406" | organization_id == "141307" | organization_id == "141301" | organization_id == "181305" | organization_id == "181307" | organization_id == "181308" | organization_id == "181306" | organization_id == "201314" | organization_id == "201313" | organization_id == "201304" | organization_id == "201408" | organization_id == "201309" | organization_id == "201312" | organization_id == "201310")
  
  department <- "教務處"

#讀取教學資料表名稱
load_tablename <- dbGetQuery(edhr, 
                             paste("
SELECT [name] FROM [plat5_edhr].[dbo].[row_tables] 
	where sheet_id = (SELECT [id] FROM [plat5_edhr].[dbo].[row_sheets] 
						          where file_id = (SELECT field_component_id FROM [plat5_edhr].[dbo].[teacher_datasets] 
											                  WHERE title = '教學資料表' AND department_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_departments] 
																						                                              WHERE title = '", department, "' AND  report_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_reports] 
																												                                                       WHERE title = '", title, "'))))", sep = "")
) %>% as.character()

#讀取教學資料表
load <- dbGetQuery(edhr, 
                   paste("SELECT * FROM [rows].[dbo].[", load_tablename, "] WHERE deleted_at IS NULL", sep = "")
) %>%
  subset(select = -c(id, created_at, deleted_at, updated_by, created_by, deleted_by))

#欄位名稱更改為設定的欄位代號
for (i in 2 : dim(load)[2]) #從2開始是因為第一的欄位是update_at
{
  colnames(load)[i] <- col_names$name[grep(paste(colnames(load)[i], "$", sep = ""), col_names$id)]
}
#格式調整
load$organization_id <- formatC(load$organization_id, dig = 0, wid = 6, format = "f", flag = "0")

#只留下審核通過之名單
load <- merge(x = load, y = list_agree, by = "organization_id", all.x = TRUE) %>%
  subset(agree == 1) %>%
  subset(select = -c(updated_at, agree))

#管區
load <- load %>%
  #subset(organization_id == "421302" | organization_id == "381302" | organization_id == "331404" | organization_id == "421303" | organization_id == "411301" | organization_id == "401302" | organization_id == "351402" | organization_id == "381305" | organization_id == "411401" | organization_id == "421301" | organization_id == "361B09" | organization_id == "331403" | organization_id == "331402" | organization_id == "341302" | organization_id == "411302" | organization_id == "401303" | organization_id == "411303" | organization_id == "401301" | organization_id == "381306" | organization_id == "381304" | organization_id == "381303" | organization_id == "331301" | organization_id == "341402" | organization_id == "361401" | organization_id == "311401" | organization_id == "331304" | organization_id == "351B09" | organization_id == "331302" | organization_id == "361301" | organization_id == "351301" | organization_id == "421404" | organization_id == "321399" | organization_id == "381301" | organization_id == "581301" | organization_id == "121318" | organization_id == "581401" | organization_id == "121417" | organization_id == "521303" | organization_id == "581302" | organization_id == "581402" | organization_id == "521301" | organization_id == "551303" | organization_id == "121307" | organization_id == "121410" | organization_id == "121302" | organization_id == "121415" | organization_id == "551402" | organization_id == "521401" | organization_id == "551301" | organization_id == "121413" | organization_id == "121405" | organization_id == "121320" | organization_id == "121306" | organization_id == "101303" | organization_id == "101304" | organization_id == "101406" | organization_id == "101405" | organization_id == "141406" | organization_id == "141307" | organization_id == "141301" | organization_id == "181305" | organization_id == "181307" | organization_id == "181308" | organization_id == "181306" | organization_id == "201314" | organization_id == "201313" | organization_id == "201304" | organization_id == "201408" | organization_id == "201309" | organization_id == "201312" | organization_id == "201310")
  
#輸出
list_of_datasets <- list("教員資料表" = teacher, "職員(工)資料表" = staff, "教學資料表" = load)
path <- "\\\\192.168.110.244\\share_tmp\\給修恒\\1101\\edhr-110t1\\dta\\edhr_110t1-202201\\110學年度上學期高級中等學校教育人力資源資料庫（20校人事及教務）.xlsx"
openxlsx :: write.xlsx(list_of_datasets, file = path, rowNames = FALSE, overwrite = TRUE)

excel_sheets(path)

data_teacher <- read_excel(path, sheet = "教員資料表")
data_staff   <- read_excel(path, sheet = "職員(工)資料表")
#data_retire   <- read_excel(path, sheet = "離退教職員(工)資料表")
data_load    <- read_excel(path, sheet = "教學資料表")
data_courseA  <- read_excel(path, sheet = "授課資料表A")
data_courseB  <- read_excel(path, sheet = "授課資料表B")

#自動化檢誤用
teacher <- data_teacher
staff <- data_staff
load <- data_load
courseA <- data_courseA
courseB <- data_courseB

#請輸入本次填報設定檔標題(字串需與標題完全相符，否則會找不到)
title <- "110學年度上學期高級中等學校教育人力資源資料庫（20校人事及教務）"

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

# 匯入上一期人事資料檔 -------------------------------------------------------------------

# 1092 20校試辦 教員資料表
#請輸入本次填報設定檔標題(字串需與標題完全相符，否則會找不到)
title_pre <- "109學年度下學期高級中等學校教育人力資源資料庫（20校人事及教務）"

department_pre <- "人事室"

#讀取審核同意之學校名單
list_agree_pre <- dbGetQuery(edhr, 
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
                                                            WHERE title = '", title_pre, "'))", sep = "")
) %>%
  distinct(organization_id, .keep_all = TRUE)

#讀取教員資料表名稱
teacher_tablename_pre <- dbGetQuery(edhr, 
                                paste("
SELECT [name] FROM [plat5_edhr].[dbo].[row_tables] 
	where sheet_id = (SELECT [id] FROM [plat5_edhr].[dbo].[row_sheets] 
						          where file_id = (SELECT field_component_id FROM [plat5_edhr].[dbo].[teacher_datasets] 
											                  WHERE title = '教員資料表' AND department_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_departments] 
																						                                              WHERE title = '", department_pre, "' AND  report_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_reports] 
																												                                                                                      WHERE title = '", title_pre, "'))))", sep = "")
) %>% as.character()

#讀取教員資料表
teacher_pre <- dbGetQuery(edhr, 
                      paste("SELECT * FROM [rows].[dbo].[", teacher_tablename_pre, "] WHERE deleted_at IS NULL", sep = "")
) %>%
  subset(select = -c(id, created_at, deleted_at, updated_by, created_by, deleted_by))

#欄位名稱更改為設定的欄位代號
col_names_pre <- dbGetQuery(edhr, "SELECT id, name, title FROM [plat5_edhr].[dbo].[row_columns]")
col_names_pre$id <- paste("C", col_names_pre$id, sep = "")
for (i in 2 : dim(teacher_pre)[2]) #從2開始是因為第一的欄位是update_at
{
  colnames(teacher_pre)[i] <- col_names_pre$name[grep(paste(colnames(teacher_pre)[i], "$", sep = ""), col_names_pre$id)]
}
#格式調整
teacher_pre$gender <- formatC(teacher_pre$gender, dig = 0, wid = 1, format = "f", flag = "0")
teacher_pre$birthdate <- formatC(teacher_pre$birthdate, dig = 0, wid = 7, format = "f", flag = "0")
teacher_pre$onbodat <- formatC(teacher_pre$onbodat, dig = 0, wid = 7, format = "f", flag = "0")
teacher_pre$desedym <- formatC(teacher_pre$desedym, dig = 0, wid = 4, format = "f", flag = "0")
teacher_pre$beobdym <- formatC(teacher_pre$beobdym, dig = 0, wid = 4, format = "f", flag = "0")
teacher_pre$organization_id <- formatC(teacher_pre$organization_id, dig = 0, wid = 6, format = "f", flag = "0")

#只留下審核通過之名單
teacher_pre <- merge(x = teacher_pre, y = list_agree_pre, by = "organization_id", all.x = TRUE) %>%
  subset(agree == 1) %>%
  subset(select = -c(updated_at, agree))

teacher_pre <- teacher_pre %>%
  mutate(dta_teacher = "教員資料表")

# 1092 20校試辦 職員(工)資料表
#讀取職員(工)資料表名稱
staff_tablename_pre <- dbGetQuery(edhr, 
                              paste("
SELECT [name] FROM [plat5_edhr].[dbo].[row_tables] 
	where sheet_id = (SELECT [id] FROM [plat5_edhr].[dbo].[row_sheets] 
						          where file_id = (SELECT field_component_id FROM [plat5_edhr].[dbo].[teacher_datasets] 
											                   WHERE title = '職員(工)資料表' AND department_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_departments] 
																							                                                 WHERE title = '", department_pre, "' AND  report_id = (SELECT id FROM [plat5_edhr].[dbo].[teacher_reports] 
																												                                                            WHERE title = '", title_pre, "'))))", sep = "")
) %>% as.character()

#讀取職員(工)資料表
staff_pre <- dbGetQuery(edhr, 
                    paste("SELECT * FROM [rows].[dbo].[", staff_tablename_pre, "] WHERE deleted_at IS NULL", sep = "")
) %>%
  subset(select = -c(id, created_at, deleted_at, updated_by, created_by, deleted_by))
#欄位名稱更改為設定的欄位代號
for (i in 2 : dim(staff_pre)[2]) #從2開始是因為第一的欄位是update_at
{
  colnames(staff_pre)[i] <- col_names_pre$name[grep(paste(colnames(staff_pre)[i], "$", sep = ""), col_names_pre$id)]
}

#格式調整
staff_pre$gender <- formatC(staff_pre$gender, dig = 0, wid = 1, format = "f", flag = "0")
staff_pre$birthdate <- formatC(staff_pre$birthdate, dig = 0, wid = 7, format = "f", flag = "0")
staff_pre$onbodat <- formatC(staff_pre$onbodat, dig = 0, wid = 7, format = "f", flag = "0")
staff_pre$desedym <- formatC(staff_pre$desedym, dig = 0, wid = 4, format = "f", flag = "0")
staff_pre$beobdym <- formatC(staff_pre$beobdym, dig = 0, wid = 4, format = "f", flag = "0")
staff_pre$organization_id <- formatC(staff_pre$organization_id, dig = 0, wid = 6, format = "f", flag = "0")

#只留下審核通過之名單
staff_pre <- merge(x = staff_pre, y = list_agree_pre, by = "organization_id", all.x = TRUE) %>%
  subset(agree == 1) %>%
  subset(select = -c(updated_at, agree))

staff_pre <- staff_pre %>%
  mutate(dta_teacher = "職員(工)資料表")

#合併教員資料表與職員(工)資料表
drev_person_pre <- bind_rows(teacher_pre, staff_pre) %>%
  rename(source = dta_teacher)

#記錄時間
time_now <- Sys.time()

# 合併人事資料表 ----------------------------------------------------------------
data_teacher <- data_teacher %>%
  mutate(source = 1)

data_staff <- data_staff %>%
  mutate(source = 2)

drev_person <- bind_rows(data_teacher, data_staff)

drev_person$source  <- factor(drev_person$source, levels = c(1, 2), labels = c("教員資料表", "職員(工)資料表"))
#這行在更改source的1和2為教員資料表及職員工資料表，levels是排序依據.





# 統計處高級中等學校科別資料 -----------------------------------------------------------
# 統計處公布的「高級中等學校科別資料」
# 網址：https://depart.moe.edu.tw/ED4500/News.aspx?n=5A930C32CC6C3818&sms=91B3AAE8C6388B96
edu_path <- "C:\\edhr-111t1\\110_base2_revise.xls"
data_schtype <- read_excel(edu_path)
data_schtype <- c("學校代碼", "學校名稱", "學程(等級)別", "學程(等級)名稱", "日夜別", "日夜別名稱", "群別代碼", "群別名稱", "科系代碼", "科系名稱", "班級數", "學生數") %>%
  data_schtype[, .]
# 改變變項名稱與形態
data_schtype <-plyr:: rename(data_schtype, c(   "學校代碼"       = "organization_id"
                                                , "學校名稱"       = "edu_name"
                                                , "學程(等級)別"   = "type_code"
                                                , "學程(等級)名稱" = "type_name"
                                                , "日夜別"         = "dn_code"
                                                , "日夜別名稱"     = "dn_name"
                                                , "群別代碼"       = "dep1_code"
                                                , "群別名稱"       = "dep1_name"
                                                , "科系代碼"       = "depcode"
                                                , "科系名稱"       = "dep2_name"
                                                , "班級數"         = "nclass"
                                                , "學生數"         = "nstudent"))
data_schtype$nclass[data_schtype$nclass == "-"]     <- NA
data_schtype$nstudent[data_schtype$nstudent == "-"] <- NA
data_schtype$nclass   <- as.numeric(data_schtype$nclass)
data_schtype$nstudent <- as.numeric(data_schtype$nstudent)

data_schtype$organization_id <- recode_factor(data_schtype$organization_id,
                                       "140222" = "140401"
                                       , "400144" = "400419")
# 參與學校名單
try <- c("010301", "013303", "013304", "013335", "013336", "013337", "013338", "013339", "013402", "013430", "013433", "013434", "014302", "014311", "014315", "014322", "014326", "014332", "014338", "014343", "014347", "014348", "014353", "014356", "014357", "014362", "014363", "014364", "014381", "014439", "014468", "020301", "020302", "020308", "020403", "020404", "020405", "020407", "020409", "024322", "024325", "030305", "030403", "033302", "033304", "033306", "033316", "033325", "033327", "033407", "033408", "034306", "034312", "034314", "034319", "034332", "034335", "034347", "034348", "034399", "040302", "040304", "040308", "044311", "044320", "050303", "050310", "050314", "050315", "050401", "050404", "050407", "054308", "054309", "054317", "054333", "060322", "060323", "063303", "063305", "063312", "063401", "063402", "063404", "063407", "063408", "064308", "064324", "064328", "064336", "064342", "064350", "064406", "070301", "070304", "070307", "070316", "070319", "070401", "070402", "070403", "070405", "070406", "070408", "070409", "070410", "070415", "074308", "074313", "074323", "074328", "074339", "080302", "080305", "080307", "080308", "080401", "080403", "080404", "080406", "080410", "084309", "090305", "090306", "090315", "090401", "090402", "090403", "090404", "090413", "094301", "094307", "094308", "094326", "100301", "100302", "100402", "104319", "104326", "110302", "110308", "110311", "110312", "110314", "110315", "110317", "110328", "110401", "110403", "110404", "110405", "110406", "110407", "110409", "110410", "114306", "114307", "120303", "120304", "120311", "120319", "120401", "120402", "120409", "124302", "124311", "124313", "124322", "124333", "124340", "130302", "130305", "130306", "130322", "130401", "130403", "130404", "130410", "130417", "134304", "134321", "134324", "134334", "140301", "140302", "140303", "140404", "140405", "140408", "144322", "150302", "150303", "150309", "150401", "150404", "150405", "150411", "154301", "160302", "160401", "170301", "170302", "170403", "170404", "173304", "173306", "173307", "173314", "180301", "180302", "180309", "180403", "180404", "183306", "183307", "183313", "190406", "193301", "193302", "193303", "193313", "193315", "193316", "193404", "193407", "194303", "194315", "200302", "200303", "200401", "200405", "200406", "200407", "210303", "210305", "210306", "210309", "210408", "210416", "213303", "213316", "313301", "313302", "323301", "323302", "323401", "323402", "330301", "333301", "333304", "333401", "343301", "343302", "343303", "353301", "353302", "353303", "363301", "363302", "373301", "373302", "380301", "383301", "383302", "383401", "393301", "393302", "393401", "403301", "403302", "403303", "403401", "413301", "413302", "413401", "423301", "423302", "383303", "523301", "533301", "533302", "533401", "533402", "540301", "543301", "543302", "553301", "553302", "553401", "563301", "563401", "573301", "580301", "583301", "593301", "593302", "593401", "610405", "613301", "710301", "710401", "720301")

data_schtype$try1 <- NA
for (x in try){
  data_schtype$try1[data_schtype$organization_id == x] <- 1
}
# 主管機關
data_schtype$authority <- NA
data_schtype$authority[substr(data_schtype$organization_id, 3, 3) == "0"] <- "國立"
data_schtype$authority[substr(data_schtype$organization_id, 3, 3) == "1"] <- "私立"
data_schtype$authority[substr(data_schtype$organization_id, 3, 3) == "3" | substr(data_schtype$organization_id, 3, 3) == "4"] <- "縣市立"

data_schtype$authority[       data_schtype$authority == "國立" & (data_schtype$organization_id == "140401" | data_schtype$organization_id == "400419")]                                     <-  "技職司管轄國立"    
data_schtype$authority[       data_schtype$authority == "國立" & (data_schtype$organization_id == "110328" | data_schtype$organization_id == "180301" | data_schtype$organization_id == "060323")] <-  "國教署與科技部共管"                         


data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 1]  <- "新北市市立"
data_schtype$authority[       data_schtype$authority ==   "私立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 1]  <- "新北市私立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 2]  <- "宜蘭縣縣立"
data_schtype$authority[       data_schtype$authority ==   "私立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 3]  <- "桃園市私立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 3]  <- "桃園市市立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 4]  <- "新竹縣縣立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 5]  <- "苗栗縣縣立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 6]  <- "臺中市市立"
data_schtype$authority[       data_schtype$authority ==   "私立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 6]  <- "臺中市私立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 7]  <- "彰化縣縣立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 8]  <- "南投縣縣立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 9]  <- "雲林縣縣立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 10] <- "嘉義縣縣立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 11] <- "臺南市市立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 12] <- "高雄市市立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 13] <- "屏東縣縣立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 14] <- "臺東縣縣立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 15] <- "花蓮縣縣立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 16] <- "澎湖縣縣立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 17] <- "基隆市市立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 18] <- "新竹市市立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 19] <- "臺中市市立"
data_schtype$authority[       data_schtype$authority ==   "私立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 19] <- "臺中市私立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 20] <- "嘉義市市立"
data_schtype$authority[       data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 21] <- "臺南市市立"

for (x in 30:42){
  data_schtype$authority[data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == x]  <- "臺北市市立"
  data_schtype$authority[data_schtype$authority ==   "私立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == x]  <- "臺北市私立"
}

for (x in 50:61){
  data_schtype$authority[data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == x]  <- "高雄市市立"
  data_schtype$authority[data_schtype$authority ==   "私立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == x]  <- "高雄市私立"
}
data_schtype$authority[data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 64] <- "高雄市市立"
data_schtype$authority[data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 66] <- "臺中市市立"
data_schtype$authority[data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 67] <- "臺南市市立"                        
data_schtype$authority[data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 71] <- "金門縣縣立"
data_schtype$authority[data_schtype$authority == "縣市立" & as.numeric(substr(data_schtype$organization_id, 1, 2)) == 72] <- "連江縣縣立"          
data_schtype$authority[data_schtype$authority == "國立"]                                                           <- "國教署管轄國立" 
data_schtype$authority[data_schtype$authority == "私立"]                                                           <- "國教署管轄私立"


# type_code
data_schtype$type_code <- recode_factor(data_schtype$type_code,
                                        "A" = "H"
                                        ,"B" = "V"
                                        ,"O" = "M")
data_schtype$type_code <- factor(data_schtype$type_code, levels = c("C", "E", "F", "H", "V", "M", "J", "U"))
data_schtype$type_code[data_schtype$dn_code == "D" & data_schtype$type_code == "U"] <- "E"
data_schtype$type_code[data_schtype$dn_code == "N" & data_schtype$type_code == "U"] <- "F"

#計算班級數
  #縣市立別
data_schtype_v2 <- data_schtype %>%
  subset(type_code != "J")
data_schtype_v2 <- data_schtype_v2 %>%
  left_join(aggregate(nclass ~ authority, data_schtype_v2, sum), by = c("authority")) %>% #縣市立別
  rename(nclass = nclass.x, nauthority = nclass.y) %>%
  left_join(aggregate(nclass ~ organization_id, data_schtype_v2, sum), by = c("organization_id")) %>% #學校別
  rename(nclass = nclass.x, nsch = nclass.y) %>%
  left_join(aggregate(nclass ~ organization_id + type_code, data_schtype_v2, sum), by = c("organization_id", "type_code")) %>% #學制別
  rename(nclass = nclass.x, ntype = nclass.y) %>%
  left_join(aggregate(nclass ~ organization_id + type_code + dep1_code, data_schtype_v2, sum), by = c("organization_id", "type_code", "dep1_code")) %>% #群別
  rename(nclass = nclass.x, ndep = nclass.y) %>%
  subset(select = c(organization_id, type_code, dep1_code, dep1_name, depcode, dep2_name, nclass, nauthority, ntype, nsch)) %>%
  rename(cinsticode = type_code, studyareacode = dep1_code)


# 各校班級的開課單位
data_schtype_wide <- data_schtype %>%
  mutate(schtype = paste("type", type_code, sep = "")) %>%
  spread(key = schtype, value = nstudent)

data_schtype_wide <- data_schtype_wide %>%
  group_by(organization_id) %>%
  mutate(typeC = sum(typeC, na.rm = TRUE)) %>%
  group_by(organization_id) %>%
  mutate(typeH = sum(typeH, na.rm = TRUE)) %>%              
  group_by(organization_id) %>%
  mutate(typeJ = sum(typeJ, na.rm = TRUE)) %>%
  group_by(organization_id) %>%
  mutate(typeV = sum(typeV, na.rm = TRUE)) %>%
  group_by(organization_id) %>%
  mutate(typeM = sum(typeM, na.rm = TRUE)) %>%
  group_by(organization_id) %>%
  mutate(typeF = sum(typeF, na.rm = TRUE)) %>%
  group_by(organization_id) %>%
  mutate(typeE = sum(typeE, na.rm = TRUE)) 

# 國際部與雙語部的名單會逐年變動。


temp <- data_schtype %>%  
  filter(dep1_code == "11", dep2_name == "雙語部") %>%
  distinct(organization_id) %>%
  `[[`(1) %>%
  as.character()

data_schtype_wide$typeD <- 0
for (x in temp){
  data_schtype_wide <- data_schtype_wide %>% 
    mutate(typeD = if_else((organization_id == x | typeD == 1), 1 ,0)) 
}

data_schtype_wide <- data_schtype_wide %>%
  mutate(typeH = if_else(as.numeric(typeH) > 1, 1, 0, missing = NULL),
         typeJ = if_else(as.numeric(typeJ) > 1, 1, 0, missing = NULL),
         typeV = if_else(as.numeric(typeV) > 1, 1, 0, missing = NULL),
         typeC = if_else(as.numeric(typeC) > 1, 1, 0, missing = NULL),
         typeM = if_else(as.numeric(typeM) > 1, 1, 0, missing = NULL),
         typeF = if_else(as.numeric(typeF) > 1, 1, 0, missing = NULL),
         typeE = if_else(as.numeric(typeE) > 1, 1, 0, missing = NULL),
         typeD = if_else(as.numeric(typeD) > 1, 1, 0, missing = 0))

# 更改人事資料表的學校名稱（若學校名單異動，哪幾間學校名稱需簡寫亦要檢視） -------------------------------------------------------------------
# edu_name*
drev_person$edu_name <- recode_factor(drev_person$organization_id,
                                      "010301" = "國立華僑高級中等學校", 
                                      "014347" = "新北市立雙溪高級中學", 
                                      "014362" = "新北市立錦和高級中學", 
                                      "063305" = "臺中市立清水高級中等學校", 
                                      "063312" = "臺中市立豐原高級中等學校", 
                                      "064308" = "臺中市立后綜高級中學", 
                                      "070408" = "國立員林高級農工職業學校", 
                                      "080307" = "國立竹山高級中學", 
                                      "100302" = "國立新港藝術高級中學", 
                                      "100402" = "國立民雄高級農工職業學校", 
                                      "104319" = "嘉義縣立竹崎高級中學", 
                                      "110308" = "國立臺南大學附屬高級中學", 
                                      "110312" = "國立新營高級中學", 
                                      "110314" = "國立後壁高級中學", 
                                      "120401" = "國立旗山高級農工職業學校", 
                                      "130322" = "國立屏北高級中學", 
                                      "130410" = "國立東港高級海事水產職業學校", 
                                      "131418" = "屏東縣私立日新高級工商職業學校", 
                                      "140301" = "國立臺東大學附屬體育高級中學", 
                                      "150405" = "國立花蓮高級商業職業學校", 
                                      "160401" = "國立澎湖高級海事水產職業學校", 
                                      "173306" = "基隆市立安樂高級中學", 
                                      "173307" = "基隆市立暖暖高級中學", 
                                      "201312" = "嘉義市私立輔仁高級中學", 
                                      "341301" = "臺北市私立中興高級中學", 
                                      "353302" = "臺北市立成功高級中學", 
                                      "380301" = "國立政治大學附屬高級中學", 
                                      "423301" = "臺北市立復興高級中學", 
                                      "423302" = "臺北市立中正高級中學", 
                                      "540301" = "國立中山大學附屬國光高級中學", 
                                      "551301" = "立志學校財團法人高雄市立志高級中學", 
                                      "710401" = "國立金門高級農工職業學校", 
                                      "321399" = "韻鏗學校財團法人臺北市私立協和祐德高級中學", 
                                      "011329" = "辭修學校財團法人新北市私立辭修高級中學", 
                                      "111321" = "興國學校財團法人臺南市興國高級中學", 
                                      "011413" = "穀保學校財團法人新北市穀保高級家事商業職業學校", 
                                      "171308" = "輔仁大學學校財團法人基隆市輔大聖心高級中學", 
                                      "101406" = "嘉義縣私立萬能高級工商職業學校", 
                                      "201304" = "嘉義市私立興華高級中學", 
                                      "201310" = "嘉義市私立嘉華高級中學", 
                                      "201408" = "嘉義市私立東吳高級工業家事職業學校", 
                                      "201309" = "嘉義市私立仁義高級中學", 
                                      "091312" = "雲林縣私立巨人高級中學", 
                                      "091410" = "雲林縣私立大成高級商工職業學校", 
                                      "171407" = "基隆市私立培德高級工業家事職業學校", 
                                      "031414" = "桃園市私立成功高級工商職業學校", 
                                      "211310" = "臺南光華學校財團法人臺南市光華高級中學", 
                                      "211318" = "臺南市德光高級中學", 
                                      "111326" = "臺南市私立新榮高級中學", 
                                      "111320" = "臺南市私立港明高級中學", 
                                      "111313" = "臺南市私立南光高級中學", 
                                      "211301" = "臺南市私立長榮高級中學", 
                                      "211302" = "臺南市私立長榮女子高級中學", 
                                      "111322" = "臺南市私立明達高級中學", 
                                      "211412" = "臺南市私立亞洲高級餐旅職業學校", 
                                      "111427" = "臺南市私立育德工業家事職業學校", 
                                      "213303" = "臺南市立南寧高級中學", 
                                      "114307" = "臺南市立永仁高級中學", 
                                      "114306" = "臺南市立大灣高級中學", 
                                      "213316" = "臺南市立土城高級中學", 
                                      "211314" = "臺南市六信高級中學", 
                                      "144322" = "臺東縣立蘭嶼高級中學", 
                                      "421301" = "臺北市私立薇閣高級中學", 
                                      "361301" = "臺北市私立靜修高級中學", 
                                      "411302" = "臺北市私立衛理女子高級中學", 
                                      "341402" = "臺北市私立稻江高級護理家事職業學校", 
                                      "361401" = "臺北市私立稻江高級商業職業學校", 
                                      "381302" = "臺北市私立滬江高級中學", 
                                      "401303" = "臺北市私立達人女子高級中學", 
                                      "351402" = "開南學校財團法人臺北市開南高級中等學校", 
                                      "331404" = "臺北市私立開平餐飲職業學校", 
                                      "411401" = "臺北市私立華岡藝術學校", 
                                      "381305" = "臺北市私立景文高級中學", 
                                      "331304" = "臺北市私立復興實驗高級中學", 
                                      "331403" = "臺北市私立喬治高級工商職業學校", 
                                      "421404" = "臺北市私立惇敘高級工商職業學校", 
                                      "351301" = "臺北市私立強恕高級中學", 
                                      "411301" = "臺北市私立泰北高級中學", 
                                      "351B09" = "臺北市私立南華高級中學職業進修學校", 
                                      "331302" = "臺北市私立金甌女子高級中學", 
                                      "331402" = "臺北市私立東方高級工商職業學校", 
                                      "193315" = "臺北市私立東山高級中學", 
                                      "381301" = "臺北市私立東山高級中學", 
                                      "331301" = "臺北市私立延平高級中學", 
                                      "311401" = "育達學校財團法人臺北市私立普林思頓高級中等學校", 
                                      "361B09" = "臺北市私立志仁高級中學職業進修學校", 
                                      "381304" = "臺北市私立再興高級中學", 
                                      "371301" = "臺北市私立立人高級中學", 
                                      "401302" = "臺北市私立方濟高級中學", 
                                      "401301" = "臺北市私立文德女子高級中學", 
                                      "421302" = "幼華學校財團法人臺北市幼華高級中等學校", 
                                      "403302" = "臺北市立麗山高級中學", 
                                      "383302" = "臺北市立萬芳高級中學", 
                                      "413301" = "臺北市立陽明高級中學", 
                                      "373301" = "臺北市立華江高級中學", 
                                      "383301" = "臺北市立景美女子高級中學", 
                                      "353303" = "臺北市立第一女子高級中學", 
                                      "363F01" = "臺北市立啟聰學校", 
                                      "413F01" = "臺北市立臺北特殊教育學校", 
                                      "413F02" = "臺北市立啟明學校", 
                                      "353301" = "臺北市立建國高級中學", 
                                      "403303" = "臺北市立南湖高級中學", 
                                      "393301" = "臺北市立南港高級中學", 
                                      "393401" = "臺北市立南港高級工業職業學校", 
                                      "323401" = "臺北市立松山高級商業家事職業學校", 
                                      "323301" = "臺北市立松山高級中學", 
                                      "323402" = "臺北市立松山高級工農職業學校", 
                                      "363301" = "臺北市立明倫高級中學", 
                                      "333301" = "臺北市立和平高級中學", 
                                      "393302" = "臺北市立育成高級中學", 
                                      "313301" = "臺北市立西松高級中學", 
                                      "413302" = "臺北市立百齡高級中學", 
                                      "363302" = "臺北市立成淵高級中學", 
                                      "323302" = "臺北市立永春高級中學", 
                                      "383401" = "臺北市立木柵高級工業職業學校", 
                                      "383F01" = "臺北市立文山特殊教育學校", 
                                      "403301" = "臺北市立內湖高級中學", 
                                      "403401" = "臺北市立內湖高級工業職業學校", 
                                      "313302" = "臺北市立中崙高級中學", 
                                      "343301" = "臺北市立中山女子高級中學", 
                                      "373302" = "臺北市立大理高級中學", 
                                      "343303" = "臺北市立大直高級中學", 
                                      "333401" = "臺北市立大安高級工業職業學校", 
                                      "343302" = "臺北市立大同高級中學", 
                                      "413401" = "臺北市立士林高級商業職業學校", 
                                      "061315" = "臺中市華盛頓高級中學", 
                                      "061316" = "臺中市青年高級中學", 
                                      "061306" = "臺中市明台高級中學", 
                                      "191314" = "臺中市私立嶺東高級中學", 
                                      "061314" = "臺中市私立僑泰高級中學", 
                                      "191305" = "臺中市私立新民高級中學", 
                                      "061321" = "臺中市私立慈明高級中學", 
                                      "061309" = "臺中市私立致用高級中學", 
                                      "061313" = "臺中市私立明道高級中學", 
                                      "061318" = "臺中市私立立人高級中學", 
                                      "061319" = "臺中市私立玉山高級中學", 
                                      "061317" = "臺中市私立弘文高級中學", 
                                      "061310" = "臺中市私立大明高級中學", 
                                      "063408" = "臺中市立霧峰農業工業高級中等學校", 
                                      "063401" = "臺中市立豐原商業高級中等學校", 
                                      "064350" = "臺中市立龍津高級中等學校", 
                                      "194303" = "臺中市立臺中第二高級中等學校", 
                                      "193302" = "臺中市立臺中第一高級中等學校", 
                                      "193F02" = "臺中市立臺中特殊教育學校", 
                                      "193404" = "臺中市立臺中家事商業高級中等學校", 
                                      "193407" = "臺中市立臺中工業高級中等學校", 
                                      "193301" = "臺中市立臺中女子高級中等學校", 
                                      "064328" = "臺中市立新社高級中學", 
                                      "193316" = "臺中市立惠文高級中學", 
                                      "193F01" = "臺中市立啟聰學校", 
                                      "063F01" = "臺中市立啟明學校", 
                                      "064336" = "臺中市立長億高級中學", 
                                      "063404" = "臺中市立東勢工業高級中等學校", 
                                      "193303" = "臺中市立忠明高級中學", 
                                      "063407" = "臺中市立沙鹿工業高級中等學校", 
                                      "193313" = "臺中市立西苑高級中學", 
                                      "194315" = "臺中市立文華高級中等學校", 
                                      "064342" = "臺中市立中港高級中學", 
                                      "064324" = "臺中市立大里高級中學", 
                                      "063303" = "臺中市立大甲高級中等學校", 
                                      "063402" = "臺中市立大甲工業高級中等學校", 
                                      "091320" = "維多利亞學校財團法人雲林縣維多利亞實驗高級中學", 
                                      "071413" = "彰化縣私立大慶高級商工職業學校", 
                                      "074308" = "彰化縣立彰化藝術高級中學", 
                                      "074323" = "彰化縣立和美高級中學", 
                                      "074339" = "彰化縣立成功高級中學", 
                                      "074328" = "彰化縣立田中高級中學", 
                                      "074313" = "彰化縣立二林高級中學", 
                                      "104326" = "嘉義縣立永慶高級中學", 
                                      "201314" = "嘉義市立仁高級中學", 
                                      "061311" = "嘉陽學校財團法人臺中市嘉陽高級中學", 
                                      "191302" = "葳格學校財團法人臺中市葳格高級中學", 
                                      "031320" = "新興學校財團法人桃園市新興高級中等學校", 
                                      "041307" = "新竹縣仰德高級中學", 
                                      "044320" = "新竹縣立湖口中學", 
                                      "044311" = "新竹縣立六家高級中學", 
                                      "183307" = "新竹市立香山高級中學", 
                                      "183313" = "新竹市立建功高級中學", 
                                      "183306" = "新竹市立成德高級中學", 
                                      "181308" = "新竹市世界高級中學", 
                                      "121306" = "新光學校財團法人高雄市新光高級中學", 
                                      "011317" = "新北市私立醒吾高級中學", 
                                      "011427" = "新北市私立豫章高級工商職業學校", 
                                      "011405" = "新北市私立樹人高級家事商業職業學校", 
                                      "011307" = "新北市私立裕德高級中等學校", 
                                      "011311" = "新北市私立聖心女子高級中學", 
                                      "011419" = "新北市私立開明高級工業商業職業學校", 
                                      "011407" = "新北市私立復興高級商工職業學校", 
                                      "011431" = "新北市私立莊敬高級工業家事職業學校", 
                                      "011318" = "新北市私立徐匯高級中學", 
                                      "011408" = "新北市私立南強高級工商職業學校", 
                                      "011324" = "新北市私立竹林高級中學", 
                                      "011C71" = "新北市私立光華高級商業職業進修學校", 
                                      "011432" = "新北市私立中華商業海事職業學校", 
                                      "014439" = "新北市立鶯歌高級工商職業學校", 
                                      "014322" = "新北市立樹林高級中學", 
                                      "014468" = "新北市立樟樹國際實創高級中等學校", 
                                      "013402" = "新北市立瑞芳高級工業職業學校", 
                                      "013337" = "新北市立新莊高級中學", 
                                      "013335" = "新北市立新店高級中學", 
                                      "013338" = "新北市立新北高級中學", 
                                      "013433" = "新北市立新北高級工業職業學校", 
                                      "013F01" = "新北市立新北特殊教育學校", 
                                      "014356" = "新北市立清水高級中學", 
                                      "013434" = "新北市立淡水高級商工職業學校", 
                                      "014302" = "新北市立海山高級中學", 
                                      "013303" = "新北市立泰山高級中學", 
                                      "014338" = "新北市立金山高級中學", 
                                      "013339" = "新北市立林口高級中學", 
                                      "013304" = "新北市立板橋高級中學", 
                                      "014326" = "新北市立明德高級中學", 
                                      "014332" = "新北市立秀峰高級中學", 
                                      "014364" = "新北市立竹圍高級中學", 
                                      "014343" = "新北市立安康高級中學", 
                                      "014363" = "新北市立光復高級中學", 
                                      "014348" = "新北市立石碇高級中學", 
                                      "014315" = "新北市立永平高級中學", 
                                      "014381" = "新北市立北大高級中學", 
                                      "014353" = "新北市立丹鳳高級中學", 
                                      "013336" = "新北市立中和高級中學", 
                                      "013430" = "新北市立三重高級商工職業學校", 
                                      "014311" = "新北市立三重高級中學", 
                                      "014357" = "新北市立三民高級中學", 
                                      "211320" = "慈濟學校財團法人臺南市私立慈濟高級中學", 
                                      "151312" = "慈濟學校財團法人慈濟大學附屬高級中學", 
                                      "091318" = "雲林縣私立義峰高級中學", 
                                      "091316" = "雲林縣私立揚子高級中學", 
                                      "094307" = "雲林縣立麥寮高級中學", 
                                      "094308" = "雲林縣立古坑華德福實驗高級中學", 
                                      "094301" = "雲林縣立斗南高級中學", 
                                      "111419" = "陽明學校財團法人臺南市陽明高級工商職業學校", 
                                      "411303" = "華興學校財團法人臺北市私立華興高級中等學校", 
                                      "121415" = "華德學校財團法人高雄市華德高級工業家事職業學校", 
                                      "011420" = "智光學校財團法人新北市智光高級商工職業學校", 
                                      "031311" = "復旦學校財團法人桃園市復旦高級中等學校", 
                                      "131308" = "陸興學校財團法人屏東縣陸興高級中學", 
                                      "011421" = "清傳學校財團法人新北市清傳高級商業職業學校", 
                                      "031319" = "清華學校財團法人桃園市清華高級中等學校", 
                                      "011302" = "康橋學校財團法人新北市康橋高級中學", 
                                      "173304" = "基隆市立中山高級中學", 
                                      "173314" = "基隆市立八斗高級中學", 
                                      "020301" = "國立蘭陽女子高級中學", 
                                      "020405" = "國立蘇澳高級海事水產職業學校", 
                                      "040304" = "國立關西高級中學", 
                                      "140404" = "國立關山高級工商職業學校", 
                                      "020404" = "國立羅東高級商業職業學校", 
                                      "020308" = "國立羅東高級中學", 
                                      "020407" = "國立羅東高級工業職業學校", 
                                      "030302" = "桃園市立龍潭高級中等學校", 
                                      "020409" = "國立頭城高級家事商業職業學校", 
                                      "130306" = "國立潮州高級中學", 
                                      "120319" = "國立鳳新高級中學", 
                                      "120409" = "國立鳳山高級商工職業學校", 
                                      "120303" = "國立鳳山高級中學", 
                                      "400419" = "國立臺灣戲曲學院", 
                                      "330301" = "國立臺灣師範大學附屬高級中學", 
                                      "210303" = "國立臺南第二高級中學", 
                                      "210305" = "國立臺南第一高級中學", 
                                      "210F01" = "國立臺南特殊教育學校", 
                                      "210408" = "國立臺南高級商業職業學校", 
                                      "210416" = "國立臺南高級海事水產職業學校", 
                                      "110409" = "國立臺南高級工業職業學校", 
                                      "210309" = "國立臺南家齊高級中等學校", 
                                      "210306" = "國立臺南女子高級中學", 
                                      "110F01" = "國立臺南大學附屬啟聰學校", 
                                      "140401" = "國立臺東專科學校", 
                                      "140C01" = "國立臺東專科學校", 
                                      "140405" = "國立臺東高級商業職業學校", 
                                      "140303" = "國立臺東高級中學", 
                                      "140302" = "國立臺東女子高級中學", 
                                      "140F01" = "國立臺東大學附屬特殊教育學校", 
                                      "030403" = "國立臺北科技大學附屬桃園農工高級中等學校", 
                                      "080308" = "國立暨南國際大學附屬高級中學", 
                                      "120311" = "國立旗美高級中學", 
                                      "070406" = "國立彰化高級商業職業學校", 
                                      "070307" = "國立彰化高級中學", 
                                      "070F02" = "國立彰化特殊教育學校", 
                                      "070401" = "國立彰化師範大學附屬高級工業職業學校", 
                                      "070301" = "國立彰化女子高級中學", 
                                      "200406" = "國立嘉義高級商業職業學校", 
                                      "200407" = "國立嘉義高級家事職業學校", 
                                      "200303" = "國立嘉義高級中學", 
                                      "200405" = "國立嘉義高級工業職業學校", 
                                      "200F01" = "國立嘉義特殊教育學校", 
                                      "200302" = "國立嘉義女子高級中學", 
                                      "070319" = "國立溪湖高級中學", 
                                      "030316" = "桃園市立楊梅高級中等學校", 
                                      "110302" = "國立新豐高級中學", 
                                      "110406" = "國立新營高級工業職業學校", 
                                      "180403" = "國立新竹高級商業職業學校", 
                                      "180309" = "國立新竹高級中學", 
                                      "180404" = "國立新竹高級工業職業學校", 
                                      "040F01" = "國立新竹特殊教育學校", 
                                      "180302" = "國立新竹女子高級中學", 
                                      "110317" = "國立新化高級中學", 
                                      "110401" = "國立新化高級工業職業學校", 
                                      "090F01" = "國立雲林特殊教育學校", 
                                      "030325" = "桃園市立陽明高級中等學校", 
                                      "200401" = "國立華南高級商業職業學校", 
                                      "110410" = "國立曾文高級農工職業學校", 
                                      "110405" = "國立曾文高級家事商業職業學校", 
                                      "110315" = "國立善化高級中學", 
                                      "070316" = "國立鹿港高級中學", 
                                      "170404" = "國立基隆高級商工職業學校", 
                                      "170403" = "國立基隆高級海事職業學校", 
                                      "170302" = "國立基隆高級中學", 
                                      "170F01" = "國立基隆特殊教育學校", 
                                      "170301" = "國立基隆女子高級中學", 
                                      "610405" = "國立高雄餐旅大學附屬餐旅高級中等學校", 
                                      "580301" = "國立高雄師範大學附屬高級中學", 
                                      "720301" = "國立馬祖高級中學", 
                                      "160302" = "國立馬公高級中學", 
                                      "080406" = "國立草屯高級商工職業學校", 
                                      "030F01" = "桃園市立桃園特殊教育學校", 
                                      "030304" = "桃園市立桃園高級中等學校", 
                                      "080403" = "國立埔里高級工業職業學校", 
                                      "070409" = "國立員林崇實高級工業職業學校", 
                                      "070410" = "國立員林高級家事商業職業學校", 
                                      "070304" = "國立員林高級中學", 
                                      "050404" = "國立苗栗高級農工職業學校", 
                                      "050407" = "國立苗栗高級商業職業學校", 
                                      "050303" = "國立苗栗高級中學", 
                                      "050F01" = "國立苗栗特殊教育學校", 
                                      "050315" = "國立苑裡高級中學", 
                                      "180301" = "國立新竹科學園區實驗高級中等學校", 
                                      "130417" = "國立恆春高級工商職業學校", 
                                      "130305" = "國立屏東高級中學", 
                                      "130403" = "國立屏東高級工業職業學校", 
                                      "130F01" = "國立屏東特殊教育學校", 
                                      "130302" = "國立屏東女子高級中學", 
                                      "110328" = "國立南科國際實驗高級中學", 
                                      "080404" = "國立南投高級商業職業學校", 
                                      "080302" = "國立南投高級中學", 
                                      "080F01" = "國立南投特殊教育學校", 
                                      "710301" = "國立金門高級中學", 
                                      "090401" = "國立虎尾高級農工職業學校", 
                                      "090315" = "國立虎尾高級中學", 
                                      "150401" = "國立花蓮高級農業職業學校", 
                                      "150303" = "國立花蓮高級中學", 
                                      "150404" = "國立花蓮高級工業職業學校", 
                                      "150F01" = "國立花蓮特殊教育學校", 
                                      "150302" = "國立花蓮女子高級中學", 
                                      "030306" = "桃園市立武陵高級中等學校", 
                                      "100301" = "國立東石高級中學", 
                                      "120402" = "國立岡山高級農工職業學校", 
                                      "120304" = "國立岡山高級中學", 
                                      "020403" = "國立宜蘭高級商業職業學校", 
                                      "020302" = "國立宜蘭高級中學", 
                                      "020F01" = "國立宜蘭特殊教育學校", 
                                      "070F01" = "國立和美實驗學校", 
                                      "050314" = "國立卓蘭高級中等學校", 
                                      "130404" = "國立佳冬高級農業職業學校", 
                                      "070405" = "國立秀水高級工業職業學校", 
                                      "090402" = "國立西螺高級農工職業學校", 
                                      "050310" = "國立竹南高級中學", 
                                      "040302" = "國立竹東高級中學", 
                                      "040308" = "國立竹北高級中學", 
                                      "140408" = "國立成功商業水產職業學校", 
                                      "210C01" = "國立成功大學附設高級工業職業進修學校", 
                                      "150411" = "國立光復高級商工職業學校", 
                                      "110403" = "國立白河高級商工職業學校", 
                                      "150309" = "國立玉里高級中學", 
                                      "110407" = "國立玉井高級工商職業學校", 
                                      "070402" = "國立永靖高級工業職業學校", 
                                      "090404" = "國立北港高級農工職業學校", 
                                      "090306" = "國立北港高級中學", 
                                      "110404" = "國立北門高級農工職業學校", 
                                      "110311" = "國立北門高級中學", 
                                      "070415" = "國立北斗高級家事商業職業學校", 
                                      "080410" = "國立水里高級商工職業學校", 
                                      "090403" = "國立斗六高級家事商業職業學校", 
                                      "090305" = "國立斗六高級中學", 
                                      "030327" = "桃園市立內壢高級中等學校", 
                                      "130401" = "國立內埔高級農工職業學校", 
                                      "080401" = "國立仁愛高級農業職業學校", 
                                      "030407" = "桃園市立中壢商業高級中等學校", 
                                      "030408" = "桃園市立中壢家事商業高級中等學校", 
                                      "080305" = "國立中興高級中學", 
                                      "190406" = "國立中興大學附屬臺中高級農業職業學校", 
                                      "060322" = "國立中興大學附屬高級中學", 
                                      "060323" = "國立中科實驗高級中學", 
                                      "030305" = "國立中央大學附屬中壢高級中學", 
                                      "050401" = "國立大湖高級農工職業學校", 
                                      "090413" = "國立土庫高級商工職業學校", 
                                      "070403" = "國立二林高級工商職業學校", 
                                      "031318" = "啟英學校財團法人桃園市啟英高級中等學校", 
                                      "551402" = "高雄市私立樹德高級家事商業職校", 
                                      "121320" = "高雄市私立義大國際高級中學", 
                                      "581301" = "高雄市私立復華高級中學", 
                                      "581401" = "高雄市私立國際高級商工職業學校", 
                                      "121318" = "高雄市私立正義高級中學", 
                                      "521401" = "高雄市私立中華高級藝術職業學校", 
                                      "521303" = "高雄市私立大榮高級中學", 
                                      "581402" = "高雄市私立三信高級家事商業職業學校", 
                                      "124340" = "高雄市立福誠高級中學", 
                                      "523301" = "高雄市立鼓山高級中學", 
                                      "124322" = "高雄市立路竹高級中學", 
                                      "593302" = "高雄市立瑞祥高級中學", 
                                      "543302" = "高雄市立楠梓高級中學", 
                                      "543F01" = "高雄市立楠梓特殊學校", 
                                      "563301" = "高雄市立新興高級中學", 
                                      "533302" = "高雄市立新莊高級中學", 
                                      "583F01" = "高雄市立高雄特殊教育學校", 
                                      "563401" = "高雄市立高雄高級商業職業學校", 
                                      "553301" = "高雄市立高雄高級中學", 
                                      "553401" = "高雄市立高雄高級工業職業學校", 
                                      "573301" = "高雄市立高雄女子高級中學", 
                                      "533401" = "高雄市立海青高級工商職業學校", 
                                      "593301" = "高雄市立前鎮高級中學", 
                                      "124311" = "高雄市立林園高級中學", 
                                      "593F01" = "高雄市立成功特殊教育學校", 
                                      "533301" = "高雄市立左營高級中學", 
                                      "124302" = "高雄市立文山高級中學", 
                                      "124333" = "高雄市立六龜高級中學", 
                                      "124313" = "高雄市立仁武高級中學", 
                                      "124F01" = "高雄市立仁武特殊教育學校", 
                                      "583301" = "高雄市立中正高級中學", 
                                      "593401" = "高雄市立中正高級工業職業學校", 
                                      "543301" = "高雄市立中山高級中學", 
                                      "613301" = "高雄市立小港高級中學", 
                                      "533402" = "高雄市立三民高級家事商業職業學校", 
                                      "553302" = "高雄市立三民高級中學", 
                                      "011325" = "財團法人新北市私立及人高級中學", 
                                      "191301" = "財團法人東海大學附屬高級中等學校", 
                                      "091319" = "福智學校財團法人雲林縣福智高級中等學校", 
                                      "011426" = "能仁學校財團法人新北市能仁高級家事商業職業學校", 
                                      "031415" = "桃園市方曙商工高級中等學校", 
                                      "031326" = "臺灣大華學校財團法人桃園市私立大華高級中等學校", 
                                      "031323" = "桃園市私立至善高級中學", 
                                      "031310" = "六和學校財團法人桃園市六和高級中等學校", 
                                      "031309" = "桃園育達學校財團法人桃園市育達高級中等學校", 
                                      "034332" = "桃園市立觀音高級中學", 
                                      "034314" = "桃園市立壽山高級中學", 
                                      "034335" = "桃園市立新屋高級中等學校", 
                                      "034306" = "桃園市立南崁高級中學", 
                                      "034347" = "桃園市立永豐高級中學", 
                                      "034319" = "桃園市立平鎮高級中學", 
                                      "034312" = "桃園市立大溪高級中學", 
                                      "034399" = "桃園市立大園國際高級中學", 
                                      "011316" = "格致學校財團法人新北市私立格致高級中學", 
                                      "011399" = "時雨學校財團法人新北市私立時雨高級中學", 
                                      "011310" = "剛恆毅學校財團法人新北市天主教恆毅高級中學", 
                                      "054317" = "苗栗縣立興華高級中學", 
                                      "054309" = "苗栗縣立苑裡高級中學", 
                                      "054333" = "苗栗縣立大同高級中學", 
                                      "054308" = "苗栗縣立三義高級中學", 
                                      "031301" = "泉僑學校財團法人桃園市漢英高級中等學校", 
                                      "131307" = "屏榮學校財團法人屏東縣屏榮高級中學", 
                                      "134321" = "屏東縣立枋寮高級中學", 
                                      "134324" = "屏東縣立東港高級中學", 
                                      "134334" = "屏東縣立來義高級中學", 
                                      "134304" = "屏東縣立大同高級中學", 
                                      "421303" = "臺北市私立奎山實驗高級中學", 
                                      "551303" = "高雄市私立南海月光實驗教育學校", 
                                      "081314" = "南投縣私立普台高級中學", 
                                      "081313" = "南投縣私立弘明實驗高級中學", 
                                      "081311" = "南投縣私立五育高級中學", 
                                      "081312" = "南投縣私立三育高級中學", 
                                      "084309" = "南投縣立旭光高級中學", 
                                      "011309" = "南山學校財團法人新北市南山高級中學", 
                                      "011306" = "金陵學校財團法人新北市金陵女子高級中學", 
                                      "154301" = "花蓮縣立體育高級中等學校", 
                                      "154399" = "花蓮縣立南平中學", 
                                      "031312" = "治平學校財團法人桃園市治平高級中等學校", 
                                      "011315" = "東海中學學校財團法人新北市東海高級中學", 
                                      "191309" = "明德學校財團法人臺中市明德高級中學", 
                                      "021301" = "宜蘭縣私立慧燈高級中學", 
                                      "024325" = "宜蘭縣立慈心華德福教育實驗高級中等學校", 
                                      "024322" = "宜蘭縣立南澳高級中學", 
                                      "191308" = "宜寧學校財團法人臺中市宜寧高級中學", 
                                      "191315" = "臺中市磊川華德福實驗教育學校", 
                                      "011312" = "私立崇義高中", 
                                      "051307" = "私立全人實驗高中", 
                                      "141301" = "均一學校財團法人臺東縣均一高級中等學校", 
                                      "121307" = "佛光山學校財團法人高雄市普門高級中學", 
                                      "171405" = "光隆學校財團法人基隆市光隆高級家事商職業學校", 
                                      "191412" = "光華學校財團法人臺中市光華高級工業職業學校", 
                                      "031317" = "光啟學校財團法人桃園市光啟高級中等學校", 
                                      "031421" = "永平學校財團法人桃園市私立永平工商高級中等學校", 
                                      "071318" = "正德學校財團法人彰化縣正德高級中學", 
                                      "101407" = "弘德學校財團法人嘉義縣弘德高級工商職業學校", 
                                      "151307" = "四維學校財團法人花蓮縣四維高級中學", 
                                      "011301" = "臺灣基督長老教會淡江學校財團法人新北市私立淡江高級中學", 
                                      "071311" = "彰化縣私立精誠高級中學", 
                                      "071414" = "彰化縣私立達德高級商工職業學校", 
                                      "071317" = "天主教私立文興高級中學", 
                                      "101304" = "嘉義縣私立協同高級中學", 
                                      "101303" = "嘉義縣私立同濟高級中學", 
                                      "201313" = "嘉義市私立宏仁女子高級中學", 
                                      "041303" = "新竹縣私立義民高級中學", 
                                      "041306" = "新竹縣私立東泰高級中學", 
                                      "041305" = "新竹縣私立忠信高級中學", 
                                      "041401" = "新竹縣私立內思高級工業職業學校", 
                                      "181306" = "新竹市私立曙光女子高級中學", 
                                      "181307" = "新竹市私立磐石高級中學", 
                                      "181305" = "新竹市私立光復高級中學", 
                                      "091307" = "雲林縣私立永年高級中學", 
                                      "091308" = "雲林縣私立正心高級中學", 
                                      "091311" = "雲林縣私立文生高級中學", 
                                      "091414" = "雲林縣私立大德高級工商職業學校", 
                                      "121410" = "高雄市私立旗美高級商工職業學校", 
                                      "121413" = "高雄市私立高英高級工商職業學校", 
                                      "121417" = "高雄市私立高苑高級工商職業學校", 
                                      "051413" = "苗栗縣私立龍德高級家事商業職業學校", 
                                      "051412" = "苗栗縣私立賢德工商職業學校", 
                                      "051306" = "苗栗縣私立建臺高級中學", 
                                      "051411" = "苗栗縣私立育民高級工業家事職業學校", 
                                      "051302" = "苗栗縣私立君毅高級中學", 
                                      "051408" = "苗栗縣私立中興高級商工職業學校", 
                                      "051305" = "苗栗縣私立大成高級中學", 
                                      "131311" = "屏東縣私立美和高級中學", 
                                      "131409" = "屏東縣私立民生高級家事商業職業學校", 
                                      "081409" = "同德學校財團法人南投縣同德高級中等學校", 
                                      "151306" = "花蓮縣私立海星高級中學", 
                                      "211315" = "臺南市私立瀛海高級中學", 
                                      "211419" = "臺南市私立慈幼高級工商職業學校", 
                                      "211317" = "臺南市私立崑山高級中學", 
                                      "211407" = "臺南市私立南英高級商工職業學校", 
                                      "141307" = "未來學校財團法人臺東縣育仁高級中等學校", 
                                      "141406" = "臺東縣私立公東高級工業職業學校", 
                                      "191313" = "臺中市私立曉明女子高級中學", 
                                      "191311" = "臺中市私立衛道高級中學", 
                                      "381303" = "臺北市私立大誠高級中學", 
                                      "341302" = "臺北市私立大同高級中學", 
                                      "061301" = "常春藤學校財團法人臺中市常春藤高級中學", 
                                      "111323" = "方濟會學校財團法人臺南市黎明高級中學", 
                                      "581302" = "天主教道明學校財團法人高雄市道明高級中學", 
                                      "211304" = "天主教聖功學校財團法人臺南市天主教聖功女子高級中學", 
                                      "011322" = "天主教崇光學校財團法人新北市崇光高級中學", 
                                      "031313" = "天主教振聲學校財團法人桃園市振聲高級中等學校", 
                                      "521301" = "天主教明誠學校財團法人高雄市明誠高級中學", 
                                      "011323" = "天主教光仁學校財團法人新北市光仁高級中學", 
                                      "021310" = "中道學校財團法人宜蘭縣中道高級中學", 
                                      "121405" = "中山學校財團法人高雄市中山高級工商職業學校", 
                                      "031324" = "大興學校財團法人桃園市大興高級中等學校", 
                                      "151410" = "上騰學校財團法人花蓮縣上騰高級工商職業學校", 
                                      "171306" = "二信學校財團法人基隆市二信高級中學", 
                                      "381306" = "臺北市私立靜心高級中等學校", 
                                      "094326" = "雲林縣立蔦松藝術高級中等學校", 
                                      "014399" = "新北市立豐珠中學", 
                                      "064406" = "臺中市立神岡工業高級中等學校", 
                                      "131302" = "高鳳學校財團法人屏東縣崇華高級中等學校", 
                                      "170403" = "國立臺灣海洋大學附屬基隆海事高級中等學校")

drev_person$edu_name2 <- recode_factor(drev_person$organization_id,
                                       "010301" = "國立華僑高級中等學校", 
                                       "011301" = "私立淡江高中", 
                                       "011302" = "私立康橋高中", 
                                       "011306" = "私立金陵女中", 
                                       "011307" = "新北市裕德高級中等學校", 
                                       "011309" = "財團法人南山高中", 
                                       "011310" = "財團法人恆毅高中", 
                                       "011311" = "私立聖心女中", 
                                       "011312" = "私立崇義高中", 
                                       "011314" = "財團法人中華高中", 
                                       "011315" = "私立東海高中", 
                                       "011316" = "私立格致高中", 
                                       "011317" = "私立醒吾高中", 
                                       "011318" = "私立徐匯高中", 
                                       "011322" = "新北市崇光高中", 
                                       "011323" = "私立光仁高中", 
                                       "011324" = "私立竹林高中", 
                                       "011325" = "私立及人高中", 
                                       "011329" = "財團法人辭修高中", 
                                       "011330" = "新北市林口康橋國際高中", 
                                       "011399" = "私立時雨高中", 
                                       "011405" = "私立樹人家商", 
                                       "011407" = "私立復興商工", 
                                       "011408" = "私立南強工商", 
                                       "011413" = "私立穀保家商", 
                                       "011419" = "私立開明工商", 
                                       "011420" = "私立智光商工", 
                                       "011421" = "私立清傳高商", 
                                       "011426" = "私立能仁家商", 
                                       "011427" = "私立豫章工商", 
                                       "011431" = "私立莊敬工家", 
                                       "011432" = "私立中華商海", 
                                       "013303" = "市立泰山高中", 
                                       "013304" = "市立板橋高中", 
                                       "013335" = "市立新店高中", 
                                       "013336" = "市立中和高中", 
                                       "013337" = "市立新莊高中", 
                                       "013338" = "市立新北高中", 
                                       "013339" = "市立林口高中", 
                                       "013402" = "市立瑞芳高工", 
                                       "013430" = "市立三重商工", 
                                       "013433" = "市立新北高工", 
                                       "013434" = "市立淡水商工", 
                                       "014302" = "市立海山高中", 
                                       "014311" = "市立三重高中", 
                                       "014315" = "市立永平高中", 
                                       "014322" = "市立樹林高中", 
                                       "014326" = "市立明德高中", 
                                       "014332" = "市立秀峰高中", 
                                       "014338" = "市立金山高中", 
                                       "014343" = "市立安康高中", 
                                       "014347" = "市立雙溪高中", 
                                       "014348" = "市立石碇高中", 
                                       "014353" = "市立丹鳳高中", 
                                       "014356" = "市立清水高中", 
                                       "014357" = "市立三民高中", 
                                       "014362" = "市立錦和高中", 
                                       "014363" = "市立光復高中", 
                                       "014364" = "市立竹圍高中", 
                                       "014381" = "市立北大高級中學", 
                                       "014399" = "市立豐珠中學", 
                                       "014439" = "市立鶯歌工商", 
                                       "014468" = "市立樟樹國際實中", 
                                       "311401" = "臺北市普林思頓高中", 
                                       "313301" = "市立西松高中", 
                                       "313302" = "市立中崙高中", 
                                       "321399" = "臺北市私立協和祐德高級中學", 
                                       "323301" = "市立松山高中", 
                                       "323302" = "市立永春高中", 
                                       "323401" = "市立松山家商", 
                                       "323402" = "市立松山工農", 
                                       "330301" = "國立師大附中", 
                                       "331301" = "私立延平中學", 
                                       "331302" = "私立金甌女中", 
                                       "331304" = "私立復興實驗高中", 
                                       "331402" = "私立東方工商", 
                                       "331403" = "私立喬治工商", 
                                       "331404" = "私立開平餐飲", 
                                       "333301" = "市立和平高中", 
                                       "333401" = "市立大安高工", 
                                       "341302" = "私立大同高中", 
                                       "341402" = "私立稻江護家", 
                                       "343301" = "市立中山女中", 
                                       "343302" = "市立大同高中", 
                                       "343303" = "市立大直高中", 
                                       "351301" = "私立強恕中學", 
                                       "351402" = "臺北市開南高中", 
                                       "353301" = "市立建國中學", 
                                       "353302" = "市立成功中學", 
                                       "353303" = "市立北一女中", 
                                       "361301" = "臺北市靜修高中", 
                                       "361401" = "私立稻江高商", 
                                       "363301" = "市立明倫高中", 
                                       "363302" = "市立成淵高中", 
                                       "373301" = "市立華江高中", 
                                       "373302" = "市立大理高中", 
                                       "380301" = "國立政大附中", 
                                       "381301" = "私立東山高中", 
                                       "381302" = "私立滬江高中", 
                                       "381303" = "私立大誠高中", 
                                       "381304" = "私立再興中學", 
                                       "381305" = "私立景文高中", 
                                       "381306" = "臺北市靜心高中", 
                                       "383301" = "市立景美女中", 
                                       "383302" = "市立萬芳高中", 
                                       "383401" = "市立木柵高工", 
                                       "393301" = "市立南港高中", 
                                       "393302" = "市立育成高中", 
                                       "393401" = "市立南港高工", 
                                       "401301" = "私立文德女中", 
                                       "401302" = "私立方濟中學", 
                                       "401303" = "私立達人女中", 
                                       "403301" = "市立內湖高中", 
                                       "403302" = "市立麗山高中", 
                                       "403303" = "市立南湖高中", 
                                       "403401" = "市立內湖高工", 
                                       "411301" = "私立泰北高中", 
                                       "411302" = "私立衛理女中", 
                                       "411303" = "私立華興中學", 
                                       "411401" = "私立華岡藝校", 
                                       "413301" = "市立陽明高中", 
                                       "413302" = "市立百齡高中", 
                                       "413401" = "市立士林高商", 
                                       "421301" = "私立薇閣高中", 
                                       "421302" = "臺北市幼華高中", 
                                       "421303" = "臺北市私立奎山實驗高級中學", 
                                       "421404" = "私立惇敘工商", 
                                       "423301" = "市立復興高中", 
                                       "423302" = "市立中正高中", 
                                       "030305" = "國立中央大學附屬中壢高中", 
                                       "030403" = "國立北科大附屬桃園農工", 
                                       "031301" = "私立漢英高中", 
                                       "031309" = "桃園市育達高中", 
                                       "031310" = "私立六和高中", 
                                       "031311" = "桃園市復旦高中", 
                                       "031312" = "桃園市治平高中", 
                                       "031313" = "桃園市振聲高中", 
                                       "031317" = "私立光啟高中", 
                                       "031318" = "桃園市啟英高中", 
                                       "031319" = "桃園市清華高中", 
                                       "031320" = "桃園市新興高中", 
                                       "031323" = "私立至善高中", 
                                       "031324" = "桃園市大興高中", 
                                       "031326" = "私立大華高中", 
                                       "031414" = "私立成功工商", 
                                       "031415" = "私立方曙商工", 
                                       "031421" = "私立永平工商", 
                                       "033302" = "市立龍潭高中", 
                                       "033304" = "市立桃園高中", 
                                       "033306" = "市立武陵高中", 
                                       "033316" = "市立楊梅高中", 
                                       "033325" = "市立陽明高中", 
                                       "033327" = "市立內壢高中", 
                                       "033407" = "市立中壢高商", 
                                       "033408" = "市立中壢家商", 
                                       "034306" = "市立南崁高中", 
                                       "034312" = "市立大溪高中", 
                                       "034314" = "市立壽山高中", 
                                       "034319" = "市立平鎮高中", 
                                       "034332" = "市立觀音高中", 
                                       "034335" = "市立新屋高級中等學校", 
                                       "034347" = "市立永豐高中", 
                                       "034399" = "市立大園國際高中", 
                                       "060322" = "國立興大附中", 
                                       "060323" = "國立中科實驗高級中學", 
                                       "061301" = "財團法人常春藤高中", 
                                       "061306" = "私立明台高中", 
                                       "061309" = "私立致用高中", 
                                       "061310" = "私立大明高中", 
                                       "061311" = "私立嘉陽高中", 
                                       "061313" = "私立明道高中", 
                                       "061314" = "私立僑泰高中", 
                                       "061315" = "私立華盛頓高中", 
                                       "061316" = "私立青年高中", 
                                       "061317" = "私立弘文高中", 
                                       "061318" = "私立立人高中", 
                                       "061319" = "私立玉山高中", 
                                       "061321" = "私立慈明高中", 
                                       "061322" = "華德福大地實驗學校", 
                                       "063303" = "市立大甲高中", 
                                       "063305" = "市立清水高中", 
                                       "063312" = "市立豐原高中", 
                                       "063401" = "市立豐原高商", 
                                       "063402" = "市立大甲高工", 
                                       "063404" = "市立東勢高工", 
                                       "063407" = "市立沙鹿高工", 
                                       "063408" = "市立霧峰農工", 
                                       "064308" = "市立后綜高中", 
                                       "064324" = "市立大里高中", 
                                       "064328" = "市立新社高中", 
                                       "064336" = "市立長億高中", 
                                       "064342" = "市立中港高中", 
                                       "064350" = "市立龍津高中", 
                                       "064406" = "市立神岡高工", 
                                       "190406" = "國立興大附農", 
                                       "191301" = "私立東大附中", 
                                       "191302" = "私立葳格高中", 
                                       "191305" = "私立新民高中", 
                                       "191308" = "私立宜寧高中", 
                                       "191309" = "私立明德高中", 
                                       "191311" = "私立衛道高中", 
                                       "191313" = "私立曉明女中", 
                                       "191314" = "私立嶺東高中", 
                                       "191315" = "私立磊川華德福實驗教育學校", 
                                       "191412" = "財團法人光華高工", 
                                       "193301" = "市立臺中女中", 
                                       "193302" = "市立臺中一中", 
                                       "193303" = "市立忠明高中", 
                                       "193313" = "市立西苑高中", 
                                       "193315" = "市立東山高中", 
                                       "193316" = "市立惠文高中", 
                                       "193404" = "市立臺中家商", 
                                       "193407" = "市立臺中高工", 
                                       "194303" = "市立臺中二中", 
                                       "194315" = "市立文華高中", 
                                       "110302" = "國立新豐高中", 
                                       "110308" = "國立臺南大學附中", 
                                       "110311" = "國立北門高中", 
                                       "110312" = "國立新營高中", 
                                       "110314" = "國立後壁高中", 
                                       "110315" = "國立善化高中", 
                                       "110317" = "國立新化高中", 
                                       "110328" = "國立南科國際實驗高中", 
                                       "110401" = "國立新化高工", 
                                       "110403" = "國立白河商工", 
                                       "110404" = "國立北門農工", 
                                       "110405" = "國立曾文家商", 
                                       "110406" = "國立新營高工", 
                                       "110407" = "國立玉井工商", 
                                       "110409" = "國立臺南高工", 
                                       "110410" = "國立曾文農工", 
                                       "111313" = "私立南光高中", 
                                       "111320" = "私立港明高中", 
                                       "111321" = "臺南市興國高中", 
                                       "111322" = "私立明達高中", 
                                       "111323" = "私立黎明高中", 
                                       "111326" = "私立新榮高中", 
                                       "111419" = "私立陽明工商", 
                                       "111427" = "私立育德工家", 
                                       "114306" = "市立大灣高中", 
                                       "114307" = "市立永仁高中", 
                                       "210303" = "國立臺南二中", 
                                       "210305" = "國立臺南一中", 
                                       "210306" = "國立臺南女中", 
                                       "210309" = "國立家齊高中", 
                                       "210408" = "國立臺南高商", 
                                       "210416" = "國立臺南海事", 
                                       "211301" = "私立長榮高中", 
                                       "211302" = "私立長榮女中", 
                                       "211304" = "財團法人聖功女中", 
                                       "211310" = "臺南市光華高中", 
                                       "211314" = "私立六信高中", 
                                       "211315" = "私立瀛海高中", 
                                       "211317" = "臺南市崑山高中", 
                                       "211318" = "私立德光高中", 
                                       "211320" = "財團法人慈濟高中", 
                                       "211407" = "私立南英商工", 
                                       "211412" = "私立亞洲餐旅", 
                                       "211419" = "私立慈幼工商", 
                                       "213303" = "市立南寧高中", 
                                       "213316" = "市立土城高中", 
                                       "120303" = "國立鳳山高中", 
                                       "120304" = "國立岡山高中", 
                                       "120311" = "國立旗美高中", 
                                       "120319" = "國立鳳新高中", 
                                       "120401" = "國立旗山農工", 
                                       "120402" = "國立岡山農工", 
                                       "120409" = "國立鳳山商工", 
                                       "121306" = "財團法人新光高中", 
                                       "121307" = "財團法人普門中學", 
                                       "121318" = "私立正義高中", 
                                       "121320" = "私立義大國際高中", 
                                       "121405" = "私立中山工商", 
                                       "121410" = "私立旗美商工", 
                                       "121413" = "私立高英工商", 
                                       "121415" = "私立華德工家", 
                                       "121417" = "私立高苑工商", 
                                       "124302" = "市立文山高中", 
                                       "124311" = "市立林園高中", 
                                       "124313" = "市立仁武高中", 
                                       "124322" = "市立路竹高中", 
                                       "124333" = "市立六龜高中", 
                                       "124340" = "市立福誠高中", 
                                       "521301" = "天主教明誠高中", 
                                       "521303" = "私立大榮高中", 
                                       "521401" = "私立中華藝校", 
                                       "523301" = "市立鼓山高中", 
                                       "533301" = "市立左營高中", 
                                       "533302" = "市立新莊高中", 
                                       "533401" = "市立海青工商", 
                                       "533402" = "市立三民家商", 
                                       "540301" = "國立中山大學附屬國光高中", 
                                       "543301" = "市立中山高中", 
                                       "543302" = "市立楠梓高中", 
                                       "551301" = "私立立志高中", 
                                       "551303" = "南海月光實驗學校", 
                                       "551402" = "私立樹德家商", 
                                       "553301" = "市立高雄中學", 
                                       "553302" = "市立三民高中", 
                                       "553401" = "市立高雄高工", 
                                       "563301" = "市立新興高中", 
                                       "563401" = "市立高雄高商", 
                                       "573301" = "市立高雄女中", 
                                       "580301" = "國立高師大附中", 
                                       "581301" = "私立復華高中", 
                                       "581302" = "天主教道明中學", 
                                       "581401" = "私立國際商工", 
                                       "581402" = "私立三信家商", 
                                       "583301" = "市立中正高中", 
                                       "593301" = "市立前鎮高中", 
                                       "593302" = "市立瑞祥高中", 
                                       "593401" = "市立中正高工", 
                                       "610405" = "國立高餐大附屬餐旅中學", 
                                       "613301" = "市立小港高中", 
                                       "020301" = "國立蘭陽女中", 
                                       "020302" = "國立宜蘭高中", 
                                       "020308" = "國立羅東高中", 
                                       "020403" = "國立宜蘭高商", 
                                       "020404" = "國立羅東高商", 
                                       "020405" = "國立蘇澳海事", 
                                       "020407" = "國立羅東高工", 
                                       "020409" = "國立頭城家商", 
                                       "021301" = "私立慧燈高中", 
                                       "021310" = "私立中道高中", 
                                       "024322" = "縣立南澳高中", 
                                       "024325" = "縣立慈心華德福實中", 
                                       "040302" = "國立竹東高中", 
                                       "040304" = "國立關西高中", 
                                       "040308" = "國立竹北高中", 
                                       "041303" = "私立義民高中", 
                                       "041305" = "私立忠信高中", 
                                       "041306" = "私立東泰高中", 
                                       "041307" = "私立仰德高中", 
                                       "041401" = "私立內思高工", 
                                       "044311" = "縣立六家高級中學", 
                                       "044320" = "縣立湖口高中", 
                                       "050303" = "國立苗栗高中", 
                                       "050310" = "國立竹南高中", 
                                       "050314" = "國立卓蘭高中", 
                                       "050315" = "國立苑裡高中", 
                                       "050401" = "國立大湖農工", 
                                       "050404" = "國立苗栗農工", 
                                       "050407" = "國立苗栗高商", 
                                       "051302" = "私立君毅高中", 
                                       "051305" = "私立大成高中", 
                                       "051306" = "私立建臺高中", 
                                       "051307" = "私立全人實驗高中", 
                                       "051408" = "私立中興商工", 
                                       "051411" = "私立育民工家", 
                                       "051412" = "私立賢德工商", 
                                       "051413" = "私立龍德家商", 
                                       "054308" = "縣立三義高中", 
                                       "054309" = "縣立苑裡高中", 
                                       "054317" = "縣立興華高中", 
                                       "054333" = "縣立大同高中", 
                                       "070301" = "國立彰化女中", 
                                       "070304" = "國立員林高中", 
                                       "070307" = "國立彰化高中", 
                                       "070316" = "國立鹿港高中", 
                                       "070319" = "國立溪湖高中", 
                                       "070401" = "國立彰師附工", 
                                       "070402" = "國立永靖高工", 
                                       "070403" = "國立二林工商", 
                                       "070405" = "國立秀水高工", 
                                       "070406" = "國立彰化高商", 
                                       "070408" = "國立員林農工", 
                                       "070409" = "國立員林崇實高工", 
                                       "070410" = "國立員林家商", 
                                       "070415" = "國立北斗家商", 
                                       "071311" = "私立精誠高中", 
                                       "071317" = "私立文興高中", 
                                       "071318" = "財團法人正德高中", 
                                       "071413" = "私立大慶商工", 
                                       "071414" = "私立達德商工", 
                                       "074308" = "縣立彰化藝術高中", 
                                       "074313" = "縣立二林高中", 
                                       "074323" = "縣立和美高中", 
                                       "074328" = "縣立田中高中", 
                                       "074339" = "縣立成功高中", 
                                       "080302" = "國立南投高中", 
                                       "080305" = "國立中興高中", 
                                       "080307" = "國立竹山高中", 
                                       "080308" = "國立暨大附中", 
                                       "080401" = "國立仁愛高農", 
                                       "080403" = "國立埔里高工", 
                                       "080404" = "國立南投高商", 
                                       "080406" = "國立草屯商工", 
                                       "080410" = "國立水里商工", 
                                       "081311" = "私立五育高中", 
                                       "081312" = "私立三育高中", 
                                       "081313" = "私立弘明實驗高中", 
                                       "081314" = "私立普台高中", 
                                       "081409" = "南投縣同德高中", 
                                       "084309" = "縣立旭光高中", 
                                       "090305" = "國立斗六高中", 
                                       "090306" = "國立北港高中", 
                                       "090315" = "國立虎尾高中", 
                                       "090401" = "國立虎尾農工", 
                                       "090402" = "國立西螺農工", 
                                       "090403" = "國立斗六家商", 
                                       "090404" = "國立北港農工", 
                                       "090413" = "國立土庫商工", 
                                       "091307" = "私立永年高中", 
                                       "091308" = "私立正心高中", 
                                       "091311" = "私立文生高中", 
                                       "091312" = "私立巨人高中", 
                                       "091316" = "私立揚子高中", 
                                       "091318" = "財團法人義峰高中", 
                                       "091319" = "福智高中", 
                                       "091320" = "雲林縣維多利亞實驗高中", 
                                       "091410" = "私立大成商工", 
                                       "091414" = "私立大德工商", 
                                       "094301" = "縣立斗南高中", 
                                       "094307" = "縣立麥寮高中", 
                                       "094308" = "縣立古坑華德福實驗高級中學", 
                                       "094326" = "縣立蔦松藝術高中", 
                                       "100301" = "國立東石高中", 
                                       "100302" = "國立新港藝術高中", 
                                       "100402" = "國立民雄農工", 
                                       "101303" = "私立同濟高中", 
                                       "101304" = "私立協同高中", 
                                       "101405" = "私立協志工商", 
                                       "101406" = "私立萬能工商", 
                                       "101407" = "私立弘德工商", 
                                       "104319" = "縣立竹崎高中", 
                                       "104326" = "縣立永慶高中", 
                                       "130302" = "國立屏東女中", 
                                       "130305" = "國立屏東高中", 
                                       "130306" = "國立潮州高中", 
                                       "130322" = "國立屏北高中", 
                                       "130401" = "國立內埔農工", 
                                       "130403" = "國立屏東高工", 
                                       "130404" = "國立佳冬高農", 
                                       "130410" = "國立東港海事", 
                                       "130417" = "國立恆春工商", 
                                       "131302" = "屏東縣崇華高中", 
                                       "131307" = "財團法人屏榮高中", 
                                       "131308" = "私立陸興高中", 
                                       "131311" = "私立美和高中", 
                                       "131409" = "私立民生家商", 
                                       "131418" = "私立日新工商", 
                                       "134304" = "縣立大同高中", 
                                       "134321" = "縣立枋寮高中", 
                                       "134324" = "縣立東港高中", 
                                       "134334" = "縣立來義高中", 
                                       "140301" = "國立臺東大學附屬體育高中", 
                                       "140302" = "國立臺東女中", 
                                       "140303" = "國立臺東高中", 
                                       "140404" = "國立關山工商", 
                                       "140405" = "國立臺東高商", 
                                       "140408" = "國立成功商水", 
                                       "141301" = "臺東縣均一高中", 
                                       "141307" = "私立育仁高中", 
                                       "141406" = "私立公東高工", 
                                       "144322" = "縣立蘭嶼高中", 
                                       "150302" = "國立花蓮女中", 
                                       "150303" = "國立花蓮高中", 
                                       "150309" = "國立玉里高中", 
                                       "150401" = "國立花蓮高農", 
                                       "150404" = "國立花蓮高工", 
                                       "150405" = "國立花蓮高商", 
                                       "150411" = "國立光復商工", 
                                       "151306" = "私立海星高中", 
                                       "151307" = "私立四維高中", 
                                       "151312" = "財團法人慈濟大學附中", 
                                       "151410" = "花蓮縣上騰工商", 
                                       "154301" = "花蓮縣立體育高中", 
                                       "154399" = "縣立南平中學", 
                                       "160302" = "國立馬公高中", 
                                       "160401" = "國立澎湖海事水產", 
                                       "170301" = "國立基隆女中", 
                                       "170302" = "國立基隆高中", 
                                       "170403" = "國立海洋大學附屬基隆海事", 
                                       "170404" = "國立基隆商工", 
                                       "171306" = "私立二信高中", 
                                       "171308" = "輔大聖心高中", 
                                       "171405" = "私立光隆家商", 
                                       "171407" = "私立培德工家", 
                                       "173304" = "市立中山高中", 
                                       "173306" = "市立安樂高中", 
                                       "173307" = "市立暖暖高中", 
                                       "173314" = "市立八斗高中", 
                                       "180301" = "國立科學工業園區實驗高中", 
                                       "180302" = "國立新竹女中", 
                                       "180309" = "國立新竹高中", 
                                       "180403" = "國立新竹高商", 
                                       "180404" = "國立新竹高工", 
                                       "181305" = "私立光復高中", 
                                       "181306" = "私立曙光女中", 
                                       "181307" = "私立磐石高中", 
                                       "181308" = "私立世界高中", 
                                       "183306" = "市立成德高中", 
                                       "183307" = "市立香山高中", 
                                       "183313" = "市立建功高中", 
                                       "200302" = "國立嘉義女中", 
                                       "200303" = "國立嘉義高中", 
                                       "200401" = "國立華南高商", 
                                       "200405" = "國立嘉義高工", 
                                       "200406" = "國立嘉義高商", 
                                       "200407" = "國立嘉義家職", 
                                       "201304" = "私立興華高中", 
                                       "201309" = "私立仁義高中", 
                                       "201310" = "私立嘉華高中", 
                                       "201312" = "私立輔仁高中", 
                                       "201313" = "私立宏仁女中", 
                                       "201314" = "私立立仁高中", 
                                       "201408" = "私立東吳工家", 
                                       "710301" = "國立金門高中", 
                                       "710401" = "國立金門農工", 
                                       "720301" = "國立馬祖高中", 
                                       "011C71" = "私立光華高商進修學校", 
                                       "351B09" = "私立南華高中進修學校", 
                                       "361B09" = "私立志仁中學進修學校", 
                                       "034348" = "桃園市立羅浮高級中等學校", 
                                       "121302" = "私立光禾華德福實驗學校", 
                                       "333304" = "臺北市芳和實驗中學", 
                                       "383303" = "市立數位實驗高中")

# 人事資料表合併學校科別資訊 ----------------------------------------------------------------
drev_person_1 <- data_schtype_wide %>%
  select(organization_id, typeC, typeD, typeM, typeH, typeV, typeJ, typeE, typeF) %>%
  distinct(organization_id, typeC, typeD, typeM, typeH, typeV, typeJ, typeE, typeF) %>%
  merge(x = drev_person, by = "organization_id", all.x = TRUE) %>%
  distinct()

# 人事資料表資料格式修正 ------------------------------------------------------------------

# 若比對目標為xxx，使用grep函數 
  # grep函數用法：grep("xxx", data名稱$欄位名稱)
  # grep函數中第一個參數：  
    # xxx：包含xxx的字串都會搜尋出來
    # ^xxx：搜尋開頭為xxx的字串
    # xxx$：搜尋結尾為xxx的字串
    # ^xxx$：搜尋完全符合xxx的字串

#調整英文字母大小寫
temp <- c("idnumber", "implcls", "skillteacher", "counselor", "speteacher", "joiteacher", "expecter", "study", "ddegreen1", "ddegreeu1", "ddegreeg1", "ddegreen2", "ddegreeu2", "ddegreeg2", "mdegreen1", "mdegreeu1", "mdegreeg1", "mdegreen2", "mdegreeu2", "mdegreeg2", "bdegreen1", "bdegreeu1", "bdegreeg1", "bdegreen2", "bdegreeu2", "bdegreeg2", "adegreen1", "adegreeu1", "adegreeg1", "adegreen2", "adegreeu2", "adegreeg2", "leave", "admintitle", "adminunit", "admintitle1", "adminunit1", "admintitle2", "adminunit2", "admintitle3", "adminunit3")
for (x in temp) {
  drev_person_1[[x]] <- drev_person_1[[x]] %>% toupper()
}

#刪除空格的方程式
trim <- function (x){
  gsub("\\s+", "", x)
}

#刪除空格 "\\s+"
#刪除字串前面或者字串後面的空格 "^\\s+|\\s+$"
temp <- names(drev_person)
temp1 <- c("name","ddegreen1", "ddegreeu1", "ddegreeg1", "ddegreen2", "ddegreeu2", "ddegreeg2", "mdegreen1", "mdegreeu1", "mdegreeg1", "mdegreen2", "mdegreeu2", "mdegreeg2", "bdegreen1", "bdegreeu1", "bdegreeg1", "bdegreen2", "bdegreeu2", "bdegreeg2", "adegreen1", "adegreeu1", "adegreeg1", "adegreen2", "adegreeu2", "adegreeg2", "source", "edu_name", "edu_name2")
for (i in temp1){
  temp <- temp[-which(temp == i)]
}
for (x in temp) {
  drev_person_1[[x]] <- trim(drev_person_1[[x]])
}

#職務名稱、服務單位 欄位名稱更名 方便分析
names(drev_person_1)[which(names(drev_person_1) == "admintitle")] <- "admintitle0"
names(drev_person_1)[which(names(drev_person_1) == "adminunit")] <- "adminunit0"

temp <- c("admintitle0", "adminunit0", "admintitle1", "adminunit1", "admintitle2", "adminunit2", "admintitle3", "adminunit3")
for (x in temp){
  drev_person_1[[x]][grep("^祕書$", drev_person_1[[x]])] <- "秘書"
  drev_person_1[[x]][grep("^圖書管主任$", drev_person_1[[x]])] <- "圖書館主任"
}	

#教員資料表無專任行政職欄位 調整成N 方便分析
temp <- c("admintitle0", "adminunit0")
for (x in temp){
  drev_person_1[[x]][is.na(drev_person_1[[x]])] <- "N"
}	

#國籍別、畢業學校國別
temp <- c("nation", "ddegreen1", "ddegreen2", "mdegreen1", "mdegreen2", "bdegreen1", "bdegreen2", "adegreen1", "adegreen2")
for (x in temp){
  drev_person_1[[x]][grep("^本國籍$", drev_person_1[[x]])] <- "本國"
  drev_person_1[[x]][grep("^澳洲$", drev_person_1[[x]])] <- "澳大利亞"
}	

#亂碼
drev_person_1[["mdegreeu1"]][grep("^國立台灣藝術⼤學$", drev_person_1[["mdegreeu1"]])] <- "國立台灣藝術大學"
drev_person_1[["mdegreeu1"]][grep("^私立中華科技⼤學$", drev_person_1[["mdegreeu1"]])] <- "私立中華科技大學"

drev_person_1[["mdegreeg1"]][grep("^⽣物科技系$", drev_person_1[["mdegreeg1"]])] <- "生物科技系"

drev_person_1[["bdegreeu1"]][grep("^國立台北科技⼤學$", drev_person_1[["bdegreeu1"]])] <- "國立台北科技大學"
drev_person_1[["bdegreeu1"]][grep("^⾼雄餐旅⼤學$", drev_person_1[["bdegreeu1"]])] <- "高雄餐旅大學"
drev_person_1[["bdegreeu1"]][grep("^銘傳⼤學$", drev_person_1[["bdegreeu1"]])] <- "銘傳大學"
drev_person_1[["bdegreeu1"]][grep("^德明財經科技⼤學$", drev_person_1[["bdegreeu1"]])] <- "德明財經科技大學"
drev_person_1[["bdegreeu1"]][grep("^國立台灣藝術⼤學$", drev_person_1[["bdegreeu1"]])] <- "國立台灣藝術大學"
drev_person_1[["bdegreeu1"]][grep("^吳鳳科技⼤學$", drev_person_1[["bdegreeu1"]])] <- "吳鳳科技大學"
drev_person_1[["bdegreeu1"]][grep("^弘光科技⼤學$", drev_person_1[["bdegreeu1"]])] <- "弘光科技大學"

drev_person_1[["bdegreeg1"]][grep("^⾞輛⼯程系$", drev_person_1[["bdegreeg1"]])] <- "車輛工程系"
drev_person_1[["bdegreeg1"]][grep("^⻄餐廚藝系$", drev_person_1[["bdegreeg1"]])] <- "西餐廚藝系"
drev_person_1[["bdegreeg1"]][grep("^化妝品應⽤管理系$", drev_person_1[["bdegreeg1"]])] <- "化妝品應用管理系"

#將gender由字串轉成數字
drev_person_1 <- drev_person_1 %>% 
  transform(gender = as.numeric(gender))

# 人事資料表與離退教職員(工)資料表合併 (inner join) -----------------------------------------------------------
drev_P_retire <- merge(x = drev_person_1, y = data_retire, by = c("organization_id", "idnumber"))

# 前一期人事資料表與離退教職員(工)資料表合併 (inner join) -----------------------------------------------------------
drev_P_retire_pre_inner <- merge(x = drev_person_pre, y = data_retire, by = c("organization_id", "idnumber"))
# 前一期人事資料表與離退教職員(工)資料表合併 (right join) -----------------------------------------------------------
drev_person_pre_1 <- drev_person_pre %>%
  mutate(pre = 1)
drev_P_retire_pre_right <- merge(x = drev_person_pre_1, y = data_retire, by = c("organization_id", "idnumber"), all.y = TRUE)

#學校名稱 (本次已上傳學校)
edu_name2 <- data.frame(
  "organization_id" = drev_person$organization_id, 
  "edu_name2" = drev_person$edu_name2
) %>%
  distinct(organization_id, .keep_all = TRUE)
#本期人事資料表與前一期人事資料表合併 -----------------------------------------------------------
drev_person_2 <- drev_person %>%
  mutate(now = 1)
drev_P_retire_merge_pre <- merge(x = drev_person_2, y = drev_person_pre_1, by = c("organization_id", "idnumber"), all.x = TRUE, all.y = TRUE)
#只留已上傳學校
drev_P_retire_merge_pre <- merge(x = drev_P_retire_merge_pre, y = edu_name2, by = c("organization_id"))
#再與本次離退表合併
data_retire_1 <- data_retire %>%
  mutate(retire = 1)
drev_P_retire_merge_pre <- merge(x = drev_P_retire_merge_pre, y = data_retire_1, by = c("organization_id", "idnumber"), all.x = TRUE, all.y = TRUE)

# 人事資料表與教學資料表合併 -----------------------------------------------------------
drev_P_load <- merge(x = drev_person_1, y = data_load, by = c("organization_id", "idnumber"), all.x = TRUE, all.y = TRUE)

# 教學資料表資料格式修正 ------------------------------------------------------------------
temp <- c("tutor", "mitleader", "classleader", "ccounselor", "adminteacher", "specurr")
for (x in temp) {
  drev_P_load[[x]] <- drev_P_load[[x]] %>% toupper()
}

temp <- c("mainsub", "tutor", "mitleader", "classleader", "ccounselor", "adminteacher", "specurr")
for (x in temp) {
  drev_P_load[[x]] <- trim(drev_P_load[[x]])
}

#主要任教科別
drev_P_load[["mainsub"]][grep("^汽事科$", drev_P_load[["mainsub"]])] <- "汽車科"

#所擔任召集人之學科名稱
drev_P_load[["mitleader"]][grep("^汽事科$", drev_P_load[["mitleader"]])] <- "汽車科"

temp <- c("basic", "cut", "othertime", "othertimeb")
for (x in temp) {
  drev_P_load[[x]][grep("^NA$", drev_P_load[[x]])] <- NA
}


# 主要任教科目 ------------------------------------------------------------------
drev_P_load$depcode_p <- case_when(
  drev_P_load$mainsub == "機械群" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "101",
  drev_P_load$mainsub == "動力機械群" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "102",
  drev_P_load$mainsub == "化工群" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "103",
  drev_P_load$mainsub == "電機與電子群" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "104",
  drev_P_load$mainsub == "土木與建築群" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "105",
  drev_P_load$mainsub == "商業與管理群" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "106",
  drev_P_load$mainsub == "外語群" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "107",
  drev_P_load$mainsub == "設計群" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "108",
  drev_P_load$mainsub == "藝術群" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "109",
  drev_P_load$mainsub == "農業群" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "110",
  drev_P_load$mainsub == "食品群" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "111",
  drev_P_load$mainsub == "餐旅群" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "112",
  drev_P_load$mainsub == "家政群" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "113",
  drev_P_load$mainsub == "海事群" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "114",
  drev_P_load$mainsub == "水產群" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "115",
  drev_P_load$mainsub == "農場經營科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "201",
  drev_P_load$mainsub == "園藝科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "202",
  drev_P_load$mainsub == "森林科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "204",
  drev_P_load$mainsub == "農業機械科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "205",
  drev_P_load$mainsub == "食品加工科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "206",
  drev_P_load$mainsub == "野生動物保育科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "214",
  drev_P_load$mainsub == "農產行銷科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "215",
  drev_P_load$mainsub == "造園科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "216",
  drev_P_load$mainsub == "畜產保健科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "217",
  drev_P_load$mainsub == "機械科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "301",
  drev_P_load$mainsub == "鑄造科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "302",
  drev_P_load$mainsub == "汽車科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "303",
  drev_P_load$mainsub == "板金科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "304",
  drev_P_load$mainsub == "資訊科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "305",
  drev_P_load$mainsub == "電子科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "306",
  drev_P_load$mainsub == "控制科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "307",
  drev_P_load$mainsub == "電機科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "308",
  drev_P_load$mainsub == "冷凍空調科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "309",
  drev_P_load$mainsub == "建築科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "311",
  drev_P_load$mainsub == "家具木工科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "312",
  drev_P_load$mainsub == "化工科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "315",
  drev_P_load$mainsub == "美工科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "316",
  drev_P_load$mainsub == "美術工藝科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "318",
  drev_P_load$mainsub == "紡織科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "319",
  drev_P_load$mainsub == "電機空調科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "321",
  drev_P_load$mainsub == "機械木模科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "332",
  drev_P_load$mainsub == "配管科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "337",
  drev_P_load$mainsub == "模具科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "338",
  drev_P_load$mainsub == "染整科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "352",
  drev_P_load$mainsub == "機電科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "360",
  drev_P_load$mainsub == "陶瓷工程科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "361",
  drev_P_load$mainsub == "製圖科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "363",
  drev_P_load$mainsub == "重機科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "364",
  drev_P_load$mainsub == "土木科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "365",
  drev_P_load$mainsub == "室內空間設計科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "366",
  drev_P_load$mainsub == "環境檢驗科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "367",
  drev_P_load$mainsub == "生物產業機電科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "372",
  drev_P_load$mainsub == "圖文傳播科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "373",
  drev_P_load$mainsub == "電腦機械製圖科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "374",
  drev_P_load$mainsub == "軌道車輛科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "380",
  drev_P_load$mainsub == "飛機修護科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "381",
  drev_P_load$mainsub == "飛修科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "381",
  drev_P_load$mainsub == "航空電子科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "384",
  drev_P_load$mainsub == "動力機械科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "392",
  drev_P_load$mainsub == "金屬工藝科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "394",
  drev_P_load$mainsub == "消防工程科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "397",
  drev_P_load$mainsub == "空間測繪科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "398",
  drev_P_load$mainsub == "家具設計科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "399",
  drev_P_load$mainsub == "商業經營科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "401",
  drev_P_load$mainsub == "商經科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "401",
  drev_P_load$mainsub == "國際貿易科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "402",
  drev_P_load$mainsub == "國貿科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "402",
  drev_P_load$mainsub == "會計事務科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "403",
  drev_P_load$mainsub == "資料處理科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "404",
  drev_P_load$mainsub == "資處科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "404",
  drev_P_load$mainsub == "廣告設計科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "406",
  drev_P_load$mainsub == "觀光事業科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "407",
  drev_P_load$mainsub == "餐飲管理科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "408",
  drev_P_load$mainsub == "不動產事務科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "418",
  drev_P_load$mainsub == "電子商務科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "425",
  drev_P_load$mainsub == "流通管理科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "426",
  drev_P_load$mainsub == "多媒體設計科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "430",
  drev_P_load$mainsub == "多媒體應用科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "431",
  drev_P_load$mainsub == "應用英語科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "433",
  drev_P_load$mainsub == "應用日語科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "434",
  drev_P_load$mainsub == "家政科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "501",
  drev_P_load$mainsub == "服裝科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "502",
  drev_P_load$mainsub == "幼兒保育科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "503",
  drev_P_load$mainsub == "美容科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "504",
  drev_P_load$mainsub == "食品科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "505",
  drev_P_load$mainsub == "室內設計科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "512",
  drev_P_load$mainsub == "時尚模特兒科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "513",
  drev_P_load$mainsub == "流行服飾科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "515",
  drev_P_load$mainsub == "時尚造型科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "516",
  drev_P_load$mainsub == "烘焙科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "517",
  drev_P_load$mainsub == "漁業科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "701",
  drev_P_load$mainsub == "輪機科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "702",
  drev_P_load$mainsub == "電子通信科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "703",
  drev_P_load$mainsub == "水產養殖科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "705",
  drev_P_load$mainsub == "水產經營科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "706",
  drev_P_load$mainsub == "航海科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "708",
  drev_P_load$mainsub == "航運管理科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "717",
  drev_P_load$mainsub == "水產食品科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "718",
  drev_P_load$mainsub == "戲劇科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "801",
  drev_P_load$mainsub == "音樂科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "802",
  drev_P_load$mainsub == "舞蹈科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "803",
  drev_P_load$mainsub == "美術科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "804",
  drev_P_load$mainsub == "影劇科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "806",
  drev_P_load$mainsub == "西樂科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "807",
  drev_P_load$mainsub == "國樂科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "808",
  drev_P_load$mainsub == "劇場藝術科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "813",
  drev_P_load$mainsub == "電影電視科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "816",
  drev_P_load$mainsub == "表演藝術科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "817",
  drev_P_load$mainsub == "多媒體動畫科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "820",
  drev_P_load$mainsub == "時尚工藝科" & (drev_P_load$typeM == 1 | drev_P_load$typeV == 1 | drev_P_load$typeC == 1) ~ "822",
  (drev_P_load$mainsub == "資訊應用學程" | drev_P_load$mainsub == "資訊應用科") & drev_P_load$typeM == 1 ~ "F01",
  drev_P_load$mainsub == "國語文" | drev_P_load$mainsub == "國文" | drev_P_load$mainsub == "國文科" ~ "01",
  drev_P_load$mainsub == "英語文" | drev_P_load$mainsub == "英語" | drev_P_load$mainsub == "英文" | drev_P_load$mainsub == "英文科" ~ "02",
  drev_P_load$mainsub == "數學" | drev_P_load$mainsub == "數學科" ~ "03",
  drev_P_load$mainsub == "歷史" | drev_P_load$mainsub == "歷史科" ~ "09",
  drev_P_load$mainsub == "地理" | drev_P_load$mainsub == "地理科" ~ "0A",
  drev_P_load$mainsub == "公民與社會" | drev_P_load$mainsub == "公民與社會科" | drev_P_load$mainsub == "公民" | drev_P_load$mainsub == "公民科" ~ "0B",
  drev_P_load$mainsub == "物理" | drev_P_load$mainsub == "物理科" ~ "0C",
  drev_P_load$mainsub == "化學" | drev_P_load$mainsub == "化學科" ~ "0D",
  drev_P_load$mainsub == "生物" | drev_P_load$mainsub == "生物科" ~ "0E",
  grepl("^地球科學", drev_P_load$mainsub) | drev_P_load$mainsub == "地科" | drev_P_load$mainsub == "地科科" ~ "0F",
  drev_P_load$mainsub == "音樂" | (drev_P_load$mainsub == "音樂科" & (drev_P_load$typeM == 0 | drev_P_load$typeV == 0 | drev_P_load$typeC == 0)) ~ "0U",
  drev_P_load$mainsub == "美術" | (drev_P_load$mainsub == "美術科" & (drev_P_load$typeM == 0 | drev_P_load$typeV == 0 | drev_P_load$typeC == 0)) ~ "0V",
  drev_P_load$mainsub == "藝術生活" | drev_P_load$mainsub == "藝術生活科" | drev_P_load$mainsub == "藝術科" | drev_P_load$mainsub == "視覺應用" | drev_P_load$mainsub == "音樂應用" | drev_P_load$mainsub == "表演藝術" ~ "0W",
  drev_P_load$mainsub == "生命教育" ~ "0X",
  drev_P_load$mainsub == "生涯規劃" ~ "0Y",
  drev_P_load$mainsub == "家政" | (drev_P_load$mainsub == "家政科" & (drev_P_load$typeM == 0 | drev_P_load$typeV == 0 | drev_P_load$typeC == 0)) ~ "0Z",
  drev_P_load$mainsub == "法律與生活" | drev_P_load$mainsub == "法律與生活科" ~ "10",
  drev_P_load$mainsub == "環境科學概論" ~ "11",
  drev_P_load$mainsub == "生活科技" | drev_P_load$mainsub == "生活科技科" | drev_P_load$mainsub == "生科科" ~ "12",
  drev_P_load$mainsub == "資訊科技" | (drev_P_load$mainsub == "資訊科" & (drev_P_load$typeM == 0 | drev_P_load$typeV == 0 | drev_P_load$typeC == 0)) ~ "13",
  grepl("健康與護理", drev_P_load$mainsub) | grepl("健康護理", drev_P_load$mainsub) | drev_P_load$mainsub == "健護" | drev_P_load$mainsub == "健護科" | drev_P_load$mainsub == "護理科" ~ "14",
  drev_P_load$mainsub == "體育" | drev_P_load$mainsub == "體育科" ~ "15",
  grepl("全民國防", drev_P_load$mainsub) | drev_P_load$mainsub == "軍訓科" | drev_P_load$mainsub == "軍訓" | drev_P_load$mainsub == "國防教育科" ~ "16",
  grepl("特殊教育", drev_P_load$mainsub) | drev_P_load$mainsub == "特教科" | drev_P_load$mainsub == "特教" | drev_P_load$mainsub == "特殊需求領域(身心障礙)" | drev_P_load$mainsub == "身心障礙" | drev_P_load$mainsub == "資賦優異" ~ "929"
)

# 需要每個學期重新調整的項目 -----------------------------------------------------------

### flag_person
# flag2、flag3；spe2、spe6。
# 以上檢查項目依最新的學校名單、群科開設狀況而定。相關資訊可上統計處查詢高級中等學校科別資料。
# flag6需檢查各表姓名是否為純中文或純英文，或者是否夾雜其他運算字元、特殊符號。
# flag8需檢查持外來人口統一證號的教職員(工)是否有填其國籍別，又其國籍別是否足以辨認。
# flag9需檢查最高學歷畢業學校國別（一）(schooln1)所填之國籍別是否足以辨認。
# flag_person <- drev_person_1 %>%
#   mutate(err_flag_2 = if_else((organization_id == "011315" | organization_id == "013430" | organization_id == "110409" | organization_id == "193404" | organization_id == "381305" | organization_id == "533402"), 1, 0),
#          err_flag_3 = if_else(organization_id == "110302", 1, 0),
#          err_spe_2  = if_else(typeD == 1 & empunit != "雙語部" & source == "教員資料表", 1, 0),
#          err_spe_6  = if_else(typeJ == 1, 1, 0),
#          err_flag_6 = if_else(name == "吳淑貞-", 1, 0),
#          err_flag_8 = if_else(nation == "外籍", 1, 0),
#          err_flag_9 = 0)
# 目前僅限於人事資料表範圍內的檢查項目暫無需檢查是否擔任「科主任」、「學程主任」，
# 因此在備分的檔案(flag_person)先把這兩個職稱拿掉。若往後有此需要，請再另行處理。
# temp <- seq(from = 18, to = 25 , by = 2)
# for (x in temp){
#   flag_person[grep("$科主任", flag_person[x]), ] <- flag_person %>%
#     slice(grep("$科主任", flag_person[x])) %>%
#     mutate(err_flag_2, recode(err_flag_2, "1 = 0"))
#   
#   flag_person[grep("$學程主任", flag_person[x]), ] <- flag_person %>%
#     slice(grep("$學程主任", flag_person[x])) %>%
#     mutate(err_flag_3, recode(err_flag_3, "1 = 0"))
# }


# 人事 -----------------------------------------------------------

# flag1: 學校（副）校長、一級單位主管名單的完整度 -------------------------------------------------------------------
# temp <- c("admintitle0", "admintitle1", "admintitle2", "admintitle3")
# for (x in temp){
#   flag_person[[x]] <- drev_person_1[[x]] %>%
#     gsub(pattern = "科主任", replacement = "") %>%
#     gsub(pattern = "主任教官", replacement = "") %>%
#     gsub(pattern = "學程主任", replacement = "")
# }

flag_person <- drev_person_1 %>%
  mutate(admin1 = 0, admin2 = 0, admin3 = 0, admin4 = 0, admin5 = 0, admin6 = 0, admin7 = 0, admin8 = 0, admin9 = 0)

flag_person$admin1 <- case_when(
  flag_person$sertype == "校長" | flag_person$admintitle1 == "校長" | flag_person$admintitle1 == "校長1" | flag_person$admintitle2 == "校長1" | flag_person$admintitle3 == "校長1" ~ 1,
  TRUE ~ flag_person$admin1
)

flag_person$admin7 <- case_when(
  flag_person$typeV == 0 ~ 1,
  TRUE ~ flag_person$admin7
)

temp <- c("0", "1", "2", "3")
for (x in temp){
  flag_person$admin2 <- case_when(
     grepl("教務", flag_person[[paste("adminunit", x, sep = "")]])                                                                                                                                     & (grepl("主任$", flag_person[[paste("admintitle", x, sep = "")]]) | grepl("主任1$", flag_person[[paste("admintitle", x, sep = "")]])) & !grepl("主任教官", flag_person[[paste("admintitle", x, sep = "")]]) & !grepl("科主任", flag_person[[paste("admintitle", x, sep = "")]]) & !grepl("學程主任", flag_person[[paste("admintitle", x, sep = "")]]) ~ 1,
    TRUE ~ flag_person$admin2
  )
}     
for (x in temp){
  flag_person$admin3 <- case_when(
    (grepl("學務", flag_person[[paste("adminunit", x, sep = "")]]) | grepl("學生事務", flag_person[[paste("adminunit", x, sep = "")]]))                                                                & (grepl("主任$", flag_person[[paste("admintitle", x, sep = "")]]) | grepl("主任1$", flag_person[[paste("admintitle", x, sep = "")]])) & !grepl("主任教官", flag_person[[paste("admintitle", x, sep = "")]]) & !grepl("科主任", flag_person[[paste("admintitle", x, sep = "")]]) & !grepl("學程主任", flag_person[[paste("admintitle", x, sep = "")]]) ~ 1,
    TRUE ~ flag_person$admin3
  )
}                
for (x in temp){
  flag_person$admin4 <- case_when(
     grepl("總務", flag_person[[paste("adminunit", x, sep = "")]])                                                                                                                                     & (grepl("主任$", flag_person[[paste("admintitle", x, sep = "")]]) | grepl("主任1$", flag_person[[paste("admintitle", x, sep = "")]])) & !grepl("主任教官", flag_person[[paste("admintitle", x, sep = "")]]) & !grepl("科主任", flag_person[[paste("admintitle", x, sep = "")]]) & !grepl("學程主任", flag_person[[paste("admintitle", x, sep = "")]]) ~ 1,
    TRUE ~ flag_person$admin4
  )
}                
for (x in temp){
  flag_person$admin5 <- case_when(
     grepl("輔導", flag_person[[paste("adminunit", x, sep = "")]])                                                                                                                                     & (grepl("主任$", flag_person[[paste("admintitle", x, sep = "")]]) | grepl("主任1$", flag_person[[paste("admintitle", x, sep = "")]])) & !grepl("主任輔導教師$", flag_person[[paste("admintitle", x, sep = "")]]) & !grepl("主任教官", flag_person[[paste("admintitle", x, sep = "")]]) & !grepl("科主任", flag_person[[paste("admintitle", x, sep = "")]]) & !grepl("學程主任", flag_person[[paste("admintitle", x, sep = "")]]) ~ 1,
    TRUE ~ flag_person$admin5
  )
}                
for (x in temp){
  flag_person$admin6 <- case_when(
     (grepl("圖書", flag_person[[paste("adminunit", x, sep = "")]]) | grepl("圖資", flag_person[[paste("adminunit", x, sep = "")]]) | grepl("圖書資訊", flag_person[[paste("adminunit", x, sep = "")]])) & ((grepl("主任$", flag_person[[paste("admintitle", x, sep = "")]]) | grepl("主任1$", flag_person[[paste("admintitle", x, sep = "")]])) & !grepl("主任教官", flag_person[[paste("admintitle", x, sep = "")]]) & !grepl("科主任", flag_person[[paste("admintitle", x, sep = "")]]) & !grepl("學程主任", flag_person[[paste("admintitle", x, sep = "")]])) | grepl("^館長$", flag_person[[paste("admintitle", x, sep = "")]]) ~ 1,
    TRUE ~ flag_person$admin6
  )
}  
for (x in temp){
  flag_person$admin7 <- case_when(
     grepl("實習", flag_person[[paste("adminunit", x, sep = "")]]) & flag_person$typeV == 1                                                                                                            & (grepl("主任$", flag_person[[paste("admintitle", x, sep = "")]]) | grepl("主任1$", flag_person[[paste("admintitle", x, sep = "")]])) & !grepl("主任教官", flag_person[[paste("admintitle", x, sep = "")]]) & !grepl("科主任", flag_person[[paste("admintitle", x, sep = "")]]) & !grepl("學程主任", flag_person[[paste("admintitle", x, sep = "")]]) ~ 1,
    TRUE ~ flag_person$admin7
  )
}     
for (x in temp){  
  flag_person$admin8 <- case_when(
    grepl("人事", flag_person[[paste("adminunit", x, sep = "")]])                                                                                                                                      & ((grepl("主任$", flag_person[[paste("admintitle", x, sep = "")]]) | grepl("主任1$", flag_person[[paste("admintitle", x, sep = "")]])) & !grepl("主任教官", flag_person[[paste("admintitle", x, sep = "")]]) & !grepl("科主任", flag_person[[paste("admintitle", x, sep = "")]]) & !grepl("學程主任", flag_person[[paste("admintitle", x, sep = "")]])) | grepl("^人事管理員$", flag_person[[paste("admintitle", x, sep = "")]]) ~ 1,
    TRUE ~ flag_person$admin8
  )
}                
for (x in temp){
  flag_person$admin9 <- case_when(
    (grepl("會計", flag_person[[paste("adminunit", x, sep = "")]]) | grepl("主計", flag_person[[paste("adminunit", x, sep = "")]]))                                                                    & ((grepl("主任$", flag_person[[paste("admintitle", x, sep = "")]]) | grepl("主任1$", flag_person[[paste("admintitle", x, sep = "")]])) & !grepl("主任教官", flag_person[[paste("admintitle", x, sep = "")]]) & !grepl("科主任", flag_person[[paste("admintitle", x, sep = "")]]) & !grepl("學程主任", flag_person[[paste("admintitle", x, sep = "")]])) | grepl("^主計員$", flag_person[[paste("admintitle", x, sep = "")]]) | grepl("^主計員$", flag_person[[paste("admintitle", x, sep = "")]]) ~ 1,
    TRUE ~ flag_person$admin9
  )
}  

#NOTE:實習處主任設置條件為：1.只有專業群科 or 2.只有專業群科和綜合高中 or 3.只有專業群科和實用技能學程
#NOTE:2021/04/20討論：只要有專業群科的學校，都需檢查有無實習處，其他類別不管

flag1 <- flag_person %>%
  group_by(organization_id) %>%
  mutate(admin1 = max(admin1),
         admin2 = max(admin2),
         admin3 = max(admin3),
         admin4 = max(admin4),
         admin5 = max(admin5),
         admin6 = max(admin6),
         admin7 = max(admin7),
         admin8 = max(admin8),
         admin9 = max(admin9)) %>%
  mutate(admin1_txt = if_else(admin1 == 0, "校長", ""),
         admin2_txt = if_else(admin2 == 0, "教務處主管", ""),
         admin3_txt = if_else(admin3 == 0, "學務處主管", ""),
         admin4_txt = if_else(admin4 == 0, "總務處主管", ""),
         admin5_txt = if_else(admin5 == 0, "輔導室主管", ""),
         admin6_txt = if_else(admin6 == 0, "圖書館主管", ""),
         admin7_txt = if_else(admin7 == 0, "實習處主管", ""),
         admin8_txt = if_else(admin8 == 0, "人事室主管", ""),
         admin9_txt = if_else(admin9 == 0, "主（會）計室主管", "")) %>%
  mutate(flag1 = paste("尚待增補之學校主管：", admin1_txt, admin2_txt, admin3_txt, admin4_txt, admin5_txt, admin6_txt, admin7_txt, admin8_txt, admin9_txt, sep = " ")) %>%
  mutate(flag1 = recode(gsub("\\s+", " ", flag1), `尚待增補之學校主管： ` = "")) %>%
  mutate(flag1 = if_else(flag1 != "", paste(flag1, "（請確認是否填報完整名單，倘貴校###主任尚未到職，請來電告知）", sep = ""), flag1)) %>%
  subset(select = c(organization_id, flag1)) %>%
  distinct(organization_id, flag1)

flag1$flag1 <- gsub("： ", replacement="：", flag1$flag1)
flag1$flag1 <- gsub(" （", replacement="（", flag1$flag1)


#偵測flag1是否存在。若不存在，則產生NA行
if('flag1' %in% ls()){
  print("flag1")
}else{
  flag1 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag1$flag1 <- ""
}

# flag2: 設有專業類科二科以上的高級中等學校，本校應設有科主任（或有同類學程時應設有學程主任）。 -------------------------------------------------------------------
  #flag2_append-------------------------------------------------------------------
data_schtype_wide_flag2 <- aggregate(typeV ~ organization_id, data_schtype_wide, sum) %>%
  subset(typeV >= 2) %>%
  rename(flag_typeV = typeV)

flag_person <- drev_person_1
flag_person <- merge(x = flag_person, y = data_schtype_wide_flag2, by = c("organization_id"), all.x = TRUE)

flag_person$err_flag <- case_when(
  flag_person$flag_typeV >= 2 ~ 1,
  TRUE ~ 0
)
  #---

temp <- c("0", "1", "2", "3")
for (x in temp){
  flag_person$err_flag <- case_when(
    (grepl("科主任", flag_person[[paste("admintitle", x, sep = "")]]) | (grepl("主任", flag_person[[paste("admintitle", x, sep = "")]]) & grepl("科$", flag_person[[paste("adminunit", x, sep = "")]]))) & flag_person$flag_typeV >= 2 ~ 0,
    TRUE ~ flag_person$err_flag
  )
}  

flag2 <- flag_person %>%
  group_by(organization_id) %>%
  mutate(err_flag = min(err_flag)) %>%
  mutate(flag2 = if_else(err_flag == 1,"請學校確認是否設置科主任或學程主任", "")) %>%
  subset(select = c(organization_id, flag2)) %>%
  distinct(organization_id, flag2)

#偵測flag2是否存在。若不存在，則產生NA行
if('flag2' %in% ls()){
  print("flag2")
}else{
  flag2 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag2$flag2 <- ""
}

# flag3: 設有專門學程總班級數四班以上的高級中等學校，本校應設有學程主任。 -------------------------------------------------------------------
  #flag3_append-------------------------------------------------------------------
data_schtype_wide_flag3 <- data_schtype_wide %>%
  subset(dep1_code != "11")
  
data_schtype_wide_flag3 <- aggregate(nclass ~ organization_id + type_code + type_name, data_schtype_wide_flag3, sum) %>%
  rename(flag_nclass = nclass) %>%
  subset(type_code == "M" & flag_nclass >= 4) %>%
  distinct(organization_id, .keep_all = TRUE)

flag_person <- drev_person_1
flag_person <- merge(x = flag_person, y = data_schtype_wide_flag3, by = c("organization_id"), all.x = TRUE)

flag_person$err_flag <- case_when(
  (flag_person$type_code == "M" & flag_person$flag_nclass >= 4) ~ 1,
  TRUE ~ 0
)
  #---

temp <- c("0", "1", "2", "3")
for (x in temp){
  flag_person$err_flag <- case_when(
    grepl("學程主任", flag_person[[paste("admintitle", x, sep = "")]]) & (flag_person$type_code == "M" & flag_person$flag_nclass >= 4) ~ 0,
    TRUE ~ flag_person$err_flag
  )
}  

flag3 <- flag_person %>%
  group_by(organization_id) %>%
  mutate(err_flag = min(err_flag)) %>%
  mutate(flag3 = if_else(err_flag == 1,"請學校確認是否設置學程主任", "")) %>%
  subset(select = c(organization_id, flag3)) %>%
  distinct(organization_id, flag3)

#偵測flag3是否存在。若不存在，則產生NA行
if('flag3' %in% ls()){
  print("flag3")
}else{
  flag3 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag3$flag3 <- ""
}

# flag6: 人事資料表的姓名應為中文或英文，不得有亂碼。 -------------------------------------------------------------------
  #flag6_append-------------------------------------------------------------------
flag_person <- drev_person_1

#檢視姓名欄位字元數不為3
view_flag6 <- distinct(flag_person, name, .keep_all = TRUE) %>%
  subset(nchar(name) != 3) %>%
  subset(select = c(organization_id, idnumber, name, edu_name2, source))

#數字、特殊符號標記為1
flag_person$err_flag <- case_when(
  grepl("\\d", flag_person$name) ~ 1,
  grepl("[[:punct:]]", flag_person$name) ~ 1,
  TRUE ~ 0
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag6 <- flag_person %>%
  subset(select = c(organization_id, idnumber, name, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ name, value.var = "name")

#合併所有name
temp <- colnames(flag_person_wide_flag6)[3 : length(colnames(flag_person_wide_flag6))]
flag_person_wide_flag6$flag6_r <- NA
for (i in temp){
  flag_person_wide_flag6$flag6_r <- paste(flag_person_wide_flag6$flag6_r, flag_person_wide_flag6[[i]], sep = " ")
}
flag_person_wide_flag6$flag6_r <- gsub("NA ", replacement="", flag_person_wide_flag6$flag6_r)
flag_person_wide_flag6$flag6_r <- gsub(" NA", replacement="", flag_person_wide_flag6$flag6_r)

#產生檢誤報告文字
flag6_temp <- flag_person_wide_flag6 %>%
  group_by(organization_id) %>%
  mutate(flag6_txt = paste(source, "需修改姓名處：", flag6_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag6_txt)) %>%
  distinct(organization_id, flag6_txt)

#根據organization_id，展開成寬資料(wide)
flag6 <- flag6_temp %>%
  dcast(organization_id ~ flag6_txt, value.var = "flag6_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag6)[2 : length(colnames(flag6))]
flag6$flag6 <- NA
for (i in temp){
  flag6$flag6 <- paste(flag6$flag6, flag6[[i]], sep = "； ")
}
flag6$flag6 <- gsub("NA； ", replacement="", flag6$flag6)
flag6$flag6 <- gsub("； NA", replacement="", flag6$flag6)

#產生檢誤報告文字
flag6 <- flag6 %>%
  subset(select = c(organization_id, flag6)) %>%
  distinct(organization_id, flag6)
}else{
#偵測flag6是否存在。若不存在，則產生NA行
if('flag6' %in% ls()){
  print("flag6")
}else{
  flag6 <- drev_person_1 %>%
  distinct(organization_id, .keep_all = TRUE) %>%
  subset(select = c(organization_id))
  flag6$flag6 <- ""
}
}
# flag7: 出生年月日換算成年齡後，過高或過低。 -------------------------------------------------------------------
flag_person <- drev_person_1

flag_person$survey_year <- 110
flag_person$birthy <- substr(flag_person$birthdate, 1, 3) %>% as.numeric()

flag_person$age <- flag_person$survey_year- flag_person$birthy

#錯誤標記
flag_person$irr_year <- 0
flag_person$irr_year <- if_else(flag_person$age < 18, 1, flag_person$irr_year)
flag_person$irr_year <- if_else(flag_person$age > 75, 1, flag_person$irr_year)
flag_person$irr_year <- if_else(flag_person$age > 75 & (flag_person$emptype == "兼任" | flag_person$emptype == "長期代課" | flag_person$emptype == "專職族語老師" | flag_person$emptype == "鐘點教師" | flag_person$emptype == "約聘僱" | flag_person$emptype == "約用" | flag_person$emptype == "派遣"), 0, flag_person$irr_year)
flag_person$irr_year <- if_else(flag_person$age > 85 & (flag_person$emptype == "兼任" | flag_person$emptype == "長期代課" | flag_person$emptype == "專職族語老師" | flag_person$emptype == "鐘點教師" | flag_person$emptype == "約聘僱" | flag_person$emptype == "約用" | flag_person$emptype == "派遣"), 1, flag_person$irr_year)

flag_person$name <- if_else(flag_person$irr_year == 1, paste(flag_person$name, "（", flag_person$birthdate, "）", sep = ""), flag_person$name)

#姓名加註出生年月日
flag_person$name <- case_when(
  flag_person$irr_year == 1 ~ paste(flag_person$name, "（", flag_person$birthdate, "）", sep = ""),
  TRUE ~ flag_person$name
)

if (dim(flag_person %>% subset(irr_year == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_flag7 <- flag_person %>%
  subset(select = c(organization_id, idnumber, name, edu_name2, source, irr_year)) %>%
  subset(irr_year == 1) %>%
  dcast(organization_id + source ~ name, value.var = "name")

#合併所有name
temp <- colnames(flag_person_flag7)[3 : length(colnames(flag_person_flag7))]
flag_person_flag7$flag7_r <- NA
for (i in temp){
  flag_person_flag7$flag7_r <- paste(flag_person_flag7$flag7_r, flag_person_flag7[[i]], sep = " ")
}
flag_person_flag7$flag7_r <- gsub("NA ", replacement="", flag_person_flag7$flag7_r)
flag_person_flag7$flag7_r <- gsub(" NA", replacement="", flag_person_flag7$flag7_r)

#產生檢誤報告文字
flag7_temp <- flag_person_flag7 %>%
  group_by(organization_id) %>%
  mutate(flag7_txt = paste(source, "：", flag7_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag7_txt)) %>%
  distinct(organization_id, flag7_txt)

#根據organization_id，展開成寬資料(wide)
flag7 <- flag7_temp %>%
  dcast(organization_id ~ flag7_txt, value.var = "flag7_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag7)[2 : length(colnames(flag7))]
flag7$flag7 <- NA
for (i in temp){
  flag7$flag7 <- paste(flag7$flag7, flag7[[i]], sep = "； ")
}
flag7$flag7 <- gsub("NA； ", replacement="", flag7$flag7)
flag7$flag7 <- gsub("； NA", replacement="", flag7$flag7)

#產生檢誤報告文字
flag7 <- flag7 %>%
  subset(select = c(organization_id, flag7)) %>%
  distinct(organization_id, flag7) %>%
  mutate(flag7 = paste(flag7, "（請確認出生年月日是否正確）", sep = ""))
}else{
#偵測flag7是否存在。若不存在，則產生NA行
if('flag7' %in% ls()){
  print("flag7")
}else{
  flag7 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag7$flag7 <- ""
}
}
# flag8: 國籍別應填入「本國籍」或者外交部網站之世界各國名稱一覽表的國家名稱（或者至少須足以辨識國家）。 -------------------------------------------------------------------
  #flag8_append-------------------------------------------------------------------
flag_person <- drev_person_1

  #檢視國籍別欄位字元數不為3
view_flag8 <- distinct(flag_person, nation, .keep_all = TRUE) %>%
  subset(nchar(nation) != 3) %>%
  subset(select = c(organization_id, idnumber, nation, edu_name2, source))

  #"外籍"標記為1
flag_person$err_flag <- case_when(
  flag_person$nation == "外籍" ~ 1,
  TRUE ~ 0
)
  #---

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag8 <- flag_person %>%
  subset(select = c(organization_id, idnumber, name, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ name, value.var = "name")

#合併所有name
temp <- colnames(flag_person_wide_flag8)[3 : length(colnames(flag_person_wide_flag8))]
flag_person_wide_flag8$flag8_r <- NA
for (i in temp){
  flag_person_wide_flag8$flag8_r <- paste(flag_person_wide_flag8$flag8_r, flag_person_wide_flag8[[i]], sep = " ")
}
flag_person_wide_flag8$flag8_r <- gsub("NA ", replacement="", flag_person_wide_flag8$flag8_r)
flag_person_wide_flag8$flag8_r <- gsub(" NA", replacement="", flag_person_wide_flag8$flag8_r)

#產生檢誤報告文字
flag8_temp <- flag_person_wide_flag8 %>%
  group_by(organization_id) %>%
  mutate(flag8_txt = paste(source, "需修改國籍別處：", flag8_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag8_txt)) %>%
  distinct(organization_id, flag8_txt)

#根據organization_id，展開成寬資料(wide)
flag8 <- flag8_temp %>%
  dcast(organization_id ~ flag8_txt, value.var = "flag8_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag8)[2 : length(colnames(flag8))]
flag8$flag8 <- NA
for (i in temp){
  flag8$flag8 <- paste(flag8$flag8, flag8[[i]], sep = "； ")
}
flag8$flag8 <- gsub("NA； ", replacement="", flag8$flag8)
flag8$flag8 <- gsub("； NA", replacement="", flag8$flag8)

#產生檢誤報告文字
flag8 <- flag8 %>%
  subset(select = c(organization_id, flag8)) %>%
  distinct(organization_id, flag8)
}else{
#偵測flag8是否存在。若不存在，則產生NA行
if('flag8' %in% ls()){
  print("flag8")
}else{
  flag8 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag8$flag8 <- ""
}
}
# flag9: 博士、碩士、學士、副學士畢業學校國別（一）～（二）應填入「本國」或者外交部網站之世界各國名稱一覽表的國家名稱（或者至少須足以辨識國家）。 -------------------------------------------------------------------
  #flag9_append-------------------------------------------------------------------
flag_person <- drev_person_1

#檢視畢業學校國別欄位字元數不為3
view_flag9 <- distinct(flag_person, ddegreen1, ddegreen2, mdegreen1, mdegreen2, bdegreen1, bdegreen2, adegreen1, adegreen2, .keep_all = TRUE) %>%
  subset(nchar(ddegreen1) != 3 | nchar(ddegreen2) != 3 | nchar(mdegreen1) != 3 | nchar(mdegreen2) != 3 | nchar(bdegreen1) != 3 | nchar(bdegreen2) != 3 | nchar(adegreen1) != 3 | nchar(adegreen2) != 3) %>%
  subset(select = c(organization_id, idnumber, ddegreen1, ddegreen2, mdegreen1, mdegreen2, bdegreen1, bdegreen2, adegreen1, adegreen2, edu_name2, source))

#"本國美國"標記為1
flag_person$err_flag <- case_when(
  flag_person$ddegreen1 == "本國美國" | flag_person$ddegreen2 == "本國美國" | flag_person$mdegreen1 == "本國美國" | flag_person$mdegreen2 == "本國美國" | flag_person$bdegreen1 == "本國美國" | flag_person$bdegreen2 == "本國美國" | flag_person$adegreen1 == "本國美國" | flag_person$adegreen2 == "本國美國" ~ 1,
    TRUE ~ 0
)
  #---

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag9 <- flag_person %>%
  subset(select = c(organization_id, idnumber, name, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ name, value.var = "name")

#合併所有name
temp <- colnames(flag_person_wide_flag9)[3 : length(colnames(flag_person_wide_flag9))]
flag_person_wide_flag9$flag9_r <- NA
for (i in temp){
  flag_person_wide_flag9$flag9_r <- paste(flag_person_wide_flag9$flag9_r, flag_person_wide_flag9[[i]], sep = " ")
}
flag_person_wide_flag9$flag9_r <- gsub("NA ", replacement="", flag_person_wide_flag9$flag9_r)
flag_person_wide_flag9$flag9_r <- gsub(" NA", replacement="", flag_person_wide_flag9$flag9_r)

#產生檢誤報告文字
flag9_temp <- flag_person_wide_flag9 %>%
  group_by(organization_id) %>%
  mutate(flag9_txt = paste(source, "需修改畢業學校國別者：", flag9_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag9_txt)) %>%
  distinct(organization_id, flag9_txt)

#根據organization_id，展開成寬資料(wide)
flag9 <- flag9_temp %>%
  dcast(organization_id ~ flag9_txt, value.var = "flag9_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag9)[2 : length(colnames(flag9))]
flag9$flag9 <- NA
for (i in temp){
  flag9$flag9 <- paste(flag9$flag9, flag9[[i]], sep = "； ")
}
flag9$flag9 <- gsub("NA； ", replacement="", flag9$flag9)
flag9$flag9 <- gsub("； NA", replacement="", flag9$flag9)

#產生檢誤報告文字
flag9 <- flag9 %>%
  subset(select = c(organization_id, flag9)) %>%
  distinct(organization_id, flag9)
}else{
#偵測flag9是否存在。若不存在，則產生NA行
if('flag9' %in% ls()){
  print("flag9")
}else{
  flag9 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag9$flag9 <- ""
}
}
# flag15: 兼任行政職職稱（一）～（三）不應填入校長或非行政職職稱，例如老師、教師、運動教練等。 -------------------------------------------------------------------
  #flag15_append-------------------------------------------------------------------
flag_person <- drev_person_1

#"本國美國"標記為1
flag_person$err_flag_admintitle1 <- 0
flag_person$err_flag_admintitle2 <- 0
flag_person$err_flag_admintitle3 <- 0
flag_person$err_flag_admintitle1 <- case_when(
  grepl("教學支援工作人員$", flag_person$admintitle1) | grepl("教學支援人員$", flag_person$admintitle1) | grepl("老師$", flag_person$admintitle1) | grepl("教師$", flag_person$admintitle1) | grepl("導師", flag_person$admintitle1) | grepl("運動教練", flag_person$admintitle1) | grepl("^校長$", flag_person$admintitle1) | grepl("教官$", flag_person$admintitle1) ~ 1,
  TRUE ~ flag_person$err_flag_admintitle1
)
flag_person$err_flag_admintitle2 <- case_when(
  grepl("教學支援工作人員$", flag_person$admintitle2) | grepl("教學支援人員$", flag_person$admintitle2) | grepl("老師$", flag_person$admintitle2) | grepl("教師$", flag_person$admintitle2) | grepl("導師", flag_person$admintitle2) | grepl("運動教練", flag_person$admintitle2) | grepl("^校長$", flag_person$admintitle2) | grepl("教官$", flag_person$admintitle2) ~ 1,
  TRUE ~ flag_person$err_flag_admintitle2
)
flag_person$err_flag_admintitle3 <- case_when(
  grepl("教學支援工作人員$", flag_person$admintitle3) | grepl("教學支援人員$", flag_person$admintitle3) | grepl("老師$", flag_person$admintitle3) | grepl("教師$", flag_person$admintitle3) | grepl("導師", flag_person$admintitle3) | grepl("運動教練", flag_person$admintitle3) | grepl("^校長$", flag_person$admintitle3) | grepl("教官$", flag_person$admintitle3) ~ 1,
  TRUE ~ flag_person$err_flag_admintitle3
)

flag_person$err_flag <- flag_person$err_flag_admintitle1 + flag_person$err_flag_admintitle2 + flag_person$err_flag_admintitle3

#加註職稱
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag_admintitle1 == 1 ~ paste(flag_person$name, "（", flag_person$admintitle1, "）", sep = ""),
  TRUE ~ flag_person$err_flag_txt
)
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag_admintitle2 == 1 ~ paste(flag_person$name, "（", flag_person$admintitle2, "）", sep = ""),
  TRUE ~ flag_person$err_flag_txt
)
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag_admintitle3 == 1 ~ paste(flag_person$name, "（", flag_person$admintitle3, "）", sep = ""),
  TRUE ~ flag_person$err_flag_txt
)
  #---

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag15 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag15)[3 : length(colnames(flag_person_wide_flag15))]
flag_person_wide_flag15$flag15_r <- NA
for (i in temp){
  flag_person_wide_flag15$flag15_r <- paste(flag_person_wide_flag15$flag15_r, flag_person_wide_flag15[[i]], sep = " ")
}
flag_person_wide_flag15$flag15_r <- gsub("NA ", replacement="", flag_person_wide_flag15$flag15_r)
flag_person_wide_flag15$flag15_r <- gsub(" NA", replacement="", flag_person_wide_flag15$flag15_r)

#產生檢誤報告文字
flag15_temp <- flag_person_wide_flag15 %>%
  group_by(organization_id) %>%
  mutate(flag15_txt = paste(source, "需修改兼任行政職職稱：", flag15_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag15_txt)) %>%
  distinct(organization_id, flag15_txt)

#根據organization_id，展開成寬資料(wide)
flag15 <- flag15_temp %>%
  dcast(organization_id ~ flag15_txt, value.var = "flag15_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag15)[2 : length(colnames(flag15))]
flag15$flag15 <- NA
for (i in temp){
  flag15$flag15 <- paste(flag15$flag15, flag15[[i]], sep = "； ")
}
flag15$flag15 <- gsub("NA； ", replacement="", flag15$flag15)
flag15$flag15 <- gsub("； NA", replacement="", flag15$flag15)

#產生檢誤報告文字
flag15 <- flag15 %>%
  subset(select = c(organization_id, flag15)) %>%
  distinct(organization_id, flag15) %>%
  mutate(flag15 = paste(flag15, "（校長、教師、教官、主任教官、族語老師、教學支援人員屬於服務身分別，若渠等教員未再兼任行政職務，如秘書、學務主任、生活輔導組組長等，請於兼任行政職職稱(單位)填“N” ）", sep = ""))
}else{
#偵測flag15是否存在。若不存在，則產生NA行
if('flag15' %in% ls()){
  print("flag15")
}else{
  flag15 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag15$flag15 <- ""
}
}
# flag16: 請假類別應依《教師請假規則》、《公務人員請假規則》以及人事行政主管機關公教員工請假給假一覽表相關規定填列。 -------------------------------------------------------------------
flag_person <- drev_person_1
  
#標記各種假別為1
flag_person$err_flag <- 1
flag_person$err_flag <- if_else(flag_person$leave == "事假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "家庭照顧假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "病假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "延長病假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "普通傷病假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "生理假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "婚假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "娩假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "分娩假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "產前假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "陪產假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "流產假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "安胎假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "喪假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "休假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "捐贈器官或骨髓假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "公假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "特別休假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "特休", 0, flag_person$err_flag)
#衛理女中確實核給"延長重病假"
flag_person$err_flag <- if_else(flag_person$leave == "延長重病假" & flag_person$organization_id == "411302", 0, flag_person$err_flag)

flag_person$err_flag <- if_else(flag_person$leave == "延長病假(安胎)", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "公假(公傷假)", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "公(傷)假", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$leave == "安胎病假、產前假、娩假", 0, flag_person$err_flag)

flag_person$err_flag <- if_else(grepl("留職停薪$", flag_person$leave) | grepl("留停$", flag_person$leave), 0, flag_person$err_flag)

flag_person$err_flag <- if_else(grepl("^育嬰假$", flag_person$leave) & flag_person$levpay == "育嬰留職停薪", 0, flag_person$err_flag)

flag_person$err_flag <- if_else(flag_person$leave == "N", 0, flag_person$err_flag)


#加註請假類別
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ paste(flag_person$name, "（", flag_person$leave, "）", sep = ""),
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag16 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")
  
#合併所有name
temp <- colnames(flag_person_wide_flag16)[3 : length(colnames(flag_person_wide_flag16))]
flag_person_wide_flag16$flag16_r <- NA
for (i in temp){
  flag_person_wide_flag16$flag16_r <- paste(flag_person_wide_flag16$flag16_r, flag_person_wide_flag16[[i]], sep = " ")
}
flag_person_wide_flag16$flag16_r <- gsub("NA ", replacement="", flag_person_wide_flag16$flag16_r)
flag_person_wide_flag16$flag16_r <- gsub(" NA", replacement="", flag_person_wide_flag16$flag16_r)
  
#產生檢誤報告文字
flag16_temp <- flag_person_wide_flag16 %>%
  group_by(organization_id) %>%
  mutate(flag16_txt = paste(source, "需修改請假類別：", flag16_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag16_txt)) %>%
  distinct(organization_id, flag16_txt)
  
#根據organization_id，展開成寬資料(wide)
flag16 <- flag16_temp %>%
  dcast(organization_id ~ flag16_txt, value.var = "flag16_txt")
  
#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag16)[2 : length(colnames(flag16))]
flag16$flag16 <- NA
for (i in temp){
  flag16$flag16 <- paste(flag16$flag16, flag16[[i]], sep = "； ")
}
flag16$flag16 <- gsub("NA； ", replacement="", flag16$flag16)
flag16$flag16 <- gsub("； NA", replacement="", flag16$flag16)
  
#產生檢誤報告文字
flag16 <- flag16 %>%
  subset(select = c(organization_id, flag16)) %>%
  distinct(organization_id, flag16) %>%
  mutate(flag16 = paste(flag16, "（請確認請假類別，或是否屬於請假，若非屬請假請填寫 “N”）", sep = ""))
}else{
#偵測flag16是否存在。若不存在，則產生NA行
if('flag16' %in% ls()){
  print("flag16")
}else{
  flag16 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag16$flag16 <- ""
}
}
# flag18: 人事資料表各欄位是否有資料分布異常的情形。 -------------------------------------------------------------------
flag_person <- drev_person_1

flag_person$count_emptype <- if_else(flag_person$emptype == "專任", 1, 0)
flag_person$count_empunit <- if_else(flag_person$empunit == "高中部日間部" | flag_person$empunit == "國中部日間部" | flag_person$empunit == "中學部", 1, 0)
flag_person$count_sertype <- if_else(flag_person$sertype == "教師", 1, 0)
flag_person$count_sertype2 <- if_else(flag_person$sertype == "校長", 1, 0)
flag_person$count_skillteacher <- if_else(flag_person$skillteacher == "N", 1, 0)
flag_person$count_counselor <- if_else(flag_person$counselor == "N", 1, 0)
flag_person$count_speteacher <- if_else(flag_person$speteacher == "N", 1, 0)
flag_person$count_joiteacher <- if_else(flag_person$joiteacher == "N", 1, 0)
flag_person$count_expecter <- if_else(flag_person$expecter == "N", 1, 0)
flag_person$count_workexp <- if_else(flag_person$workexp == "N", 1, 0)
flag_person$count_study <- if_else(flag_person$study == "N", 1, 0)

#調整NA
temp <- c("count_emptype", "count_empunit", "count_sertype", "count_sertype2", "count_skillteacher", "count_counselor", "count_speteacher", "count_counselor", "count_speteacher", "count_joiteacher", "count_expecter", "count_workexp", "count_study")
for (x in temp){
  flag_person[[x]][is.na(flag_person[[x]])] <- 0
}

flag_person$jj <- 1

flag_person_wide_flag18 <- aggregate(cbind(count_emptype, count_empunit, count_sertype, count_sertype2, count_skillteacher, count_counselor, count_speteacher, count_joiteacher, count_expecter, count_workexp, count_study, jj) ~ organization_id + source, flag_person, sum)

flag_person_wide_flag18$flag_err <- 0
flag_person_wide_flag18$err_emptype <- if_else(flag_person_wide_flag18$count_emptype / flag_person_wide_flag18$jj < 0.5 & flag_person_wide_flag18$source == "教員資料表", "教員資料表專任教學人員人數偏低，請再協助確認實際聘任情況，或請確認是否填報完整教員名單資料。", "")
flag_person_wide_flag18$err_emptype <- if_else(flag_person_wide_flag18$count_emptype / flag_person_wide_flag18$jj < 0.5 & flag_person_wide_flag18$source == "職員工資料表", "職員(工)資料表專任人員人數偏低，請再協助確認實際聘任情況，或請確認是否填報完整職員(工)名單資料。", "")
flag_person_wide_flag18$err_empunit <- if_else(flag_person_wide_flag18$count_empunit / flag_person_wide_flag18$jj < 0.5 & flag_person_wide_flag18$source == "教員資料表", "教員資料表主聘單位各類別人數分布異常，請再協助確認實際聘任情況。", "")
flag_person_wide_flag18$err_empunit <- if_else(flag_person_wide_flag18$count_empunit / flag_person_wide_flag18$jj < 0.5 & flag_person_wide_flag18$source == "職員工資料表", "職員(工)資料表主聘單位各類別人數分布異常，請再協助確認實際聘任情況。", "")
flag_person_wide_flag18$err_sertype <- if_else(flag_person_wide_flag18$count_sertype / flag_person_wide_flag18$jj < 0.5 & flag_person_wide_flag18$source == "教員資料表", "教師人數偏低，請再協助確認實際聘任情況。", "")
flag_person_wide_flag18$err_sertype <- if_else(flag_person_wide_flag18$count_sertype2 > 1 & flag_person_wide_flag18$source == "教員資料表", "校長人數超過一位，請再協助確認實際聘任情況。", "")
flag_person_wide_flag18$err_skillteacher <- if_else(flag_person_wide_flag18$count_skillteacher / flag_person_wide_flag18$jj < 0.5 & flag_person_wide_flag18$source == "教員資料表", "專業及技術教師人數偏多，請再協助確認實際聘任情況。", "")
flag_person_wide_flag18$err_counselor <- if_else(flag_person_wide_flag18$count_counselor / flag_person_wide_flag18$jj < 0.5 & flag_person_wide_flag18$source == "教員資料表", "專任輔導教師人數偏多，請再協助確認實際聘任情況。", "")
flag_person_wide_flag18$err_speteacher <- if_else(flag_person_wide_flag18$count_speteacher / flag_person_wide_flag18$jj < 0.5 & flag_person_wide_flag18$source == "教員資料表", "特教班專職教師人數偏多，請再協助確認實際聘任情況。", "")
flag_person_wide_flag18$err_joiteacher <- if_else(flag_person_wide_flag18$count_joiteacher / flag_person_wide_flag18$jj < 0.5 & flag_person_wide_flag18$source == "教員資料表", "合聘教師人數偏多。（請確認校內教師是否與他校合聘：如有與他校合聘者，本校又為『主聘學校』，再請於『是否為合聘教師』一欄填入『1』，若以本校為『從聘學校』請於『是否為合聘教師』一欄填入『2』；若沒有與他校合聘，則『是否為合聘教師』一欄請填『N』）", "")
flag_person_wide_flag18$err_expecter <- if_else(flag_person_wide_flag18$count_expecter / flag_person_wide_flag18$jj < 0.5 & flag_person_wide_flag18$source == "教員資料表", "業界專家人數偏多，請再協助確認實際聘任情況，或請確認是否將專業及技術教師誤填為業界專家。", "")
flag_person_wide_flag18$err_workexp <- if_else(flag_person_wide_flag18$count_workexp / flag_person_wide_flag18$jj < 0.5 & flag_person_wide_flag18$source == "教員資料表", "一年以上與任教領域相關之業界實務工作經驗人數偏多。（請再協助確認，『是否具備一年以上與任教領域相關之業界實務工作經驗』填寫『Y』之教員，是否確依欄位說明具備此經驗）", "")
flag_person_wide_flag18$err_study <- if_else(flag_person_wide_flag18$count_study / flag_person_wide_flag18$jj < 0.5 & flag_person_wide_flag18$source == "教員資料表", "近六年內進行與專業或技術有關之研習或研究的人數偏多，請再協助確認實際聘任情況。", "")

flag_person_wide_flag18$err_flag_txt <- paste(flag_person_wide_flag18$err_emptype, flag_person_wide_flag18$err_empunit, flag_person_wide_flag18$err_sertype, flag_person_wide_flag18$err_skillteacher, flag_person_wide_flag18$err_counselor, flag_person_wide_flag18$err_speteacher, flag_person_wide_flag18$err_joiteacher, flag_person_wide_flag18$err_expecter, flag_person_wide_flag18$err_workexp, flag_person_wide_flag18$err_study, sep = " ")

# #產生檢誤報告文字
# flag18_temp <- flag_person_wide_flag18 %>%
#   group_by(organization_id) %>%
#   mutate(flag18_txt = paste(source, "需修改請假類別：", flag18_r, sep = ""), "") %>%
#   subset(select = c(organization_id, flag18_txt)) %>%
#   distinct(organization_id, flag18_txt)

if (dim(flag_person %>% subset(grepl("\\S", flag_person_wide_flag18$err_flag_txt)))[1] != 0){
#根據organization_id，展開成寬資料(wide)
flag18 <- flag_person_wide_flag18 %>%
  subset(grepl("\\S", flag_person_wide_flag18$err_flag_txt)) %>%
  dcast(organization_id ~ err_flag_txt, value.var = "err_flag_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag18)[2 : length(colnames(flag18))]
flag18$flag18 <- NA
for (i in temp){
  flag18$flag18 <- paste(flag18$flag18, flag18[[i]], sep = "； ")
}
flag18$flag18 <- gsub("NA； ", replacement="", flag18$flag18)
flag18$flag18 <- gsub("； NA", replacement="", flag18$flag18)

#產生檢誤報告文字
flag18 <- flag18 %>%
  subset(select = c(organization_id, flag18)) %>%
  distinct(organization_id, flag18)

#刪除字串最後異常空格
trim_t <- function (x){
  gsub("\\s+|\\s+$", "", x)
}

flag18$flag18 <- trim_t(flag18$flag18) ##test
}else{
#偵測flag18是否存在。若不存在，則產生NA行
if('flag18' %in% ls()){
  print("flag18")
}else{
  flag18 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag18$flag18 <- ""
}
}
# flag19: 填寫外來人口統一證號者，國籍別應非「本國籍」。 -------------------------------------------------------------------
flag_person <- drev_person_1

#外來人口統一證號：第二碼為A B C D 8 9
flag_person$err_flag <- 0
flag_person$err_flag <- if_else(substr(flag_person$idnumber, 2, 2) == "A" & (flag_person$nation == "本國籍" | flag_person$nation == "本國"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(substr(flag_person$idnumber, 2, 2) == "B" & (flag_person$nation == "本國籍" | flag_person$nation == "本國"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(substr(flag_person$idnumber, 2, 2) == "C" & (flag_person$nation == "本國籍" | flag_person$nation == "本國"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(substr(flag_person$idnumber, 2, 2) == "D" & (flag_person$nation == "本國籍" | flag_person$nation == "本國"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(substr(flag_person$idnumber, 2, 2) == "8" & (flag_person$nation == "本國籍" | flag_person$nation == "本國"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(substr(flag_person$idnumber, 2, 2) == "9" & (flag_person$nation == "本國籍" | flag_person$nation == "本國"), 1, flag_person$err_flag)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag19 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag19)[3 : length(colnames(flag_person_wide_flag19))]
flag_person_wide_flag19$flag19_r <- NA
for (i in temp){
  flag_person_wide_flag19$flag19_r <- paste(flag_person_wide_flag19$flag19_r, flag_person_wide_flag19[[i]], sep = " ")
}
flag_person_wide_flag19$flag19_r <- gsub("NA ", replacement="", flag_person_wide_flag19$flag19_r)
flag_person_wide_flag19$flag19_r <- gsub(" NA", replacement="", flag_person_wide_flag19$flag19_r)

#產生檢誤報告文字
flag19_temp <- flag_person_wide_flag19 %>%
  group_by(organization_id) %>%
  mutate(flag19_txt = paste(source, "：", flag19_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag19_txt)) %>%
  distinct(organization_id, flag19_txt)

#根據organization_id，展開成寬資料(wide)
flag19 <- flag19_temp %>%
  dcast(organization_id ~ flag19_txt, value.var = "flag19_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag19)[2 : length(colnames(flag19))]
flag19$flag19 <- NA
for (i in temp){
  flag19$flag19 <- paste(flag19$flag19, flag19[[i]], sep = "； ")
}
flag19$flag19 <- gsub("NA； ", replacement="", flag19$flag19)
flag19$flag19 <- gsub("； NA", replacement="", flag19$flag19)

#產生檢誤報告文字
flag19 <- flag19 %>%
  subset(select = c(organization_id, flag19)) %>%
  distinct(organization_id, flag19) %>%
  mutate(flag19 = paste(flag19, "（請修正該員所屬國籍別）", sep = ""))
}else{
#偵測flag19是否存在。若不存在，則產生NA行
if('flag19' %in% ls()){
  print("flag19")
}else{
  flag19 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag19$flag19 <- ""
}
}
# flag20: 教職員工畢業學校若為專科學校，學歷資訊應於「副學士」畢業學校欄位填列。 -------------------------------------------------------------------
flag_person <- drev_person_1

#學士學位畢業學校名稱不可出現"專科學校"
flag_person$err_flag_bdegreeu1 <- 0
flag_person$err_flag_bdegreeu1 <- if_else(grepl("專科", flag_person$bdegreeu1), 1, flag_person$err_flag_bdegreeu1)
flag_person$err_flag_bdegreeu1 <- if_else(grepl("二專", flag_person$bdegreeu1), 1, flag_person$err_flag_bdegreeu1)
flag_person$err_flag_bdegreeu1 <- if_else(grepl("五專", flag_person$bdegreeu1), 1, flag_person$err_flag_bdegreeu1)
flag_person$err_flag_bdegreeu1 <- if_else(grepl("海專", flag_person$bdegreeu1), 1, flag_person$err_flag_bdegreeu1)
flag_person$err_flag_bdegreeu1 <- if_else(grepl("工專", flag_person$bdegreeu1), 1, flag_person$err_flag_bdegreeu1)
flag_person$err_flag_bdegreeu1 <- if_else(grepl("商專", flag_person$bdegreeu1), 1, flag_person$err_flag_bdegreeu1)
flag_person$err_flag_bdegreeu1 <- if_else(grepl("藝專", flag_person$bdegreeu1), 1, flag_person$err_flag_bdegreeu1)
flag_person$err_flag_bdegreeu1 <- if_else(grepl("農專", flag_person$bdegreeu1), 1, flag_person$err_flag_bdegreeu1)
flag_person$err_flag_bdegreeu1 <- if_else(grepl("護專", flag_person$bdegreeu1), 1, flag_person$err_flag_bdegreeu1)
flag_person$err_flag_bdegreeu1 <- if_else(grepl("家專", flag_person$bdegreeu1), 1, flag_person$err_flag_bdegreeu1)
flag_person$err_flag_bdegreeu1 <- if_else(grepl("行專", flag_person$bdegreeu1), 1, flag_person$err_flag_bdegreeu1)
flag_person$err_flag_bdegreeu1 <- if_else(grepl("師專", flag_person$bdegreeu1), 1, flag_person$err_flag_bdegreeu1)
flag_person$err_flag_bdegreeu1 <- if_else(grepl("藥專", flag_person$bdegreeu1), 1, flag_person$err_flag_bdegreeu1)
flag_person$err_flag_bdegreeu1 <- if_else(grepl("^台南家專學校財團法人台南應用科技大學$", flag_person$bdegreeu1), 0, flag_person$err_flag_bdegreeu1)
#陸軍官校專科班為學士學位
flag_person$err_flag_bdegreeu1 <- if_else(grepl("^陸軍官校專科班$", flag_person$bdegreeu1), 0, flag_person$err_flag_bdegreeu1)

flag_person$err_flag_bdegreeu2 <- 0
flag_person$err_flag_bdegreeu2 <- if_else(grepl("專科", flag_person$bdegreeu2), 1, flag_person$err_flag_bdegreeu2)
flag_person$err_flag_bdegreeu2 <- if_else(grepl("二專", flag_person$bdegreeu2), 1, flag_person$err_flag_bdegreeu2)
flag_person$err_flag_bdegreeu2 <- if_else(grepl("五專", flag_person$bdegreeu2), 1, flag_person$err_flag_bdegreeu2)
flag_person$err_flag_bdegreeu2 <- if_else(grepl("海專", flag_person$bdegreeu2), 1, flag_person$err_flag_bdegreeu2)
flag_person$err_flag_bdegreeu2 <- if_else(grepl("工專", flag_person$bdegreeu2), 1, flag_person$err_flag_bdegreeu2)
flag_person$err_flag_bdegreeu2 <- if_else(grepl("商專", flag_person$bdegreeu2), 1, flag_person$err_flag_bdegreeu2)
flag_person$err_flag_bdegreeu2 <- if_else(grepl("藝專", flag_person$bdegreeu2), 1, flag_person$err_flag_bdegreeu2)
flag_person$err_flag_bdegreeu2 <- if_else(grepl("農專", flag_person$bdegreeu2), 1, flag_person$err_flag_bdegreeu2)
flag_person$err_flag_bdegreeu2 <- if_else(grepl("護專", flag_person$bdegreeu2), 1, flag_person$err_flag_bdegreeu2)
flag_person$err_flag_bdegreeu2 <- if_else(grepl("家專", flag_person$bdegreeu2), 1, flag_person$err_flag_bdegreeu2)
flag_person$err_flag_bdegreeu2 <- if_else(grepl("行專", flag_person$bdegreeu2), 1, flag_person$err_flag_bdegreeu2)
flag_person$err_flag_bdegreeu2 <- if_else(grepl("師專", flag_person$bdegreeu2), 1, flag_person$err_flag_bdegreeu2)
flag_person$err_flag_bdegreeu2 <- if_else(grepl("藥專", flag_person$bdegreeu2), 1, flag_person$err_flag_bdegreeu2)
flag_person$err_flag_bdegreeu2 <- if_else(grepl("^台南家專學校財團法人台南應用科技大學$", flag_person$bdegreeu2), 0, flag_person$err_flag_bdegreeu2)
#陸軍官校專科班為學士學位
flag_person$err_flag_bdegreeu2 <- if_else(grepl("^陸軍官校專科班$", flag_person$bdegreeu2), 0, flag_person$err_flag_bdegreeu2)

flag_person$err_flag <- flag_person$err_flag_bdegreeu1 + flag_person$err_flag_bdegreeu2

#加註學士學位畢業學校名稱
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag_bdegreeu1 == 1 ~ paste(flag_person$name, "（學士學位畢業學校（一）：", flag_person$bdegreeu1, "）", sep = ""),
  flag_person$err_flag_bdegreeu2 == 1 ~ paste(flag_person$name, "（學士學位畢業學校（二）：", flag_person$bdegreeu2, "）", sep = ""),
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag20 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag20)[3 : length(colnames(flag_person_wide_flag20))]
flag_person_wide_flag20$flag20_r <- NA
for (i in temp){
  flag_person_wide_flag20$flag20_r <- paste(flag_person_wide_flag20$flag20_r, flag_person_wide_flag20[[i]], sep = " ")
}
flag_person_wide_flag20$flag20_r <- gsub("NA ", replacement="", flag_person_wide_flag20$flag20_r)
flag_person_wide_flag20$flag20_r <- gsub(" NA", replacement="", flag_person_wide_flag20$flag20_r)

#產生檢誤報告文字
flag20_temp <- flag_person_wide_flag20 %>%
  group_by(organization_id) %>%
  mutate(flag20_txt = paste(source, "：", flag20_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag20_txt)) %>%
  distinct(organization_id, flag20_txt)

#根據organization_id，展開成寬資料(wide)
flag20 <- flag20_temp %>%
  dcast(organization_id ~ flag20_txt, value.var = "flag20_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag20)[2 : length(colnames(flag20))]
flag20$flag20 <- NA
for (i in temp){
  flag20$flag20 <- paste(flag20$flag20, flag20[[i]], sep = "； ")
}
flag20$flag20 <- gsub("NA； ", replacement="", flag20$flag20)
flag20$flag20 <- gsub("； NA", replacement="", flag20$flag20)

#產生檢誤報告文字
flag20 <- flag20 %>%
  subset(select = c(organization_id, flag20)) %>%
  distinct(organization_id, flag20)
}else{
#偵測flag20是否存在。若不存在，則產生NA行
if('flag20' %in% ls()){
  print("flag20")
}else{
  flag20 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag20$flag20 <- ""
}
}
# flag24: 本校到職日期與填報基準日的差距，不應小於本校任職需扣除年資。 -------------------------------------------------------------------
flag_person <- drev_person_1

flag_person$survey_year <- 2022
flag_person$survey_mon <- 9
flag_person$onbodaty <- ""
flag_person$onbodatm <- ""
flag_person$onbodatd <- ""

flag_person$onbodaty <- if_else(nchar(flag_person$onbodat) == 6, substr(flag_person$onbodat, 1, 2), flag_person$onbodaty)
flag_person$onbodatm <- if_else(nchar(flag_person$onbodat) == 6, substr(flag_person$onbodat, 3, 4), flag_person$onbodatm)
flag_person$onbodatd <- if_else(nchar(flag_person$onbodat) == 6, substr(flag_person$onbodat, 5, 6), flag_person$onbodatd)
flag_person$onbodaty <- if_else(nchar(flag_person$onbodat) == 7, substr(flag_person$onbodat, 1, 3), flag_person$onbodaty)
flag_person$onbodatm <- if_else(nchar(flag_person$onbodat) == 7, substr(flag_person$onbodat, 4, 5), flag_person$onbodatm)
flag_person$onbodatd <- if_else(nchar(flag_person$onbodat) == 7, substr(flag_person$onbodat, 6, 7), flag_person$onbodatd)

flag_person$onbodaty <- as.numeric(flag_person$onbodaty)
flag_person$onbodatm <- as.numeric(flag_person$onbodatm)
flag_person$onbodatd <- as.numeric(flag_person$onbodatd)

#本校服務年資
flag_person$tser <- 0
flag_person$tser <- if_else(flag_person$survey_year %% 4 != 0, ((flag_person$survey_year-1911) + 9/12 + 30/365) - (flag_person$onbodaty + (flag_person$onbodatm/12) + (flag_person$onbodatd/365)), flag_person$tser)
flag_person$tser <- if_else(flag_person$survey_year %% 4 == 0, ((flag_person$survey_year-1911) + 9/12 + 30/366) - (flag_person$onbodaty + (flag_person$onbodatm/12) + (flag_person$onbodatd/366)), flag_person$tser)

#本次本校任職需扣除之年資
flag_person$desey <- substr(flag_person$desedym, 1, 2) %>% as.numeric()
flag_person$desem <- substr(flag_person$desedym, 3, 4) %>% as.numeric()

flag_person$dese <- (flag_person$desey + (flag_person$desem / 12))

#本校服務年資-本校任職需扣除之年資 才是實際在本校的服務年資
flag_person$tser <- flag_person$tser - flag_person$dese

#本校到職前學校服務總年資
flag_person$beoby <- substr(flag_person$beobdym, 1, 2) %>% as.numeric
flag_person$beobm <- substr(flag_person$beobdym, 3, 4) %>% as.numeric

flag_person$beob <- (flag_person$beoby + (flag_person$beobm / 12))

#學校教學工作總年資
flag_person$tsch <- flag_person$tser + flag_person$beob

#tser要小於-0.00137而不是0的原因：本校到職日期+本次本校任職需扣除之年資可能為4月1日，剛好超過資料基準日3月31日一天
#tser改成要小於-0.0041而不是-0.00137的原因：本校到職日期+本次本校任職需扣除之年資可能為10月1日，剛好超過資料基準日9月30日一天
#因扣除年資未滿一個月以一個月計，下學期基準日為2/28，可能造成誤差

flag_person$err_flag <- 0
flag_person$err_flag <- if_else(flag_person$tser < -.0806, 1, flag_person$err_flag)

#若spe3錯 則不應出現在flag24
flag_person$survey_year <- flag_person$survey_year - 1911

flag_person$arvy <- substr(flag_person$onbodat, 1, 3) %>% as.numeric()
flag_person$arvm <- substr(flag_person$onbodat, 4, 5) %>% as.numeric()

flag_person$err_spe = if_else((flag_person$arvy * 12 + flag_person$arvm) > (flag_person$survey_year * 12 + flag_person$survey_mon), 1, 0)

flag_person$err_flag <- if_else(flag_person$err_spe == 1, 0, flag_person$err_flag)

#換算到職至資料基準日的日期
flag_person$tser_ndese <- flag_person$tser + flag_person$dese
flag_person$tser_ndesey <- floor(flag_person$tser_ndese)
flag_person$tser_ndesem <- ceiling((flag_person$tser_ndese - floor(flag_person$tser_ndese)) * 12)
flag_person$ndesey <- floor(flag_person$dese)
flag_person$ndesem <- ceiling((flag_person$dese - floor(flag_person$dese)) * 12)

temp <- c("tser_ndesey", "tser_ndesem", "ndesey", "ndesem")
for (x in temp) {
  flag_person[[x]] <- flag_person[[x]] %>% as.character()
}

#加註到職至資料基準日的時間，和扣除年資
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ paste(flag_person$name, flag_person$onbodat,"到職（到職至資料基準日為", flag_person$tser_ndesey, "年", flag_person$tser_ndesem, "個月，但扣除年資為", flag_person$ndesey, "年", flag_person$ndesem, "個月", "）", sep = ""),
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag24 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag24)[3 : length(colnames(flag_person_wide_flag24))]
flag_person_wide_flag24$flag24_r <- NA
for (i in temp){
  flag_person_wide_flag24$flag24_r <- paste(flag_person_wide_flag24$flag24_r, flag_person_wide_flag24[[i]], sep = " ")
}
flag_person_wide_flag24$flag24_r <- gsub("NA ", replacement="", flag_person_wide_flag24$flag24_r)
flag_person_wide_flag24$flag24_r <- gsub(" NA", replacement="", flag_person_wide_flag24$flag24_r)

#產生檢誤報告文字
flag24_temp <- flag_person_wide_flag24 %>%
  group_by(organization_id) %>%
  mutate(flag24_txt = paste(source, "：", flag24_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag24_txt)) %>%
  distinct(organization_id, flag24_txt)

#根據organization_id，展開成寬資料(wide)
flag24 <- flag24_temp %>%
  dcast(organization_id ~ flag24_txt, value.var = "flag24_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag24)[2 : length(colnames(flag24))]
flag24$flag24 <- NA
for (i in temp){
  flag24$flag24 <- paste(flag24$flag24, flag24[[i]], sep = "； ")
}
flag24$flag24 <- gsub("NA； ", replacement="", flag24$flag24)
flag24$flag24 <- gsub("； NA", replacement="", flag24$flag24)

#產生檢誤報告文字
flag24 <- flag24 %>%
  subset(select = c(organization_id, flag24)) %>%
  distinct(organization_id, flag24) %>%
  mutate(flag24 = paste("請確認該員之「本校到職日期」、「本校任職需扣除之年資」，", flag24, sep = ""))
}else{
#偵測flag24是否存在。若不存在，則產生NA行
if('flag24' %in% ls()){
  print("flag24")
}else{
  flag24 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag24$flag24 <- ""
}
}
# flag39: 學校工作總年資（本校服務年資+本校到職前學校服務總年資），與年齡之差距過大或過小。 -------------------------------------------------------------------
flag_person <- drev_person_1

#年齡
#創設變項出生年月日：birthy birthm birthd
flag_person$birthy <- ""
flag_person$birthm <- ""
flag_person$birthd <- ""

flag_person$birthy <- if_else(nchar(flag_person$birthdate) == 6, substr(flag_person$birthdate, 1, 2), flag_person$birthy)
flag_person$birthm <- if_else(nchar(flag_person$birthdate) == 6, substr(flag_person$birthdate, 3, 4), flag_person$birthm)
flag_person$birthd <- if_else(nchar(flag_person$birthdate) == 6, substr(flag_person$birthdate, 5, 6), flag_person$birthd)
flag_person$birthy <- if_else(nchar(flag_person$birthdate) == 7, substr(flag_person$birthdate, 1, 3), flag_person$birthy)
flag_person$birthm <- if_else(nchar(flag_person$birthdate) == 7, substr(flag_person$birthdate, 4, 5), flag_person$birthm)
flag_person$birthd <- if_else(nchar(flag_person$birthdate) == 7, substr(flag_person$birthdate, 6, 7), flag_person$birthd)

flag_person$birthy <- as.numeric(flag_person$birthy)
flag_person$birthm <- as.numeric(flag_person$birthm)
flag_person$birthd <- as.numeric(flag_person$birthd)

flag_person$survey_year <- 2022

#創設變項年齡（以年為單位）：age
flag_person$age <- 0
flag_person$age <- if_else(flag_person$survey_year %% 4 != 0, ((flag_person$survey_year-1911) + 9/12 + 30/365) - (flag_person$birthy + (flag_person$birthm/12) + (flag_person$birthd/365)), flag_person$age)
flag_person$age <- if_else(flag_person$survey_year %% 4 == 0, ((flag_person$survey_year-1911) + 9/12 + 30/366) - (flag_person$birthy + (flag_person$birthm/12) + (flag_person$birthd/366)), flag_person$age)

flag_person$survey_year <- 2022
flag_person$onbodaty <- ""
flag_person$onbodatm <- ""
flag_person$onbodatd <- ""
flag_person$onbodatd <- ""

flag_person$onbodaty <- if_else(nchar(flag_person$onbodat) == 6, substr(flag_person$onbodat, 1, 2), flag_person$onbodaty)
flag_person$onbodatm <- if_else(nchar(flag_person$onbodat) == 6, substr(flag_person$onbodat, 3, 4), flag_person$onbodatm)
flag_person$onbodatd <- if_else(nchar(flag_person$onbodat) == 6, substr(flag_person$onbodat, 5, 6), flag_person$onbodatd)
flag_person$onbodaty <- if_else(nchar(flag_person$onbodat) == 7, substr(flag_person$onbodat, 1, 3), flag_person$onbodaty)
flag_person$onbodatm <- if_else(nchar(flag_person$onbodat) == 7, substr(flag_person$onbodat, 4, 5), flag_person$onbodatm)
flag_person$onbodatd <- if_else(nchar(flag_person$onbodat) == 7, substr(flag_person$onbodat, 6, 7), flag_person$onbodatd)

flag_person$onbodaty <- as.numeric(flag_person$onbodaty)
flag_person$onbodatm <- as.numeric(flag_person$onbodatm)
flag_person$onbodatd <- as.numeric(flag_person$onbodatd)

#本校服務年資
flag_person$tser <- 0
flag_person$tser <- if_else(flag_person$survey_year %% 4 != 0, ((flag_person$survey_year-1911) + 9/12 + 30/365) - (flag_person$onbodaty + (flag_person$onbodatm/12) + (flag_person$onbodatd/365)), flag_person$tser)
flag_person$tser <- if_else(flag_person$survey_year %% 4 == 0, ((flag_person$survey_year-1911) + 9/12 + 30/366) - (flag_person$onbodaty + (flag_person$onbodatm/12) + (flag_person$onbodatd/366)), flag_person$tser)

#本次本校任職需扣除之年資
flag_person$desey <- substr(flag_person$desedym, 1, 2) %>% as.numeric()
flag_person$desem <- substr(flag_person$desedym, 3, 4) %>% as.numeric()

flag_person$dese <- (flag_person$desey + (flag_person$desem / 12))

#本校服務年資-本校任職需扣除資年資 才是實際在本校的服務年資
flag_person$tser <- flag_person$tser - flag_person$dese

#避免掉年資小於零的情況（因本校到職日期+本次本校任職需扣除之年資可能為8/1的情況）
flag_person$tser <- if_else(flag_person$tser < 0, 0, flag_person$tser)

#本校到職前學校服務總年資
flag_person$beoby <- substr(flag_person$beobdym, 1, 2) %>% as.numeric
flag_person$beobm <- substr(flag_person$beobdym, 3, 4) %>% as.numeric

flag_person$beob <- (flag_person$beoby + (flag_person$beobm / 12))

#學校教學工作總年資
flag_person$tsch <- flag_person$tser + flag_person$beob

flag_person$err_flag <- 0
flag_person$err_flag <- if_else(flag_person$age - flag_person$tsch <= 17, 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$age - flag_person$tsch > 75, 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$age - flag_person$tsch > 75 & (flag_person$emptype == "兼任" | flag_person$emptype == "長期代課" | flag_person$emptype == "專職族語老師" | flag_person$emptype == "鐘點教師" | flag_person$emptype == "約聘僱" | flag_person$emptype == "約用" | flag_person$emptype == "派遣"), 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$age - flag_person$tsch > 85 & (flag_person$emptype == "兼任" | flag_person$emptype == "長期代課" | flag_person$emptype == "專職族語老師" | flag_person$emptype == "鐘點教師" | flag_person$emptype == "約聘僱" | flag_person$emptype == "約用" | flag_person$emptype == "派遣"), 1, flag_person$err_flag)

# *確實為實際情況
# *市立瑞芳高工	013402	flag39	沈秋梅確實是17歲開始工作
# replace irr_seniority = 0 if organization_id == "013402" & idnumber == "F222463581"
# *國立蘭陽女中	020301	flag39	職員(工)資料表：莊惠娟59歲，但學校工作總年資有43年（約16歲開始工作）
# replace irr_seniority = 0 if organization_id == "020301" & idnumber == "G220091122"
# *國立羅東高中	020308	flag39	職員(工)資料表：朱文祥45歲，但學校工作總年資有30年（約15歲開始工作） 是屬實的
# replace irr_seniority = 0 if organization_id == "020308" & idnumber == "G121094054"
# *國立鹿港高中	070316	flag39	職員蔡奉育 52歲，但學校工作總年資有36年 (約16歲開始工作) 這是正確的
# replace irr_seniority = 0 if organization_id == "070316" & idnumber == "N220359428"
# *國立員林崇實高工	070409	flag39	王麗粧確實年資確實是38年。
# replace irr_seniority = 0 if organization_id == "070409" & idnumber == "N220792807"
# *國立員林家商	070410	flag39	黃美鳳  確實是15歲開始工作
# replace irr_seniority = 0 if organization_id == "070410" & idnumber == "N220798694"
# *國立中興高中	080305	flag39	李美紅57歲，但學校工作總年資有40年（約17歲開始工作） 黃鳳微63歲，但學校工作總年資有46年（約17歲開始工作） 是正確的 
# replace irr_seniority = 0 if organization_id == "080305" & idnumber == "N221676053"
# replace irr_seniority = 0 if organization_id == "080305" & idnumber == "M220125707"
# *國立新營高工	110406	flag39	職員方春寶　　確實１５歲任職，他是工友。
# replace irr_seniority = 0 if organization_id == "110406" & idnumber == "Q220737273"
# *國立臺南海事	210416	flag39	黃梅玉(16歲開始工作)、張秀惠(17歲開始工作)的確在該年紀就開始工作  故資料是正確的
# replace irr_seniority = 0 if organization_id == "210416" & idnumber == "D220464962"
# replace irr_seniority = 0 if organization_id == "210416" & idnumber == "D221156834"

flag_person$age <- floor(flag_person$age)
flag_person$tsch <- floor(flag_person$tsch)
flag_person$gowork <- flag_person$age - flag_person$tsch

temp <- c("age", "tsch", "gowork")
for (x in temp) {
  flag_person[[x]] <- flag_person[[x]] %>% as.character()
}

#加註學校工作總年資及工作起始歲數
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ paste(flag_person$name, flag_person$age, "歲，但學校工作總年資有", flag_person$tsch, "年（約", flag_person$gowork, "歲開始工作）", sep = ""),
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag39 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag39)[3 : length(colnames(flag_person_wide_flag39))]
flag_person_wide_flag39$flag39_r <- NA
for (i in temp){
  flag_person_wide_flag39$flag39_r <- paste(flag_person_wide_flag39$flag39_r, flag_person_wide_flag39[[i]], sep = " ")
}
flag_person_wide_flag39$flag39_r <- gsub("NA ", replacement="", flag_person_wide_flag39$flag39_r)
flag_person_wide_flag39$flag39_r <- gsub(" NA", replacement="", flag_person_wide_flag39$flag39_r)

#產生檢誤報告文字
flag39_temp <- flag_person_wide_flag39 %>%
  group_by(organization_id) %>%
  mutate(flag39_txt = paste(source, "：", flag39_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag39_txt)) %>%
  distinct(organization_id, flag39_txt)

#根據organization_id，展開成寬資料(wide)
flag39 <- flag39_temp %>%
  dcast(organization_id ~ flag39_txt, value.var = "flag39_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag39)[2 : length(colnames(flag39))]
flag39$flag39 <- NA
for (i in temp){
  flag39$flag39 <- paste(flag39$flag39, flag39[[i]], sep = "； ")
}
flag39$flag39 <- gsub("NA； ", replacement="", flag39$flag39)
flag39$flag39 <- gsub("； NA", replacement="", flag39$flag39)

#產生檢誤報告文字
flag39 <- flag39 %>%
  subset(select = c(organization_id, flag39)) %>%
  distinct(organization_id, flag39) %>%
  mutate(flag39 = paste("請確認該員之「本校到職日期」、「本校任職需扣除之年資」、「本校到職前學校服務總年資」，", flag39, sep = ""))
}else{
#偵測flag39是否存在。若不存在，則產生NA行
if('flag39' %in% ls()){
  print("flag39")
}else{
  flag39 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag39$flag39 <- ""
}
}
# flag45: 聘任科別應填入服務身分別為「教師」、「主任教官」、「教官」之聘任科別中文名稱。 -------------------------------------------------------------------
flag_person <- drev_person_1

#聘任科別不合理處
flag_person$err_flag <- 0
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & (flag_person$emsub == "NA" | flag_person$emsub == "N") & (flag_person$sertype == "教師" | flag_person$sertype == "主任教官" | flag_person$sertype == "教官"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & (flag_person$emsub == "不分科") & (flag_person$sertype == "主任教官" | flag_person$sertype == "教官"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & (flag_person$emsub == "教師") & (flag_person$sertype == "教師" | flag_person$sertype == "主任教官" | flag_person$sertype == "教官"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & (flag_person$emsub == "教官") & (flag_person$sertype == "教師" | flag_person$sertype == "主任教官" | flag_person$sertype == "教官"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & (flag_person$emsub == "主任教官") & (flag_person$sertype == "教師" | flag_person$sertype == "主任教官" | flag_person$sertype == "教官"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & (flag_person$emsub == "副校長") & (flag_person$sertype == "教師" | flag_person$sertype == "主任教官" | flag_person$sertype == "教官"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("室$", flag_person$emsub), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("處$", flag_person$emsub), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("^資處$", flag_person$emsub), 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("教官室", flag_person$emsub), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("教務處", flag_person$emsub), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("學務處", flag_person$emsub), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("人事室", flag_person$emsub), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("總務處", flag_person$emsub), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("會計室", flag_person$emsub), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("輔導室", flag_person$emsub), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("實習處", flag_person$emsub), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("圖書館", flag_person$emsub), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("校長室", flag_person$emsub), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("(全時支援他校)", flag_person$emsub), 0, flag_person$err_flag)

#社團 聘任類別為"鐘點教師"或"兼任"
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("^社團$", flag_person$emsub), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("社$", flag_person$emsub), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("^管樂$", flag_person$emsub) & (flag_person$emptype == "鐘點教師" | flag_person$emptype == "兼任"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("^合唱$", flag_person$emsub) & (flag_person$emptype == "鐘點教師" | flag_person$emptype == "兼任"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("^中正之家$", flag_person$emsub) & (flag_person$emptype == "鐘點教師" | flag_person$emptype == "兼任"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("^熱門音樂$", flag_person$emsub) & (flag_person$emptype == "鐘點教師" | flag_person$emptype == "兼任"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("^吉他$", flag_person$emsub) & (flag_person$emptype == "鐘點教師" | flag_person$emptype == "兼任"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("^魔術$", flag_person$emsub) & (flag_person$emptype == "鐘點教師" | flag_person$emptype == "兼任"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("^話劇$", flag_person$emsub) & (flag_person$emptype == "鐘點教師" | flag_person$emptype == "兼任"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("^國術$", flag_person$emsub) & (flag_person$emptype == "鐘點教師" | flag_person$emptype == "兼任"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("^劍道$", flag_person$emsub) & (flag_person$emptype == "鐘點教師" | flag_person$emptype == "兼任"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("^飛盤$", flag_person$emsub) & (flag_person$emptype == "鐘點教師" | flag_person$emptype == "兼任"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("^機器人研究$", flag_person$emsub) & (flag_person$emptype == "鐘點教師" | flag_person$emptype == "兼任"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("^儀隊$", flag_person$emsub) & (flag_person$emptype == "鐘點教師" | flag_person$emptype == "兼任"), 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & grepl("^滑板$", flag_person$emsub) & (flag_person$emptype == "鐘點教師" | flag_person$emptype == "兼任"), 1, flag_person$err_flag)
#「專門」指導學生「社團活動」之外聘指導教員，暫不納入填報。請依欄位說明，確認貴校教職員工名單是否正確。

#若校長的服務身份別填錯，且聘任科別填「NA」，則flag45不呈現，在flag47呈現
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & (flag_person$emsub == "NA" | flag_person$emsub == "N") & (flag_person$sertype == "校長"), 0, flag_person$err_flag)
#若聘任科別填校長，需抓出
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & (flag_person$emsub == "校長"), 1, flag_person$err_flag)

#加註聘任科別
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ paste(flag_person$name, "（", flag_person$emsub, "）", sep = ""),
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag45 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag45)[3 : length(colnames(flag_person_wide_flag45))]
flag_person_wide_flag45$flag45_r <- NA
for (i in temp){
  flag_person_wide_flag45$flag45_r <- paste(flag_person_wide_flag45$flag45_r, flag_person_wide_flag45[[i]], sep = " ")
}
flag_person_wide_flag45$flag45_r <- gsub("NA ", replacement="", flag_person_wide_flag45$flag45_r)
flag_person_wide_flag45$flag45_r <- gsub(" NA", replacement="", flag_person_wide_flag45$flag45_r)

#產生檢誤報告文字
flag45_temp <- flag_person_wide_flag45 %>%
  group_by(organization_id) %>%
  mutate(flag45_txt = paste(source, "需修改聘任科別(括號內為該員所對應之聘任科別欄位內容)：", flag45_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag45_txt)) %>%
  distinct(organization_id, flag45_txt)

#根據organization_id，展開成寬資料(wide)
flag45 <- flag45_temp %>%
  dcast(organization_id ~ flag45_txt, value.var = "flag45_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag45)[2 : length(colnames(flag45))]
flag45$flag45 <- NA
for (i in temp){
  flag45$flag45 <- paste(flag45$flag45, flag45[[i]], sep = "； ")
}
flag45$flag45 <- gsub("NA； ", replacement="", flag45$flag45)
flag45$flag45 <- gsub("； NA", replacement="", flag45$flag45)

#產生檢誤報告文字
flag45 <- flag45 %>%
  subset(select = c(organization_id, flag45)) %>%
  distinct(organization_id, flag45) %>%
  mutate(flag45 = paste(flag45, "（請依欄位說明，修正「教師」、「主任教官」、「教官」之聘任科別中文名稱，「教師」、「主任教官」、「教官」以外其他服務身分別教員之聘任科別請修正為NA。）", sep = ""))
}else{
#偵測flag45是否存在。若不存在，則產生NA行
if('flag45' %in% ls()){
  print("flag45")
}else{
  flag45 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag45$flag45 <- ""
}
}
# flag47: 兼任行政職職稱(一)若填寫“校長"，代表服務身分別填答有誤，故應核對服務身分別與兼任行政職職稱(一)。 -------------------------------------------------------------------
flag_person <- drev_person_1

#標記服務身分別不為"校長，且兼任行政職稱為"校長"
flag_person$err_flag <- 0
flag_person$err_flag <- if_else(flag_person$source == "教員資料表" & flag_person$sertype != "校長" & flag_person$admintitle1 == "校長", 1, flag_person$err_flag)

#加註
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag47 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag47)[3 : length(colnames(flag_person_wide_flag47))]
flag_person_wide_flag47$flag47_r <- NA
for (i in temp){
  flag_person_wide_flag47$flag47_r <- paste(flag_person_wide_flag47$flag47_r, flag_person_wide_flag47[[i]], sep = " ")
}
flag_person_wide_flag47$flag47_r <- gsub("NA ", replacement="", flag_person_wide_flag47$flag47_r)
flag_person_wide_flag47$flag47_r <- gsub(" NA", replacement="", flag_person_wide_flag47$flag47_r)

#產生檢誤報告文字
flag47_temp <- flag_person_wide_flag47 %>%
  group_by(organization_id) %>%
  mutate(flag47_txt = paste(source, "需核對「服務身分別」：", flag47_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag47_txt)) %>%
  distinct(organization_id, flag47_txt)

#根據organization_id，展開成寬資料(wide)
flag47 <- flag47_temp %>%
  dcast(organization_id ~ flag47_txt, value.var = "flag47_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag47)[2 : length(colnames(flag47))]
flag47$flag47 <- NA
for (i in temp){
  flag47$flag47 <- paste(flag47$flag47, flag47[[i]], sep = "； ")
}
flag47$flag47 <- gsub("NA； ", replacement="", flag47$flag47)
flag47$flag47 <- gsub("； NA", replacement="", flag47$flag47)

#產生檢誤報告文字
flag47 <- flag47 %>%
  subset(select = c(organization_id, flag47)) %>%
  distinct(organization_id, flag47) %>%
  mutate(flag47 = paste(flag47, "（請依實際情況並按欄位說明修正）", sep = ""))
}else{
#偵測flag47是否存在。若不存在，則產生NA行
if('flag47' %in% ls()){
  print("flag47")
}else{
  flag47 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag47$flag47 <- ""
}
}
# flag48: 1. 職務名稱與兼任行政職職稱(一)，兩者不應填相同職稱。2. 兼任行政職職稱(一)~(三)，三者不應填相同職稱。-------------------------------------------------------------------
flag_person <- drev_person_1

flag_person$err_admintitle0  <- paste(flag_person$adminunit0, flag_person$admintitle0, sep = "")
flag_person$err_admintitle1  <- paste(flag_person$adminunit1, flag_person$admintitle1, sep = "")
flag_person$err_admintitle2  <- paste(flag_person$adminunit2, flag_person$admintitle2, sep = "")
flag_person$err_admintitle3  <- paste(flag_person$adminunit3, flag_person$admintitle3, sep = "")

#職務名稱與兼任行政職職稱不合理處
flag_person$err_flag <- 0
flag_person$err_flag <- if_else(flag_person$err_admintitle0 == flag_person$err_admintitle1, 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$err_admintitle0 == "NN" & flag_person$err_admintitle1 == "NN", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$err_admintitle0 == flag_person$err_admintitle2, 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$err_admintitle0 == "NN" & flag_person$err_admintitle2 == "NN", 0, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$err_admintitle0 == flag_person$err_admintitle3, 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$err_admintitle0 == "NN" & flag_person$err_admintitle3 == "NN", 0, flag_person$err_flag)
flag_person$err_flag <- if_else((flag_person$err_admintitle1 == flag_person$err_admintitle2) & flag_person$err_admintitle1 != "NN", 1, flag_person$err_flag)
flag_person$err_flag <- if_else((flag_person$err_admintitle1 == flag_person$err_admintitle3) & flag_person$err_admintitle1 != "NN", 1, flag_person$err_flag)
flag_person$err_flag <- if_else((flag_person$err_admintitle2 == flag_person$err_admintitle3) & flag_person$err_admintitle2 != "NN", 1, flag_person$err_flag)

#加註姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag48 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag48)[3 : length(colnames(flag_person_wide_flag48))]
flag_person_wide_flag48$flag48_r <- NA
for (i in temp){
  flag_person_wide_flag48$flag48_r <- paste(flag_person_wide_flag48$flag48_r, flag_person_wide_flag48[[i]], sep = " ")
}
flag_person_wide_flag48$flag48_r <- gsub("NA ", replacement="", flag_person_wide_flag48$flag48_r)
flag_person_wide_flag48$flag48_r <- gsub(" NA", replacement="", flag_person_wide_flag48$flag48_r)

#產生檢誤報告文字
flag48_temp <- flag_person_wide_flag48 %>%
  mutate(flag48_txt = 
    case_when(
      flag48_temp$source == "教員資料表" ~ paste(flag48_temp$source, "需核對「服務身分別」與「兼任行政職職稱(一)」：", flag48_temp$flag48_r, sep = ""),
      flag48_temp$source == "職員工資料表" ~ paste(flag48_temp$source, "「職務名稱」與「兼任行政職職稱」重複：", flag48_temp$flag48_r, sep = "")
    )) %>%
  group_by(organization_id) %>%
  subset(select = c(organization_id, flag48_txt)) %>%
  distinct(organization_id, flag48_txt)

#根據organization_id，展開成寬資料(wide)
flag48 <- flag48_temp %>%
  dcast(organization_id ~ flag48_txt, value.var = "flag48_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag48)[2 : length(colnames(flag48))]
flag48$flag48 <- NA
for (i in temp){
  flag48$flag48 <- paste(flag48$flag48, flag48[[i]], sep = "； ")
}
flag48$flag48 <- gsub("NA； ", replacement="", flag48$flag48)
flag48$flag48 <- gsub("； NA", replacement="", flag48$flag48)

#產生檢誤報告文字
flag48 <- flag48 %>%
  subset(select = c(organization_id, flag48)) %>%
  distinct(organization_id, flag48) %>%
  mutate(flag48 = paste(flag48, "（以上人員之專職工作職稱請填入『職務名稱』，非『兼任行政職職稱』。併請確認以上人員除本職職務外，是否再兼任其他職務）", sep = ""))
}else{
#偵測flag48是否存在。若不存在，則產生NA行
if('flag48' %in% ls()){
  print("flag48")
}else{
  flag48 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag48$flag48 <- ""
}
}
# flag49: 1.	職員(工)的「職務名稱」不應填N（全型或半型皆不行）。-------------------------------------------------------------------
# 2. 職員(工)的「服務單位」不應填N（全型或半型皆不行），且應入填入對應職稱的學校內部單位。
flag_person <- drev_person_1

#標記職務名稱、服務單位為N或非學校內部單位
flag_person$err_adm1 <- 1
flag_person$err_adm1 <- if_else(grepl("組", flag_person$adminunit0), 0, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(grepl("室", flag_person$adminunit0), 0, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(grepl("科", flag_person$adminunit0), 0, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(grepl("中心", flag_person$adminunit0), 0, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(grepl("部", flag_person$adminunit0), 0, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(grepl("辦公室", flag_person$adminunit0), 0, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(grepl("館", flag_person$adminunit0), 0, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(grepl("處", flag_person$adminunit0), 0, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(grepl("部", flag_person$adminunit0), 0, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(grepl("^社區大學$", flag_person$adminunit0), 0, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(flag_person$adminunit0 == "", 0, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(flag_person$adminunit0 == "董事會", 0, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(flag_person$adminunit0 == "實習農場", 0, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(flag_person$adminunit0 == "NA", 1, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(flag_person$adminunit0 == "n", 1, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(flag_person$adminunit0 == "N", 1, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(flag_person$adminunit0 == "Ｎ", 1, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(flag_person$adminunit0 == "國教署", 1, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(flag_person$adminunit0 == "y", 1, flag_person$err_adm1)
flag_person$err_adm1 <- if_else(flag_person$adminunit0 == "Y", 1, flag_person$err_adm1)

flag_person$err_adm2 <- 0
flag_person$err_adm2 <- if_else(flag_person$admintitle0 == "NA", 1, flag_person$err_adm2)
flag_person$err_adm2 <- if_else(flag_person$admintitle0 == "N", 1, flag_person$err_adm2)
flag_person$err_adm2 <- if_else(flag_person$admintitle0 == "Ｎ", 1, flag_person$err_adm2)
flag_person$err_adm2 <- if_else(flag_person$admintitle0 == "n", 1, flag_person$err_adm2)
flag_person$err_adm2 <- if_else(flag_person$admintitle0 == "Y", 1, flag_person$err_adm2)
flag_person$err_adm2 <- if_else(flag_person$admintitle0 == "y", 1, flag_person$err_adm2)
flag_person$err_adm2 <- if_else(grepl("Ｎ", flag_person$admintitle0), 1, flag_person$err_adm2)

flag_person$err_flag <- flag_person$err_adm1 + flag_person$err_adm2
flag_person$err_adm <- 0
flag_person$err_adm <- if_else(flag_person$err_flag != 0 & flag_person$source == "職員工資料表", 1, flag_person$err_adm)

#加註
flag_person$name <- paste(flag_person$name, "（", sep = "")
flag_person$name <- if_else(flag_person$err_adm2 != 0, paste(flag_person$name, "職務名稱：", flag_person$admintitle0, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_adm1 != 0, paste(flag_person$name, "服務單位：", flag_person$adminunit0, "；", sep = ""), flag_person$name)
flag_person$name <- paste(flag_person$name, "）", sep = "")
flag_person$name <- gsub("；）", replacement = "）", flag_person$name)
flag_person$name <- gsub("（）", replacement = "", flag_person$name)


flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_adm == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_adm == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag49 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_adm)) %>%
  subset(err_adm == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag49)[3 : length(colnames(flag_person_wide_flag49))]
flag_person_wide_flag49$flag49_r <- NA
for (i in temp){
  flag_person_wide_flag49$flag49_r <- paste(flag_person_wide_flag49$flag49_r, flag_person_wide_flag49[[i]], sep = " ")
}
flag_person_wide_flag49$flag49_r <- gsub("NA ", replacement="", flag_person_wide_flag49$flag49_r)
flag_person_wide_flag49$flag49_r <- gsub(" NA", replacement="", flag_person_wide_flag49$flag49_r)

#產生檢誤報告文字
flag49_temp <- flag_person_wide_flag49 %>%
  group_by(organization_id) %>%
  mutate(flag49_txt = paste(source, "：", flag49_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag49_txt)) %>%
  distinct(organization_id, flag49_txt)

#根據organization_id，展開成寬資料(wide)
flag49 <- flag49_temp %>%
  dcast(organization_id ~ flag49_txt, value.var = "flag49_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag49)[2 : length(colnames(flag49))]
flag49$flag49 <- NA
for (i in temp){
  flag49$flag49 <- paste(flag49$flag49, flag49[[i]], sep = "； ")
}
flag49$flag49 <- gsub("NA； ", replacement="", flag49$flag49)
flag49$flag49 <- gsub("； NA", replacement="", flag49$flag49)

#產生檢誤報告文字
flag49 <- flag49 %>%
  subset(select = c(organization_id, flag49)) %>%
  distinct(organization_id, flag49) %>%
  mutate(flag49 = paste(flag49, "（請確認『職務名稱』、『服務單位』）", sep = ""))
}else{
#偵測flag49是否存在。若不存在，則產生NA行
if('flag49' %in% ls()){
  print("flag49")
}else{
  flag49 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag49$flag49 <- ""
}
}
# flag50: 留職停薪原因為「應徵入伍留職停薪」、「奉派協助友邦留職停薪」、「育嬰留職停薪」、「侍親留職停薪」、「依親留職停薪」、「出國進修或研究留職停薪」、「易服勞役留職停薪」、「延長留職停薪」、「照護配偶或子女留職停薪」、「國內外進修期滿延長留職停薪」、「延長病假期滿留職停薪」、「因公傷病公假期滿留職停薪」、「留職停薪/停聘」、「其他情事留職停薪」，在借調類別應填寫N。 -------------------------------------------------------------------
flag_person <- drev_person_1

flag_person$brtype <- if_else(flag_person$brtype == "NA", "N", flag_person$brtype)

#留職停薪原因不合理處
flag_person$err_lev <- 0
flag_person$err_lev <- if_else(flag_person$levpay == "應徵入伍留職停薪" & flag_person$brtype != "N", 1, flag_person$err_lev)
flag_person$err_lev <- if_else(flag_person$levpay == "奉派協助友邦留職停薪" & flag_person$brtype != "N", 1, flag_person$err_lev)
flag_person$err_lev <- if_else(flag_person$levpay == "育嬰留職停薪" & flag_person$brtype != "N", 1, flag_person$err_lev)
flag_person$err_lev <- if_else(flag_person$levpay == "侍親留職停薪" & flag_person$brtype != "N", 1, flag_person$err_lev)
flag_person$err_lev <- if_else(flag_person$levpay == "依親留職停薪" & flag_person$brtype != "N", 1, flag_person$err_lev)
flag_person$err_lev <- if_else(flag_person$levpay == "出國進修或研究留職停薪" & flag_person$brtype != "N", 1, flag_person$err_lev)
flag_person$err_lev <- if_else(flag_person$levpay == "易服勞役留職停薪" & flag_person$brtype != "N", 1, flag_person$err_lev)
flag_person$err_lev <- if_else(flag_person$levpay == "延長留職停薪" & flag_person$brtype != "N", 1, flag_person$err_lev)
flag_person$err_lev <- if_else(flag_person$levpay == "易服勞役留職停薪" & flag_person$brtype != "N", 1, flag_person$err_lev)
flag_person$err_lev <- if_else(flag_person$levpay == "照護配偶或子女留職停薪" & flag_person$brtype != "N", 1, flag_person$err_lev)
flag_person$err_lev <- if_else(flag_person$levpay == "國內外進修期滿延長留職停薪" & flag_person$brtype != "N", 1, flag_person$err_lev)
flag_person$err_lev <- if_else(flag_person$levpay == "延長病假期滿留職停薪" & flag_person$brtype != "N", 1, flag_person$err_lev)
flag_person$err_lev <- if_else(flag_person$levpay == "留職停薪/停聘" & flag_person$brtype != "N", 1, flag_person$err_lev)
flag_person$err_lev <- if_else(flag_person$levpay == "其他情事留職停薪" & flag_person$brtype != "N", 1, flag_person$err_lev)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_lev == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_lev == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag50 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_lev)) %>%
  subset(err_lev == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag50)[3 : length(colnames(flag_person_wide_flag50))]
flag_person_wide_flag50$flag50_r <- NA
for (i in temp){
  flag_person_wide_flag50$flag50_r <- paste(flag_person_wide_flag50$flag50_r, flag_person_wide_flag50[[i]], sep = " ")
}
flag_person_wide_flag50$flag50_r <- gsub("NA ", replacement="", flag_person_wide_flag50$flag50_r)
flag_person_wide_flag50$flag50_r <- gsub(" NA", replacement="", flag_person_wide_flag50$flag50_r)

#產生檢誤報告文字
flag50_temp <- flag_person_wide_flag50 %>%
  group_by(organization_id) %>%
  mutate(flag50_txt = paste(source, "需核對「留職停薪原因」與「借調類別」：", flag50_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag50_txt)) %>%
  distinct(organization_id, flag50_txt)

#根據organization_id，展開成寬資料(wide)
flag50 <- flag50_temp %>%
  dcast(organization_id ~ flag50_txt, value.var = "flag50_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag50)[2 : length(colnames(flag50))]
flag50$flag50 <- NA
for (i in temp){
  flag50$flag50 <- paste(flag50$flag50, flag50[[i]], sep = "； ")
}
flag50$flag50 <- gsub("NA； ", replacement="", flag50$flag50)
flag50$flag50 <- gsub("； NA", replacement="", flag50$flag50)

#產生檢誤報告文字
flag50 <- flag50 %>%
  subset(select = c(organization_id, flag50)) %>%
  distinct(organization_id, flag50) %>%
  mutate(flag50 = paste(flag50, "", sep = ""))
}else{
#偵測flag50是否存在。若不存在，則產生NA行
if('flag50' %in% ls()){
  print("flag50")
}else{
  flag50 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag50$flag50 <- ""
}
}
# flag51: 原則上，「留職停薪原因」與「借調類別」填寫應相符:1.	借調公民營事業機構留職停薪↔借調至公民營事業機構 2.	借調行政法人機關留職停薪↔借調至行政法人機關 3.借調法定實驗學校留職停薪↔借調至法定實驗學校-------------------------------------------------------------------
flag_person <- drev_person_1

#「留職停薪原因」與「借調類別」不合理處
flag_person$err_lev <- 0
flag_person$err_lev <- if_else(flag_person$brtype == "借調至公民營事業機構"  & flag_person$levpay != "借調公民營事業機構留職停薪", 1, flag_person$err_lev)
flag_person$err_lev <- if_else(flag_person$brtype == "借調至行政法人機關"  & flag_person$levpay != "借調行政法人機關留職停薪", 1, flag_person$err_lev)
flag_person$err_lev <- if_else(flag_person$brtype == "借調至法定實驗學校"  & flag_person$levpay != "借調法定實驗學校留職停薪", 1, flag_person$err_lev)

flag_person$err_lev <- if_else(flag_person$brtype != "借調至公民營事業機構"  & flag_person$levpay == "借調公民營事業機構留職停薪", 1, flag_person$err_lev)
flag_person$err_lev <- if_else(flag_person$brtype != "借調至行政法人機關"  & flag_person$levpay == "借調行政法人機關留職停薪", 1, flag_person$err_lev)
flag_person$err_lev <- if_else(flag_person$brtype != "借調至法定實驗學校"  & flag_person$levpay == "借調法定實驗學校留職停薪", 1, flag_person$err_lev)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_lev == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_lev == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag51 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_lev)) %>%
  subset(err_lev == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag51)[3 : length(colnames(flag_person_wide_flag51))]
flag_person_wide_flag51$flag51_r <- NA
for (i in temp){
  flag_person_wide_flag51$flag51_r <- paste(flag_person_wide_flag51$flag51_r, flag_person_wide_flag51[[i]], sep = " ")
}
flag_person_wide_flag51$flag51_r <- gsub("NA ", replacement="", flag_person_wide_flag51$flag51_r)
flag_person_wide_flag51$flag51_r <- gsub(" NA", replacement="", flag_person_wide_flag51$flag51_r)

#產生檢誤報告文字
flag51_temp <- flag_person_wide_flag51 %>%
  group_by(organization_id) %>%
  mutate(flag51_txt = paste(source, "需核對「留職停薪原因」與「借調類別」：", flag51_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag51_txt)) %>%
  distinct(organization_id, flag51_txt)

#根據organization_id，展開成寬資料(wide)
flag51 <- flag51_temp %>%
  dcast(organization_id ~ flag51_txt, value.var = "flag51_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag51)[2 : length(colnames(flag51))]
flag51$flag51 <- NA
for (i in temp){
  flag51$flag51 <- paste(flag51$flag51, flag51[[i]], sep = "； ")
}
flag51$flag51 <- gsub("NA； ", replacement="", flag51$flag51)
flag51$flag51 <- gsub("； NA", replacement="", flag51$flag51)

#產生檢誤報告文字
flag51 <- flag51 %>%
  subset(select = c(organization_id, flag51)) %>%
  distinct(organization_id, flag51) %>%
  mutate(flag51 = paste(flag51, "", sep = ""))
}else{
#偵測flag51是否存在。若不存在，則產生NA行
if('flag51' %in% ls()){
  print("flag51")
}else{
  flag51 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag51$flag51 <- ""
}
}
# flag52: 留職停薪原因若填寫「應徵入伍留職停薪」、「奉派協助友邦留職停薪」、「育嬰留職停薪」、「侍親留職停薪」、「依親留職停薪」、「出國進修或研究留職停薪」、「易服勞役留職停薪」、「延長留職停薪」、「照護配偶或子女留職停薪」、「國內外進修期滿延長留職停薪」、「延長病假期滿留職停薪」、「因公傷病公假期滿留職停薪」、「留職停薪/停聘」、「其他情事留職停薪」、「借調公務機關留職停薪」、「借調公民營事業機構留職停薪」、「借調行政法人機關留職停薪」、「借調法定實驗學校留職停薪」，在商借類別應填寫N。 -------------------------------------------------------------------
flag_person <- drev_person_1

flag_person$levpay <- if_else(flag_person$levpay == "NA", "N", flag_person$levpay)
flag_person$negle <- if_else(flag_person$negle == "NA", "N", flag_person$negle)

#「留職停薪原因」、「商借類別」不合理處
flag_person$err_lev <- 0
flag_person$err_lev <- if_else(flag_person$levpay != "N"  & flag_person$negle != "N" & flag_person$source == "教員資料表", 1, flag_person$err_lev)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_lev == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_lev == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag52 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_lev)) %>%
  subset(err_lev == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag52)[3 : length(colnames(flag_person_wide_flag52))]
flag_person_wide_flag52$flag52_r <- NA
for (i in temp){
  flag_person_wide_flag52$flag52_r <- paste(flag_person_wide_flag52$flag52_r, flag_person_wide_flag52[[i]], sep = " ")
}
flag_person_wide_flag52$flag52_r <- gsub("NA ", replacement="", flag_person_wide_flag52$flag52_r)
flag_person_wide_flag52$flag52_r <- gsub(" NA", replacement="", flag_person_wide_flag52$flag52_r)

#產生檢誤報告文字
flag52_temp <- flag_person_wide_flag52 %>%
  group_by(organization_id) %>%
  mutate(flag52_txt = paste(source, "需核對「留職停薪原因」與「商借類別」：", flag52_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag52_txt)) %>%
  distinct(organization_id, flag52_txt)

#根據organization_id，展開成寬資料(wide)
flag52 <- flag52_temp %>%
  dcast(organization_id ~ flag52_txt, value.var = "flag52_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag52)[2 : length(colnames(flag52))]
flag52$flag52 <- NA
for (i in temp){
  flag52$flag52 <- paste(flag52$flag52, flag52[[i]], sep = "； ")
}
flag52$flag52 <- gsub("NA； ", replacement="", flag52$flag52)
flag52$flag52 <- gsub("； NA", replacement="", flag52$flag52)

#產生檢誤報告文字
flag52 <- flag52 %>%
  subset(select = c(organization_id, flag52)) %>%
  distinct(organization_id, flag52) %>%
  mutate(flag52 = paste(flag52, "", sep = ""))
}else{
#偵測flag52是否存在。若不存在，則產生NA行
if('flag52' %in% ls()){
  print("flag52")
}else{
  flag52 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag52$flag52 <- ""
}
}
# flag57: 學歷資料各學位別畢業學校國別/校名/科系所之（一）與（二）內容不應相同，請依學位取得實際情況修正。-------------------------------------------------------------------
flag_person <- drev_person_1

#「留職停薪原因」與「借調類別」不合理處
flag_person$err_degree <- 0
flag_person$err_degree <- if_else((flag_person$ddegreen1 == flag_person$ddegreen2 & flag_person$ddegreen1 != "N") & (flag_person$ddegreeu1 == flag_person$ddegreeu2 & flag_person$ddegreeu1 != "N") & (flag_person$ddegreeg1 == flag_person$ddegreeg2 & flag_person$ddegreeg1 != "N"), 1, flag_person$err_degree)
flag_person$err_degree <- if_else((flag_person$mdegreen1 == flag_person$mdegreen2 & flag_person$mdegreen1 != "N") & (flag_person$mdegreeu1 == flag_person$ddegreeu2 & flag_person$mdegreeu1 != "N") & (flag_person$mdegreeg1 == flag_person$ddegreeg2 & flag_person$mdegreeg1 != "N"), 1, flag_person$err_degree)
flag_person$err_degree <- if_else((flag_person$bdegreen1 == flag_person$bdegreen2 & flag_person$bdegreen1 != "N") & (flag_person$bdegreeu1 == flag_person$bdegreeu2 & flag_person$bdegreeu1 != "N") & (flag_person$bdegreeg1 == flag_person$bdegreeg2 & flag_person$bdegreeg1 != "N"), 1, flag_person$err_degree)
flag_person$err_degree <- if_else((flag_person$adegreen1 == flag_person$adegreen2 & flag_person$adegreen1 != "N") & (flag_person$adegreeu1 == flag_person$adegreeu2 & flag_person$adegreeu1 != "N") & (flag_person$adegreeg1 == flag_person$adegreeg2 & flag_person$adegreeg1 != "N"), 1, flag_person$err_degree)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_degree == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_degree == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag57 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_degree)) %>%
  subset(err_degree == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag57)[3 : length(colnames(flag_person_wide_flag57))]
flag_person_wide_flag57$flag57_r <- NA
for (i in temp){
  flag_person_wide_flag57$flag57_r <- paste(flag_person_wide_flag57$flag57_r, flag_person_wide_flag57[[i]], sep = " ")
}
flag_person_wide_flag57$flag57_r <- gsub("NA ", replacement="", flag_person_wide_flag57$flag57_r)
flag_person_wide_flag57$flag57_r <- gsub(" NA", replacement="", flag_person_wide_flag57$flag57_r)

#產生檢誤報告文字
flag57_temp <- flag_person_wide_flag57 %>%
  group_by(organization_id) %>%
  mutate(flag57_txt = paste("請檢視修正學歷資訊內容：", source, "：", flag57_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag57_txt)) %>%
  distinct(organization_id, flag57_txt)

#根據organization_id，展開成寬資料(wide)
flag57 <- flag57_temp %>%
  dcast(organization_id ~ flag57_txt, value.var = "flag57_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag57)[2 : length(colnames(flag57))]
flag57$flag57 <- NA
for (i in temp){
  flag57$flag57 <- paste(flag57$flag57, flag57[[i]], sep = "； ")
}
flag57$flag57 <- gsub("NA； ", replacement="", flag57$flag57)
flag57$flag57 <- gsub("； NA", replacement="", flag57$flag57)

#產生檢誤報告文字
flag57 <- flag57 %>%
  subset(select = c(organization_id, flag57)) %>%
  distinct(organization_id, flag57) %>%
  mutate(flag57 = paste(flag57, "", sep = ""))
}else{
#偵測flag57是否存在。若不存在，則產生NA行
if('flag57' %in% ls()){
  print("flag57")
}else{
  flag57 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag57$flag57 <- ""
}
}
# flag59: 校長之聘任類別需為「專任」。-------------------------------------------------------------------
flag_person <- drev_person_1

#校長之聘任類別不為專任 不合理
flag_person$err_flag <- 0
flag_person$err_flag <- if_else(flag_person$sertype == "校長" & flag_person$emptype != "專任", 1, flag_person$err_flag)
#師大附中及高師大附中校長為兼任
flag_person$err_flag <- if_else(flag_person$organization_id == "330301" | flag_person$organization_id == "580301", 0, flag_person$err_flag)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag59 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag59)[3 : length(colnames(flag_person_wide_flag59))]
flag_person_wide_flag59$flag59_r <- NA
for (i in temp){
  flag_person_wide_flag59$flag59_r <- paste(flag_person_wide_flag59$flag59_r, flag_person_wide_flag59[[i]], sep = " ")
}
flag_person_wide_flag59$flag59_r <- gsub("NA ", replacement="", flag_person_wide_flag59$flag59_r)
flag_person_wide_flag59$flag59_r <- gsub(" NA", replacement="", flag_person_wide_flag59$flag59_r)

#產生檢誤報告文字
flag59_temp <- flag_person_wide_flag59 %>%
  group_by(organization_id) %>%
  mutate(flag59_txt = paste("校長之聘任類別需為「專任」。", sep = ""), "") %>%
  subset(select = c(organization_id, flag59_txt)) %>%
  distinct(organization_id, flag59_txt)

#根據organization_id，展開成寬資料(wide)
flag59 <- flag59_temp %>%
  dcast(organization_id ~ flag59_txt, value.var = "flag59_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag59)[2 : length(colnames(flag59))]
flag59$flag59 <- NA
for (i in temp){
  flag59$flag59 <- paste(flag59$flag59, flag59[[i]], sep = "； ")
}
flag59$flag59 <- gsub("NA； ", replacement="", flag59$flag59)
flag59$flag59 <- gsub("； NA", replacement="", flag59$flag59)

#產生檢誤報告文字
flag59 <- flag59 %>%
  subset(select = c(organization_id, flag59)) %>%
  distinct(organization_id, flag59) %>%
  mutate(flag59 = paste(flag59, "", sep = ""))
}else{
#偵測flag59是否存在。若不存在，則產生NA行
if('flag59' %in% ls()){
  print("flag59")
}else{
  flag59 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag59$flag59 <- ""
}
}
# flag62: 職務名稱及兼任行政職職稱（一）～（三）資料內容是否完整正確。-------------------------------------------------------------------
#如：1.	職務名稱及兼任行政職職稱（一）～（三）填入非職稱內容。
#2.	服務單位及兼任行政職服務單位（一）～（三）填入非服務單位內容。
#3.	校長、教官、主任教官屬於教員，故應填至教員資料表。
#職員工的「職務名稱」不應有教師、老師等非行政工作之名稱。
flag_person <- drev_person_1

#職務名稱
flag_person$err_admintitle0 <- 1
flag_person$err_admintitle0 <- if_else(grepl("主任$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("佐理員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("助理$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("人員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("助理員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("技士", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("技工", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("技佐", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("防護員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("組長$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("組員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("管理員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("管理師$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("輔導員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("工友$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^約僱", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^約聘", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("書記$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("幹事", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^學務創新", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("創新人力", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("營養師$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^職輔員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("護士$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("護理師$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^均質化承辦人$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^均職化承辦人$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^優質化協辦人$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("校安$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("心理師$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("技術員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("職輔員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("廚工$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("編制外行政人力$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("司機$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("秘書$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("祕書$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("舍監$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("辦事員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("職務代理$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("職代$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("職務代理人$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("救生員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("值機員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("監督$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("三副$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("社工師$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("校護$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("專員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("雇員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("僱員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^充實行政人力$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("1", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("工讀生$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("警衛$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^實習餐廳經理$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^清潔員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^清潔$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("佐理$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^會計員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^水電$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^總機$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^園藝$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("電工$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^木工$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^守衛$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("校長", flag_person$admintitle0), 1, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^副校長$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("館員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^出納$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("庶務$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("環保$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("體衛$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^書院Coach$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("執行長$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("顧問$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^助教$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^督導$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("教師$", flag_person$admintitle0), 1, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("老師$", flag_person$admintitle0), 1, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("導師$", flag_person$admintitle0), 1, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("教學支援人員$", flag_person$admintitle0), 1, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^董事長$", flag_person$admintitle0), 1, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("指導員$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("補充行政人力$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^學創人力$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("書記\\(控障-公務人員\\)$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^校警$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^專任行政人力$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("安心上工", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("護理員", flag_person$admintitle0), 0, flag_person$err_admintitle0)
#光禾華德福才可有以下職稱：總務、學務、教務、輔導、人事、國中部行政、高中部行政、會計
flag_person$err_admintitle0 <- if_else(grepl("^國中部行政$", flag_person$admintitle0) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^高中部行政$", flag_person$admintitle0) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^總務$", flag_person$admintitle0) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^學務$", flag_person$admintitle0) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^教務$", flag_person$admintitle0) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^輔導$", flag_person$admintitle0) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^人事$", flag_person$admintitle0) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^會計$", flag_person$admintitle0) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle0)
#東方工商才可有以下職稱：職員
flag_person$err_admintitle0 <- if_else(grepl("^職員$", flag_person$admintitle0) & flag_person$organization_id == "331402", 0, flag_person$err_admintitle0)
#仁義高中才可有以下職稱：會計
flag_person$err_admintitle0 <- if_else(grepl("^會計$", flag_person$admintitle0) & flag_person$organization_id == "201309", 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(flag_person$source == 1, 0, flag_person$err_admintitle0)
#運動教練已在flag34檢查
flag_person$err_admintitle0 <- if_else(grepl("教練$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
#教官已在flag15檢查
flag_person$err_admintitle0 <- if_else(grepl("教官$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
#N或NA已在flag49檢查
flag_person$err_admintitle0 <- if_else(grepl("^N$", flag_person$admintitle0), 0, flag_person$err_admintitle0)
flag_person$err_admintitle0 <- if_else(grepl("^NA$", flag_person$admintitle0), 0, flag_person$err_admintitle0)


#服務單位
flag_person$err_adminunit0 <- 1
flag_person$err_adminunit0 <- if_else(grepl("^人事室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^主計室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^校長室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^副校長室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^特教辦公室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^秘書室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^祕書室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^小學部$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^國小部$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^國中部$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^中學部$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^國教署$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^教官室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^教官室\\(軍訓室\\)$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^教務處", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^進修部$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^進修學校$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^圖書館$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^圖書室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^實習處", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^實習室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^實習農場$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^實習輔導處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^輔導室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("輔導處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^學生事務處", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^學務處", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^總務處", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^雙語部$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^專案辦公室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^軍訓室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^會計室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^資訊室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^國際部$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^董事會$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("中心$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("研究發展處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^保健室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^招生處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^公關室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^農場經營$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^藝文中心", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("科$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("1", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("組$", flag_person$adminunit0) & (grepl("*處*", flag_person$adminunit0) | grepl("*中心*", flag_person$adminunit0) | grepl("*部*", flag_person$adminunit0) | grepl("*室*", flag_person$adminunit0) | grepl("*館*", flag_person$adminunit0)), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^國中部教務處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^高中部教務處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^圖資室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^圖資處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^實習就業處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^總務室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^社區大學$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^國際教育處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^研發處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^資訊處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^實輔處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^住校處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^發展事務處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^家具設計發展處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^國際處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^研發室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^圖資室兼技術交流處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^教育推廣處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^學輔處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^顧問室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^研究發展室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^小學部籌備處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^綜合高中$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^技術交流處$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("實驗室$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^資源班$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
#立志高中才有"高國中部"
flag_person$err_adminunit0 <- if_else(grepl("^高國中部$", flag_person$adminunit0) & flag_person$organization_id == "551301", 0, flag_person$err_adminunit0)
#光禾華德福才可有以下服務單位：國中部日間部、高中部日間部
flag_person$err_adminunit0 <- if_else(grepl("^國中部日間部$", flag_person$adminunit0) & flag_person$organization_id == "121302", 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^高中部日間部$", flag_person$adminunit0) & flag_person$organization_id == "121302", 0, flag_person$err_adminunit0)
#N或NA已在flag49檢查
flag_person$err_adminunit0 <- if_else(grepl("^N$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(grepl("^NA$", flag_person$adminunit0), 0, flag_person$err_adminunit0)
flag_person$err_adminunit0 <- if_else(flag_person$source == 1, 0, flag_person$err_adminunit0)


#兼任行政職職稱（一）
flag_person$err_admintitle1 <- 1
flag_person$err_admintitle1 <- if_else(grepl("主任$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("秘書$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("組長$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("組員$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^副校長$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^均質化承辦人$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^均職化承辦人$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^優質化協辦人$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("校安$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("心理師$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("技術員$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("人員$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("職輔員$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("廚工$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("營養師$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("編制外行政人力$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("司機$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("秘書$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("祕書$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("舍監$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("辦事員$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^幹事$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("職務代理$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("職代$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("職務代理人$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("救生員$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("值機員$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("監督$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("三副$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("社工師$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("助理$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("專員$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("政風$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("1", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^N$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("輔導員$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^警衛$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^實習餐廳經理$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("管理師$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^清潔員$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("佐理$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("技佐$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^會計員$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("管理員$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^書記$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("佐理員$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("館員$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("科主席$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("護理師$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("助理員$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^助教$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("庶務$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("文書$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^電競專案教練$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^出納$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^助教$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^督導$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("訓育業務$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("指導員$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
#光禾華德福才可有以下職稱：總務、學務、教務、輔導、人事、國中部行政、高中部行政、會計
flag_person$err_admintitle1 <- if_else(grepl("^國中部行政$", flag_person$admintitle1) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^高中部行政$", flag_person$admintitle1) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^總務$", flag_person$admintitle1) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^學務$", flag_person$admintitle1) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^教務$", flag_person$admintitle1) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^輔導$", flag_person$admintitle1) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^人事$", flag_person$admintitle1) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^會計$", flag_person$admintitle1) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^校長$", flag_person$admintitle1), 1, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("教官$", flag_person$admintitle1), 1, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("教師$", flag_person$admintitle1), 1, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("運動教練$", flag_person$admintitle1), 1, flag_person$err_admintitle1)
flag_person$err_admintitle1 <- if_else(grepl("^董事長$", flag_person$admintitle1), 1, flag_person$err_admintitle1)
#校長已在flag15檢查
flag_person$err_admintitle1 <- if_else(grepl("^校長$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
#導師已在flag15檢查
flag_person$err_admintitle1 <- if_else(grepl("導師$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
#教師已在flag15檢查
flag_person$err_admintitle1 <- if_else(grepl("教師$", flag_person$admintitle1), 0, flag_person$err_admintitle1)
#教官已在flag15檢查
flag_person$err_admintitle1 <- if_else(grepl("教官$", flag_person$admintitle1), 0, flag_person$err_admintitle1)


#兼任行政職服務單位（一）
flag_person$err_adminunit1 <- if_else(grepl("^校長室$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^副校長室$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^秘書室$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^國小部", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^小學部$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^國中部$", flag_person$adminunit1) | grepl("^國民中學部$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^中學部$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^教官室$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^教務處", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^進修部", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^進修學校$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^資訊室$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^圖書館", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^圖書室$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^實習處", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^實習室$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^輔導室", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("輔導處", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^學生事務處", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^學務處", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^總務處", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^雙語部", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^軍訓室$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^會計室$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^主計室$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^資訊室$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^國際部$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^董事會$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("中心$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("研究發展處$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("人事室$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^招生處$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^公關室$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^農場經營$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^藝文中心", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("科$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("1", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("組$", flag_person$adminunit1) & (grepl("*處*", flag_person$adminunit1) | grepl("*中心*", flag_person$adminunit1) | grepl("*部*", flag_person$adminunit1) | grepl("*室*", flag_person$adminunit1) | grepl("*館*", flag_person$adminunit1)), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^N$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^國中部教務處$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^高中部教務處$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^圖資室$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^圖資處$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^國際教育處$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^研發處$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^資訊處$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^實輔處$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^家具設計發展處$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^國際處$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^研發室$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^圖資室兼技術交流處$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^教育推廣處$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^學輔處$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^綜合高中$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^技術交流處$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^補校$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^國際交流處$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^專案研究室$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("分校$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^創發處$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^補校教學組$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^補校教務組$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^補校訓育組$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^實驗研究組$", flag_person$adminunit1), 0, flag_person$err_adminunit1)
#立志高中才有"高國中部"
flag_person$err_adminunit1 <- if_else(grepl("^高國中部$", flag_person$adminunit1) & flag_person$organization_id == "551301", 0, flag_person$err_adminunit1)
#光禾華德福才可有以下服務單位：國中部日間部、高中部日間部
flag_person$err_adminunit1 <- if_else(grepl("^國中部日間部$", flag_person$adminunit1) & flag_person$organization_id == "121302", 0, flag_person$err_adminunit1)
flag_person$err_adminunit1 <- if_else(grepl("^高中部日間部$", flag_person$adminunit1) & flag_person$organization_id == "121302", 0, flag_person$err_adminunit1)
#磐石高中才有"高中部"
flag_person$err_adminunit1 <- if_else(grepl("^高中部$", flag_person$adminunit1) & flag_person$organization_id == "181307", 0, flag_person$err_adminunit1)


#兼任行政職職稱（二）
flag_person$err_admintitle2 <- 1
flag_person$err_admintitle2 <- if_else(grepl("主任$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("秘書$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("組長$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("組員$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^副校長$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^均質化承辦人$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^均職化承辦人$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^優質化協辦人$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("校安$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("心理師$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("技術員$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("人員$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("職輔員$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("廚工$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("營養師$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("編制外行政人力$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("司機$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("秘書$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("祕書$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("舍監$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("辦事員$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^幹事$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("職務代理$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("職代$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("職務代理人$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("救生員$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("值機員$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("監督$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("三副$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("社工師$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("助理$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("專員$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("政風$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("1", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^N$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("輔導員$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^警衛$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^實習餐廳經理$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("管理師$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^清潔員$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("佐理$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("技佐$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^會計員$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("管理員$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^書記$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("佐理員$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("館員$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("科主席$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("護理師$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("助理員$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^助教$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("庶務$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("文書$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^電競專案教練$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^出納$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^助教$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^督導$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("訓育業務$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("指導員$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
#光禾華德福才可有以下職稱：總務、學務、教務、輔導、人事、國中部行政、高中部行政、會計
flag_person$err_admintitle2 <- if_else(grepl("^國中部行政$", flag_person$admintitle2) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^高中部行政$", flag_person$admintitle2) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^總務$", flag_person$admintitle2) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^學務$", flag_person$admintitle2) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^教務$", flag_person$admintitle2) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^輔導$", flag_person$admintitle2) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^人事$", flag_person$admintitle2) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^會計$", flag_person$admintitle2) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^校長$", flag_person$admintitle2), 1, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("教官$", flag_person$admintitle2), 1, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("教師$", flag_person$admintitle2), 1, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("運動教練$", flag_person$admintitle2), 1, flag_person$err_admintitle2)
flag_person$err_admintitle2 <- if_else(grepl("^董事長$", flag_person$admintitle2), 1, flag_person$err_admintitle2)
#校長已在flag15檢查
flag_person$err_admintitle2 <- if_else(grepl("^校長$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
#導師已在flag15檢查
flag_person$err_admintitle2 <- if_else(grepl("導師$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
#教師已在flag15檢查
flag_person$err_admintitle2 <- if_else(grepl("教師$", flag_person$admintitle2), 0, flag_person$err_admintitle2)
#教官已在flag15檢查
flag_person$err_admintitle2 <- if_else(grepl("教官$", flag_person$admintitle2), 0, flag_person$err_admintitle2)


#兼任行政職服務單位（二）
flag_person$err_adminunit2 <- if_else(grepl("^校長室$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^副校長室$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^秘書室$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^國小部", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^小學部$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^國中部$", flag_person$adminunit2) | grepl("^國民中學部$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^中學部$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^教官室$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^教務處", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^進修部", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^進修學校$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^資訊室$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^圖書館", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^圖書室$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^實習處", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^實習室$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^輔導室", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("輔導處", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^學生事務處", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^學務處", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^總務處", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^雙語部", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^軍訓室$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^會計室$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^主計室$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^資訊室$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^國際部$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^董事會$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("中心$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("研究發展處$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("人事室$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^招生處$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^公關室$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^農場經營$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^藝文中心", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("科$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("1", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("組$", flag_person$adminunit2) & (grepl("*處*", flag_person$adminunit2) | grepl("*中心*", flag_person$adminunit2) | grepl("*部*", flag_person$adminunit2) | grepl("*室*", flag_person$adminunit2) | grepl("*館*", flag_person$adminunit2)), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^N$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^國中部教務處$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^高中部教務處$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^圖資室$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^圖資處$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^國際教育處$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^研發處$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^資訊處$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^實輔處$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^家具設計發展處$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^國際處$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^研發室$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^圖資室兼技術交流處$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^教育推廣處$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^學輔處$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^綜合高中$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^技術交流處$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^補校$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^國際交流處$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^專案研究室$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("分校$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^創發處$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^補校教學組$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^補校教務組$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^補校訓育組$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^實驗研究組$", flag_person$adminunit2), 0, flag_person$err_adminunit2)
#立志高中才有"高國中部"
flag_person$err_adminunit2 <- if_else(grepl("^高國中部$", flag_person$adminunit2) & flag_person$organization_id == "551301", 0, flag_person$err_adminunit2)
#光禾華德福才可有以下服務單位：國中部日間部、高中部日間部
flag_person$err_adminunit2 <- if_else(grepl("^國中部日間部$", flag_person$adminunit2) & flag_person$organization_id == "121302", 0, flag_person$err_adminunit2)
flag_person$err_adminunit2 <- if_else(grepl("^高中部日間部$", flag_person$adminunit2) & flag_person$organization_id == "121302", 0, flag_person$err_adminunit2)
#磐石高中才有"高中部"
flag_person$err_adminunit2 <- if_else(grepl("^高中部$", flag_person$adminunit2) & flag_person$organization_id == "181307", 0, flag_person$err_adminunit2)


#兼任行政職職稱（三）
flag_person$err_admintitle3 <- 1
flag_person$err_admintitle3 <- if_else(grepl("主任$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("秘書$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("組長$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("組員$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^副校長$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^均質化承辦人$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^均職化承辦人$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^優質化協辦人$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("校安$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("心理師$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("技術員$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("人員$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("職輔員$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("廚工$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("營養師$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("編制外行政人力$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("司機$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("秘書$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("祕書$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("舍監$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("辦事員$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^幹事$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("職務代理$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("職代$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("職務代理人$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("救生員$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("值機員$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("監督$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("三副$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("社工師$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("助理$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("專員$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("政風$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("1", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^N$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("輔導員$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^警衛$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^實習餐廳經理$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("管理師$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^清潔員$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("佐理$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("技佐$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^會計員$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("管理員$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^書記$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("佐理員$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("館員$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("科主席$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("護理師$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("助理員$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^助教$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("庶務$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("文書$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^電競專案教練$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^出納$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^助教$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^督導$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("訓育業務$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("指導員$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
#光禾華德福才可有以下職稱：總務、學務、教務、輔導、人事、國中部行政、高中部行政、會計
flag_person$err_admintitle3 <- if_else(grepl("^國中部行政$", flag_person$admintitle3) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^高中部行政$", flag_person$admintitle3) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^總務$", flag_person$admintitle3) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^學務$", flag_person$admintitle3) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^教務$", flag_person$admintitle3) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^輔導$", flag_person$admintitle3) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^人事$", flag_person$admintitle3) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^會計$", flag_person$admintitle3) & flag_person$organization_id == "121302", 0, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^校長$", flag_person$admintitle3), 1, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("教官$", flag_person$admintitle3), 1, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("教師$", flag_person$admintitle3), 1, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("運動教練$", flag_person$admintitle3), 1, flag_person$err_admintitle3)
flag_person$err_admintitle3 <- if_else(grepl("^董事長$", flag_person$admintitle3), 1, flag_person$err_admintitle3)
#校長已在flag15檢查
flag_person$err_admintitle3 <- if_else(grepl("^校長$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
#導師已在flag15檢查
flag_person$err_admintitle3 <- if_else(grepl("導師$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
#教師已在flag15檢查
flag_person$err_admintitle3 <- if_else(grepl("教師$", flag_person$admintitle3), 0, flag_person$err_admintitle3)
#教官已在flag15檢查
flag_person$err_admintitle3 <- if_else(grepl("教官$", flag_person$admintitle3), 0, flag_person$err_admintitle3)


#兼任行政職服務單位（三）
flag_person$err_adminunit3 <- if_else(grepl("^校長室$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^副校長室$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^秘書室$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^國小部", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^小學部$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^國中部$", flag_person$adminunit3) | grepl("^國民中學部$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^中學部$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^教官室$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^教務處", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^進修部", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^進修學校$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^資訊室$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^圖書館", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^圖書室$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^實習處", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^實習室$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^輔導室", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("輔導處", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^學生事務處", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^學務處", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^總務處", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^雙語部", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^軍訓室$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^會計室$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^主計室$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^資訊室$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^國際部$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^董事會$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("中心$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("研究發展處$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("人事室$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^招生處$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^公關室$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^農場經營$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^藝文中心", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("科$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("1", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("組$", flag_person$adminunit3) & (grepl("*處*", flag_person$adminunit3) | grepl("*中心*", flag_person$adminunit3) | grepl("*部*", flag_person$adminunit3) | grepl("*室*", flag_person$adminunit3) | grepl("*館*", flag_person$adminunit3)), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^N$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^國中部教務處$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^高中部教務處$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^圖資室$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^圖資處$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^國際教育處$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^研發處$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^資訊處$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^實輔處$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^家具設計發展處$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^國際處$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^研發室$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^圖資室兼技術交流處$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^教育推廣處$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^學輔處$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^綜合高中$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^技術交流處$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^補校$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^國際交流處$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^專案研究室$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("分校$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^創發處$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^補校教學組$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^補校教務組$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^補校訓育組$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^實驗研究組$", flag_person$adminunit3), 0, flag_person$err_adminunit3)
#立志高中才有"高國中部"
flag_person$err_adminunit3 <- if_else(grepl("^高國中部$", flag_person$adminunit3) & flag_person$organization_id == "551301", 0, flag_person$err_adminunit3)
#光禾華德福才可有以下服務單位：國中部日間部、高中部日間部
flag_person$err_adminunit3 <- if_else(grepl("^國中部日間部$", flag_person$adminunit3) & flag_person$organization_id == "121302", 0, flag_person$err_adminunit3)
flag_person$err_adminunit3 <- if_else(grepl("^高中部日間部$", flag_person$adminunit3) & flag_person$organization_id == "121302", 0, flag_person$err_adminunit3)
#磐石高中才有"高中部"
flag_person$err_adminunit3 <- if_else(grepl("^高中部$", flag_person$adminunit3) & flag_person$organization_id == "181307", 0, flag_person$err_adminunit3)

#教員資料表之各兼任行政職資料不完整或不正確：請依欄位說明確認並正確填列行政單位名稱，如為二級單位主管，請敘明一級與二級單位名稱。如教務處教學組，學生事務處生活輔導組。
#教員資料表之兼任行政職資料不完整或不正確：請依欄位說明確認並正確填列行政單位名稱，如為二級單位主管，請敘明一級與二級單位名稱。如教務處教學組，學生事務處生活輔導組。
#職員(工)資料表之各服務單位資料不完整或不正確：請依欄位說明確認並正確填列服務單位名稱，如為二級單位主管，請敘明一級與二級單位名稱。如總務處出納組，學生事務處生活輔導組。
#教員資料表及職員(工)資料表之(兼任)職稱及服務單位資料不完整或不正確：請依欄位說明確認並正確填列行政單位名稱，如為二級單位主管，請敘明一級與二級單位名稱。如教務處教學組，總務處出納組。
#上述職員(工)資料表中人員，若未再兼任或代理行政職務者，僅須填寫「職務名稱」與「服務單位」，且二級單位名稱請於「服務單位」所填一級單位名稱後面加註。
#請確認*員之職稱或服務身分別，若確為教師，請將資料填至教員資料表。

#（職員(工)資料表之服務單位資料不完整或不正確：請依欄位說明確認並正確填列行政單位名稱，如為二級單位，請敘明一級與二級單位名稱，如學務處體育組，總務處出納組。另請再確認資源班是否為行政單位名稱。）

err_title0 <- data.frame(admintitle0 = flag_person$admintitle0, 
                         adminunit0 = flag_person$adminunit0, 
                         organization_id = flag_person$organization_id)
flag_person$err_title0 <- if_else(duplicated(err_title0) & !err_title0$admintitle0 %in% "N" & err_title0$admintitle0 %in% "組長", 1, 0)

err_title1 <- data.frame(admintitle1 = flag_person$admintitle1, 
                         adminunit1 = flag_person$adminunit1, 
                         organization_id = flag_person$organization_id)
flag_person$err_title1 <- if_else(duplicated(err_title1) & !err_title1$admintitle1 %in% "N" & err_title1$admintitle1 %in% "組長", 1, 0)

err_title2 <- data.frame(admintitle2 = flag_person$admintitle2, 
                         adminunit2 = flag_person$adminunit2, 
                         organization_id = flag_person$organization_id)
flag_person$err_title2 <- if_else(duplicated(err_title2) & !err_title2$admintitle2 %in% "N" & err_title2$admintitle2 %in% "組長", 1, 0)

err_title3 <- data.frame(admintitle3 = flag_person$admintitle3, 
                         adminunit3 = flag_person$adminunit3, 
                         organization_id = flag_person$organization_id)
flag_person$err_title3 <- if_else(duplicated(err_title3) & !err_title3$admintitle3 %in% "N" & err_title3$admintitle3 %in% "組長", 1, 0)

flag_person$err_flag_62 <- flag_person$err_admintitle0 + flag_person$err_adminunit0 + flag_person$err_admintitle1 + flag_person$err_adminunit1 + flag_person$err_admintitle2 + flag_person$err_adminunit2 + flag_person$err_admintitle3 + flag_person$err_adminunit3 + flag_person$err_title0 + flag_person$err_title1 + flag_person$err_title2 + flag_person$err_title3

flag_person$err_flag <- if_else(flag_person$err_flag_62 != 0, 1, 0)

#加註
flag_person$name <- paste(flag_person$name, "（", sep = "")
flag_person$name <- if_else(flag_person$err_admintitle0 != 0, paste(flag_person$name, "職務名稱：", flag_person$admintitle0, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_adminunit0 != 0, paste(flag_person$name, "服務單位：", flag_person$adminunit0, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_admintitle1 != 0, paste(flag_person$name, "兼任行政職職稱(一)：", flag_person$admintitle1, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_adminunit1 != 0, paste(flag_person$name, "兼任行政職服務單位(一)：", flag_person$adminunit1, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_admintitle2 != 0, paste(flag_person$name, "兼任行政職職稱(二)：", flag_person$admintitle2, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_adminunit2 != 0, paste(flag_person$name, "兼任行政職服務單位(二)：", flag_person$adminunit2, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_admintitle3 != 0, paste(flag_person$name, "兼任行政職職稱(三)：", flag_person$admintitle3, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_adminunit3 != 0, paste(flag_person$name, "兼任行政職服務單位(三)：", flag_person$adminunit3, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_title0 != 0, paste(flag_person$name, "服務單位：", flag_person$adminunit0, " ", "職務名稱：", flag_person$admintitle0, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_title1 != 0, paste(flag_person$name, "兼任行政職服務單位(一)：", flag_person$adminunit1, " ", "兼任行政職職稱(一)：", flag_person$admintitle1, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_title2 != 0, paste(flag_person$name, "兼任行政職服務單位(二)：", flag_person$adminunit2, " ", "兼任行政職職稱(二)：", flag_person$admintitle2, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_title3 != 0, paste(flag_person$name, "兼任行政職服務單位(三)：", flag_person$adminunit3, " ", "兼任行政職職稱(三)：", flag_person$admintitle3, "；", sep = ""), flag_person$name)
flag_person$name <- paste(flag_person$name, "）", sep = "")
flag_person$name <- gsub("；）", replacement = "）", flag_person$name)
flag_person$name <- gsub("（）", replacement = "", flag_person$name)


flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag62 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag62)[3 : length(colnames(flag_person_wide_flag62))]
flag_person_wide_flag62$flag62_r <- NA
for (i in temp){
  flag_person_wide_flag62$flag62_r <- paste(flag_person_wide_flag62$flag62_r, flag_person_wide_flag62[[i]], sep = " ")
}
flag_person_wide_flag62$flag62_r <- gsub("NA ", replacement="", flag_person_wide_flag62$flag62_r)
flag_person_wide_flag62$flag62_r <- gsub(" NA", replacement="", flag_person_wide_flag62$flag62_r)

#產生檢誤報告文字
flag62_temp <- flag_person_wide_flag62 %>%
  group_by(organization_id) %>%
  mutate(flag62_txt = paste(source, "之行政職資料不完整或不正確：", flag62_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag62_txt)) %>%
  distinct(organization_id, flag62_txt)

#根據organization_id，展開成寬資料(wide)
flag62 <- flag62_temp %>%
  dcast(organization_id ~ flag62_txt, value.var = "flag62_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag62)[2 : length(colnames(flag62))]
flag62$flag62 <- NA
for (i in temp){
  flag62$flag62 <- paste(flag62$flag62, flag62[[i]], sep = "； ")
}
flag62$flag62 <- gsub("NA； ", replacement="", flag62$flag62)
flag62$flag62 <- gsub("； NA", replacement="", flag62$flag62)

#產生檢誤報告文字
flag62 <- flag62 %>%
  subset(select = c(organization_id, flag62)) %>%
  distinct(organization_id, flag62) %>%
  mutate(flag62 = paste(flag62, "（請再確認該員職務正確名稱）", sep = ""))
}else{
#偵測flag62是否存在。若不存在，則產生NA行
if('flag62' %in% ls()){
  print("flag62")
}else{
  flag62 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag62$flag62 <- ""
}
}
# flag64: 本校任職需扣除年資非0000者分布偏高。 -------------------------------------------------------------------
flag_person <- drev_person_1

flag_person$dese <- 0
flag_person$dese <- if_else(flag_person$desedym != "0000", 1, flag_person$dese)


flag_person$jj <- 1

flag_person_wide_flag64 <- aggregate(cbind(dese, jj) ~ organization_id, flag_person, sum)

flag_person_wide_flag64$flag_err <- 0
flag_person_wide_flag64$err_flag_txt <- if_else(flag_person_wide_flag64$dese / flag_person_wide_flag64$jj > 0.25, "扣除年資不為零的人數偏高，請再依欄位說明確認。", "")

if (dim(flag_person_wide_flag64 %>% subset(err_flag_txt != ""))[1] != 0){
#根據organization_id，展開成寬資料(wide)
flag64 <- flag_person_wide_flag64 %>%
  subset(err_flag_txt != "") %>%
  dcast(organization_id ~ err_flag_txt, value.var = "err_flag_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag64)[2 : length(colnames(flag64))]
flag64$flag64 <- NA
for (i in temp){
  flag64$flag64 <- paste(flag64$flag64, flag64[[i]], sep = "； ")
}
flag64$flag64 <- gsub("NA； ", replacement="", flag64$flag64)
flag64$flag64 <- gsub("； NA", replacement="", flag64$flag64)

#產生檢誤報告文字
flag64 <- flag64 %>%
  subset(select = c(organization_id, flag64)) %>%
  distinct(organization_id, flag64)
}else{
#偵測flag64是否存在。若不存在，則產生NA行
if('flag64' %in% ls()){
  print("flag64")
}else{
  flag64 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag64$flag64 <- ""
}
}
# flag80: 代理教師、兼任教師、鐘點教師、長期代課教師、約用教師、約聘僱教師的「本校到職日期」非屬本學期，請再確認。-------------------------------------------------------------------
flag_person <- drev_person_1

#代理教師、兼任教師、長期代課教師、鐘點教師、約聘僱教師、約用教師到職日過早

#兼任教師、長期代課教師、鐘點教師、約聘僱教師、約用教師到職日應為上一個學期開學日(每學期(年)需修改emp_year、emp_mon的時間)
flag_person$emp_year1 <- 111
flag_person$emp_mon1 <- 2

flag_person$arvy1 <- substr(flag_person$onbodat, 1, 3) %>% as.numeric()
flag_person$arvm1 <- substr(flag_person$onbodat, 4, 5) %>% as.numeric()

flag_person$err_emp1 <- if_else(((flag_person$emptype == "兼任" | flag_person$emptype == "長期代課" | flag_person$emptype == "鐘點教師" | flag_person$emptype == "約聘僱" | flag_person$emptype == "約用") & flag_person$sertype == "教師") & (flag_person$arvy1 * 12 + flag_person$arvm1) < (flag_person$emp_year1 * 12 + flag_person$emp_mon1), 1, 0)

#代理教師到職日應為上一個學期開學日-2年，依法規代理教師得續聘2次(每學期(年)需修改emp_year、emp_mon的時間)
flag_person$emp_year2 <- flag_person$emp_year1 - 2
flag_person$emp_mon2 <- flag_person$emp_mon1

flag_person$arvy2 <- substr(flag_person$onbodat, 1, 3) %>% as.numeric()
flag_person$arvm2 <- substr(flag_person$onbodat, 4, 5) %>% as.numeric()

flag_person$err_emp2 <- if_else((flag_person$emptype == "代理" & flag_person$sertype == "教師") & (flag_person$arvy2 * 12 + flag_person$arvm2) < (flag_person$emp_year2 * 12 + flag_person$emp_mon2), 1, 0)

flag_person$emptypesertype <- paste(flag_person$emptype, flag_person$sertype, sep = "")
flag_person$emptypesertype <- if_else(flag_person$emptypesertype == "鐘點教師教師", "鐘點教師", flag_person$emptypesertype)

flag_person$err_flag <- if_else(flag_person$err_emp1 == 1 | flag_person$err_emp2 == 1, 1, 0)

#加註
flag_person$name <- paste(flag_person$name, "（", flag_person$emptypesertype, " 到職日:", flag_person$onbodat, "）", sep = "")
flag_person$name <- gsub("；）", replacement = "）", flag_person$name)
flag_person$name <- gsub("（）", replacement = "", flag_person$name)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag80 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag80)[3 : length(colnames(flag_person_wide_flag80))]
flag_person_wide_flag80$flag80_r <- NA
for (i in temp){
  flag_person_wide_flag80$flag80_r <- paste(flag_person_wide_flag80$flag80_r, flag_person_wide_flag80[[i]], sep = " ")
}
flag_person_wide_flag80$flag80_r <- gsub("NA ", replacement="", flag_person_wide_flag80$flag80_r)
flag_person_wide_flag80$flag80_r <- gsub(" NA", replacement="", flag_person_wide_flag80$flag80_r)

#產生檢誤報告文字
flag80_temp <- flag_person_wide_flag80 %>%
  group_by(organization_id) %>%
  mutate(flag80_txt = paste(source, "需核對「本校到職日期」：",flag80_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag80_txt)) %>%
  distinct(organization_id, flag80_txt)

#根據organization_id，展開成寬資料(wide)
flag80 <- flag80_temp %>%
  dcast(organization_id ~ flag80_txt, value.var = "flag80_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag80)[2 : length(colnames(flag80))]
flag80$flag80 <- NA
for (i in temp){
  flag80$flag80 <- paste(flag80$flag80, flag80[[i]], sep = "； ")
}
flag80$flag80 <- gsub("NA； ", replacement="", flag80$flag80)
flag80$flag80 <- gsub("； NA", replacement="", flag80$flag80)

#產生檢誤報告文字
flag80 <- flag80 %>%
  subset(select = c(organization_id, flag80)) %>%
  distinct(organization_id, flag80) %>%
  mutate(flag80 = paste(flag80, "", sep = ""))
}else{
#偵測flag80是否存在。若不存在，則產生NA行
if('flag80' %in% ls()){
  print("flag80")
}else{
  flag80 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag80$flag80 <- ""
}
}
# flag82: 若請假類別填寫「留職停薪」，則留職停薪原因須填寫內容。 -------------------------------------------------------------------
flag_person <- drev_person_1

#抓出:請假類別填"留職停薪"，但留職停薪原因填N
flag_person$err_flag <- 0
flag_person$err_flag <- if_else(grepl("留職停薪", flag_person$leave) & flag_person$levpay == "N", 1, flag_person$err_flag)
flag_person$err_flag <- if_else(grepl("留職停薪", flag_person$leave) & flag_person$levpay == "Ｎ", 1, flag_person$err_flag)
flag_person$err_flag <- if_else(grepl("留停", flag_person$leave) & flag_person$levpay == "N", 1, flag_person$err_flag)
flag_person$err_flag <- if_else(grepl("留停", flag_person$leave) & flag_person$levpay == "Ｎ", 1, flag_person$err_flag)

#加註
flag_person$name <- paste(flag_person$name, "（", "請假類別：", flag_person$leave, "；留職停薪原因：", flag_person$levpay, "）", sep = "")
flag_person$name <- gsub("；）", replacement = "）", flag_person$name)
flag_person$name <- gsub("（）", replacement = "", flag_person$name)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag82 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag82)[3 : length(colnames(flag_person_wide_flag82))]
flag_person_wide_flag82$flag82_r <- NA
for (i in temp){
  flag_person_wide_flag82$flag82_r <- paste(flag_person_wide_flag82$flag82_r, flag_person_wide_flag82[[i]], sep = " ")
}
flag_person_wide_flag82$flag82_r <- gsub("NA ", replacement="", flag_person_wide_flag82$flag82_r)
flag_person_wide_flag82$flag82_r <- gsub(" NA", replacement="", flag_person_wide_flag82$flag82_r)

#產生檢誤報告文字
flag82_temp <- flag_person_wide_flag82 %>%
  group_by(organization_id) %>%
  mutate(flag82_txt = paste(source, "需修改請假類別、留職停薪原因：", flag82_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag82_txt)) %>%
  distinct(organization_id, flag82_txt)

#根據organization_id，展開成寬資料(wide)
flag82 <- flag82_temp %>%
  dcast(organization_id ~ flag82_txt, value.var = "flag82_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag82)[2 : length(colnames(flag82))]
flag82$flag82 <- NA
for (i in temp){
  flag82$flag82 <- paste(flag82$flag82, flag82[[i]], sep = "； ")
}
flag82$flag82 <- gsub("NA； ", replacement="", flag82$flag82)
flag82$flag82 <- gsub("； NA", replacement="", flag82$flag82)

#產生檢誤報告文字
flag82 <- flag82 %>%
  subset(select = c(organization_id, flag82)) %>%
  distinct(organization_id, flag82) %>%
  mutate(flag82 = paste(flag82, "", sep = ""))
}else{
#偵測flag82是否存在。若不存在，則產生NA行
if('flag82' %in% ls()){
  print("flag82")
}else{
  flag82 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag82$flag82 <- ""
}
}
# flag89: 專任教師、代理教師原則須具大專以上學歷，請再確認實際情況及所填資料。 -------------------------------------------------------------------
flag_person <- drev_person_1

#專任、代理教師最高學歷是否為大專以上不應為N
flag_person$err_flag <- 0
flag_person$err_flag <- if_else(flag_person$degree == "N" & flag_person$emptype == "專任" & flag_person$sertype == "教師", 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$degree == "N" & flag_person$emptype == "代理" & flag_person$sertype == "教師", 1, flag_person$err_flag)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag89 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag89)[3 : length(colnames(flag_person_wide_flag89))]
flag_person_wide_flag89$flag89_r <- NA
for (i in temp){
  flag_person_wide_flag89$flag89_r <- paste(flag_person_wide_flag89$flag89_r, flag_person_wide_flag89[[i]], sep = " ")
}
flag_person_wide_flag89$flag89_r <- gsub("NA ", replacement="", flag_person_wide_flag89$flag89_r)
flag_person_wide_flag89$flag89_r <- gsub(" NA", replacement="", flag_person_wide_flag89$flag89_r)

#產生檢誤報告文字
flag89_temp <- flag_person_wide_flag89 %>%
  group_by(organization_id) %>%
  mutate(flag89_txt = paste(source, "：", flag89_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag89_txt)) %>%
  distinct(organization_id, flag89_txt)

#根據organization_id，展開成寬資料(wide)
flag89 <- flag89_temp %>%
  dcast(organization_id ~ flag89_txt, value.var = "flag89_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag89)[2 : length(colnames(flag89))]
flag89$flag89 <- NA
for (i in temp){
  flag89$flag89 <- paste(flag89$flag89, flag89[[i]], sep = "； ")
}
flag89$flag89 <- gsub("NA； ", replacement="", flag89$flag89)
flag89$flag89 <- gsub("； NA", replacement="", flag89$flag89)

#產生檢誤報告文字
flag89 <- flag89 %>%
  subset(select = c(organization_id, flag89)) %>%
  distinct(organization_id, flag89) %>%
  mutate(flag89 = paste(flag89, "（請確認渠等人員畢業學歷）", sep = ""))
}else{
#偵測flag89是否存在。若不存在，則產生NA行
if('flag89' %in% ls()){
  print("flag89")
}else{
  flag89 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag89$flag89 <- ""
}
}
# flag90: 校內行政職務，原則由專任或代理教師兼任，請再確認實際情況及所填資料。 -------------------------------------------------------------------
flag_person <- drev_person_1

#兼任、長期代課、專職族語老師、鐘點教師、約聘僱、約用教師，不應有兼任行政職務，也不可擔任導師
flag_person$err_flag <- 0
flag_person$err_flag <- if_else(flag_person$emptype == "兼任" & flag_person$sertype == "教師" & flag_person$admintitle1 != "N" , 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$emptype == "長期代課" & flag_person$sertype == "教師" & flag_person$admintitle1 != "N" , 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$emptype == "專職族語老師" & flag_person$sertype == "教師" & flag_person$admintitle1 != "N" , 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$emptype == "鐘點教師" & flag_person$sertype == "教師" & flag_person$admintitle1 != "N" , 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$emptype == "約聘僱" & flag_person$sertype == "教師" & flag_person$admintitle1 != "N" , 1, flag_person$err_flag)
flag_person$err_flag <- if_else(flag_person$emptype == "約用" & flag_person$sertype == "教師" & flag_person$admintitle1 != "N" , 1, flag_person$err_flag)

#加註
flag_person$name <- paste(flag_person$name, "（", flag_person$emptype, "）", "（人事資料顯示該教師兼任行政職務；）", sep = "")
flag_person$name <- gsub("；）", replacement = "）", flag_person$name)
flag_person$name <- gsub("（）", replacement = "", flag_person$name)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag90 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag90)[3 : length(colnames(flag_person_wide_flag90))]
flag_person_wide_flag90$flag90_r <- NA
for (i in temp){
  flag_person_wide_flag90$flag90_r <- paste(flag_person_wide_flag90$flag90_r, flag_person_wide_flag90[[i]], sep = " ")
}
flag_person_wide_flag90$flag90_r <- gsub("NA ", replacement="", flag_person_wide_flag90$flag90_r)
flag_person_wide_flag90$flag90_r <- gsub(" NA", replacement="", flag_person_wide_flag90$flag90_r)

#產生檢誤報告文字
flag90_temp <- flag_person_wide_flag90 %>%
  group_by(organization_id) %>%
  mutate(flag90_txt = paste("姓名：", flag90_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag90_txt)) %>%
  distinct(organization_id, flag90_txt)

#根據organization_id，展開成寬資料(wide)
flag90 <- flag90_temp %>%
  dcast(organization_id ~ flag90_txt, value.var = "flag90_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag90)[2 : length(colnames(flag90))]
flag90$flag90 <- NA
for (i in temp){
  flag90$flag90 <- paste(flag90$flag90, flag90[[i]], sep = "； ")
}
flag90$flag90 <- gsub("NA； ", replacement="", flag90$flag90)
flag90$flag90 <- gsub("； NA", replacement="", flag90$flag90)

#產生檢誤報告文字
flag90 <- flag90 %>%
  subset(select = c(organization_id, flag90)) %>%
  distinct(organization_id, flag90) %>%
  mutate(flag90 = paste(flag90, "（校內行政職務原則由專任或代理教師兼任，請協助再確認上述教師是否兼任行政職，或協助再確認上述教師之聘任類別）", sep = ""))
}else{
#偵測flag90是否存在。若不存在，則產生NA行
if('flag90' %in% ls()){
  print("flag90")
}else{
  flag90 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag90$flag90 <- ""
}
}
# flag94: 職員（工）的「職務名稱」與「聘任類別」不相符應。 -------------------------------------------------------------------
flag_person <- drev_person_1

#職員工若為專任，職務名稱不可出現"約僱"、"約聘雇"、"約雇"、"約聘雇"、"約聘"之關鍵字
flag_person$err_flag <- 0
flag_person$err_flag <- if_else(grepl("約僱", flag_person$admintitle0) & flag_person$emptype == "專任" & flag_person$source == "職員工資料表", 1, flag_person$err_flag)
flag_person$err_flag <- if_else(grepl("約聘僱", flag_person$admintitle0) & flag_person$emptype == "專任" & flag_person$source == "職員工資料表", 1, flag_person$err_flag)
flag_person$err_flag <- if_else(grepl("約雇", flag_person$admintitle0) & flag_person$emptype == "專任" & flag_person$source == "職員工資料表", 1, flag_person$err_flag)
flag_person$err_flag <- if_else(grepl("約聘雇", flag_person$admintitle0) & flag_person$emptype == "專任" & flag_person$source == "職員工資料表", 1, flag_person$err_flag)
flag_person$err_flag <- if_else(grepl("約聘", flag_person$admintitle0) & flag_person$emptype == "專任" & flag_person$source == "職員工資料表", 1, flag_person$err_flag)

#加註
flag_person$name <- paste(flag_person$name, "（職務名稱：", flag_person$admintitle0, "；）", sep = "")
flag_person$name <- gsub("；）", replacement = "）", flag_person$name)
flag_person$name <- gsub("（）", replacement = "", flag_person$name)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag94 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag94)[3 : length(colnames(flag_person_wide_flag94))]
flag_person_wide_flag94$flag94_r <- NA
for (i in temp){
  flag_person_wide_flag94$flag94_r <- paste(flag_person_wide_flag94$flag94_r, flag_person_wide_flag94[[i]], sep = " ")
}
flag_person_wide_flag94$flag94_r <- gsub("NA ", replacement="", flag_person_wide_flag94$flag94_r)
flag_person_wide_flag94$flag94_r <- gsub(" NA", replacement="", flag_person_wide_flag94$flag94_r)

#產生檢誤報告文字
flag94_temp <- flag_person_wide_flag94 %>%
  group_by(organization_id) %>%
  mutate(flag94_txt = paste(source, "姓名：", flag94_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag94_txt)) %>%
  distinct(organization_id, flag94_txt)

#根據organization_id，展開成寬資料(wide)
flag94 <- flag94_temp %>%
  dcast(organization_id ~ flag94_txt, value.var = "flag94_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag94)[2 : length(colnames(flag94))]
flag94$flag94 <- NA
for (i in temp){
  flag94$flag94 <- paste(flag94$flag94, flag94[[i]], sep = "； ")
}
flag94$flag94 <- gsub("NA； ", replacement="", flag94$flag94)
flag94$flag94 <- gsub("； NA", replacement="", flag94$flag94)

#產生檢誤報告文字
flag94 <- flag94 %>%
  subset(select = c(organization_id, flag94)) %>%
  distinct(organization_id, flag94) %>%
  mutate(flag94 = paste(flag94, "（請確認上開職員(工)之『聘任類別』及『職務名稱』。凡以簽訂契約方式任用之人員，無論是否為編制內員額，其『聘任類別』皆請修正為『約聘僱』。並請再協助確認上開職員(工)『職務名稱』是否正確。惟貴校職員(工)如具正式公務人員身分者，則其『聘任類別』原則應是『專任』。）（貴校如僅有本項檢查須再確認修正資料，則不列入國教署催辦範圍，惟請儘速確認修正送出資料。）", sep = ""))
}else{
#偵測flag94是否存在。若不存在，則產生NA行
if('flag94' %in% ls()){
  print("flag94")
}else{
  flag94 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag94$flag94 <- ""
}
}
# spe3: 本校到職日期晚於填報基準日。 -------------------------------------------------------------------
flag_person <- drev_person_1

#本校到職日期晚於填報基準日。
flag_person$survey_year <- 111
flag_person$survey_mon <- 9

flag_person$arvy <- substr(flag_person$onbodat, 1, 3) %>% as.numeric()
flag_person$arvm <- substr(flag_person$onbodat, 4, 5) %>% as.numeric()

flag_person$err_spe <- if_else((flag_person$arvy * 12 + flag_person$arvm) > (flag_person$survey_year * 12 + flag_person$survey_mon), 1, 0)

#加註
flag_person$name <- paste(flag_person$name, "（", flag_person$onbodat, "）", sep = "")
flag_person$name <- gsub("；）", replacement = "）", flag_person$name)
flag_person$name <- gsub("（）", replacement = "", flag_person$name)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_spe == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_spe == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_spe3 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_spe)) %>%
  subset(err_spe == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_spe3)[3 : length(colnames(flag_person_wide_spe3))]
flag_person_wide_spe3$spe3_r <- NA
for (i in temp){
  flag_person_wide_spe3$spe3_r <- paste(flag_person_wide_spe3$spe3_r, flag_person_wide_spe3[[i]], sep = " ")
}
flag_person_wide_spe3$spe3_r <- gsub("NA ", replacement="", flag_person_wide_spe3$spe3_r)
flag_person_wide_spe3$spe3_r <- gsub(" NA", replacement="", flag_person_wide_spe3$spe3_r)

#產生檢誤報告文字
spe3_temp <- flag_person_wide_spe3 %>%
  group_by(organization_id) %>%
  mutate(spe3_txt = paste(source, "：", spe3_r, sep = ""), "") %>%
  subset(select = c(organization_id, spe3_txt)) %>%
  distinct(organization_id, spe3_txt)

#根據organization_id，展開成寬資料(wide)
spe3 <- spe3_temp %>%
  dcast(organization_id ~ spe3_txt, value.var = "spe3_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(spe3)[2 : length(colnames(spe3))]
spe3$spe3 <- NA
for (i in temp){
  spe3$spe3 <- paste(spe3$spe3, spe3[[i]], sep = "； ")
}
spe3$spe3 <- gsub("NA； ", replacement="", spe3$spe3)
spe3$spe3 <- gsub("； NA", replacement="", spe3$spe3)

#產生檢誤報告文字
spe3 <- spe3 %>%
  subset(select = c(organization_id, spe3)) %>%
  distinct(organization_id, spe3) %>%
  mutate(spe3 = paste(spe3, "（請確認到職日期，並請以資料基準日111年9月30日當時情況為準）", sep = ""))
}else{
#偵測spe3是否存在。若不存在，則產生NA行
if('spe3' %in% ls()){
  print("spe3")
}else{
  spe3 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  spe3$spe3 <- ""
}
}
# spe5: 教職員工畢業學校若為(科技)大學或(技術)學院，學歷資訊原則於「學士」、「碩士」或「博士」學歷欄位填列，而非「副學士」。 -------------------------------------------------------------------
flag_person <- drev_person_1

#副學士學位畢業學校名稱不可出現"大學"及"學院"，可出現"專科學校"
flag_person$err_flag_adegreeu1 <- 0
flag_person$err_flag_adegreeu2 <- 0
flag_person$err_flag_adegreeu1 <- if_else(grepl("大學", flag_person$adegreeu1), 1, flag_person$err_flag_adegreeu1)
flag_person$err_flag_adegreeu1 <- if_else(grepl("學院", flag_person$adegreeu1), 1, flag_person$err_flag_adegreeu1)
flag_person$err_flag_adegreeu2 <- if_else(grepl("大學", flag_person$adegreeu2), 1, flag_person$err_flag_adegreeu2)
flag_person$err_flag_adegreeu2 <- if_else(grepl("學院", flag_person$adegreeu2), 1, flag_person$err_flag_adegreeu2)
#達姆施塔特工業大學（德語：Technische Universität Darmstadt），是德國歷史悠久的理工大學
flag_person$err_flag_adegreeu2 <- if_else(grepl("Darmstadt$", flag_person$adegreeu2), 1, flag_person$err_flag_adegreeu2)
#副學士的情況
flag_person$err_flag_adegreeu1 <- if_else(grepl("專科", flag_person$adegreeu1), 0, flag_person$err_flag_adegreeu1)
flag_person$err_flag_adegreeu1 <- if_else(grepl("二專", flag_person$adegreeu1), 0, flag_person$err_flag_adegreeu1)
flag_person$err_flag_adegreeu1 <- if_else(grepl("五專", flag_person$adegreeu1), 0, flag_person$err_flag_adegreeu1)
flag_person$err_flag_adegreeu1 <- if_else(grepl("五年制", flag_person$adegreeu1), 0, flag_person$err_flag_adegreeu1)
flag_person$err_flag_adegreeu1 <- if_else(grepl("商專", flag_person$adegreeu1), 0, flag_person$err_flag_adegreeu1)
flag_person$err_flag_adegreeu1 <- if_else(grepl("農專", flag_person$adegreeu1), 0, flag_person$err_flag_adegreeu1)
flag_person$err_flag_adegreeu1 <- if_else(grepl("空專", flag_person$adegreeu1), 0, flag_person$err_flag_adegreeu1)
flag_person$err_flag_adegreeu1 <- if_else(grepl("三專", flag_person$adegreeu1), 0, flag_person$err_flag_adegreeu1)
flag_person$err_flag_adegreeu2 <- if_else(grepl("專科", flag_person$adegreeu2), 0, flag_person$err_flag_adegreeu2)
flag_person$err_flag_adegreeu2 <- if_else(grepl("二專", flag_person$adegreeu2), 0, flag_person$err_flag_adegreeu2)
flag_person$err_flag_adegreeu2 <- if_else(grepl("五專", flag_person$adegreeu2), 0, flag_person$err_flag_adegreeu2)
flag_person$err_flag_adegreeu2 <- if_else(grepl("五年制", flag_person$adegreeu2), 0, flag_person$err_flag_adegreeu2)
flag_person$err_flag_adegreeu2 <- if_else(grepl("商專", flag_person$adegreeu2), 0, flag_person$err_flag_adegreeu2)
flag_person$err_flag_adegreeu2 <- if_else(grepl("農專", flag_person$adegreeu2), 0, flag_person$err_flag_adegreeu2)
flag_person$err_flag_adegreeu2 <- if_else(grepl("空專", flag_person$adegreeu2), 0, flag_person$err_flag_adegreeu2)
flag_person$err_flag_adegreeu2 <- if_else(grepl("三專", flag_person$adegreeu2), 0, flag_person$err_flag_adegreeu2)

flag_person$err_flag <- flag_person$err_flag_adegreeu1 + flag_person$err_flag_adegreeu2

#加註學士學位畢業學校名稱
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag_adegreeu1 == 1 ~ paste(flag_person$name, "（副學士學位畢業學校（一）：", flag_person$adegreeu1, "）", sep = ""),
  flag_person$err_flag_adegreeu2 == 1 ~ paste(flag_person$name, "（副學士學位畢業學校（二）：", flag_person$adegreeu2, "）", sep = ""),
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_spe5 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_spe5)[3 : length(colnames(flag_person_wide_spe5))]
flag_person_wide_spe5$spe5_r <- NA
for (i in temp){
  flag_person_wide_spe5$spe5_r <- paste(flag_person_wide_spe5$spe5_r, flag_person_wide_spe5[[i]], sep = " ")
}
flag_person_wide_spe5$spe5_r <- gsub("NA ", replacement="", flag_person_wide_spe5$spe5_r)
flag_person_wide_spe5$spe5_r <- gsub(" NA", replacement="", flag_person_wide_spe5$spe5_r)

#產生檢誤報告文字
spe5_temp <- flag_person_wide_spe5 %>%
  group_by(organization_id) %>%
  mutate(spe5_txt = paste(source, "：", spe5_r, sep = ""), "") %>%
  subset(select = c(organization_id, spe5_txt)) %>%
  distinct(organization_id, spe5_txt)

#根據organization_id，展開成寬資料(wide)
spe5 <- spe5_temp %>%
  dcast(organization_id ~ spe5_txt, value.var = "spe5_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(spe5)[2 : length(colnames(spe5))]
spe5$spe5 <- NA
for (i in temp){
  spe5$spe5 <- paste(spe5$spe5, spe5[[i]], sep = "； ")
}
spe5$spe5 <- gsub("NA； ", replacement="", spe5$spe5)
spe5$spe5 <- gsub("； NA", replacement="", spe5$spe5)

#產生檢誤報告文字
spe5 <- spe5 %>%
  subset(select = c(organization_id, spe5)) %>%
  distinct(organization_id, spe5)
}else{
#偵測spe5是否存在。若不存在，則產生NA行
if('spe5' %in% ls()){
  print("spe5")
}else{
  spe5 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  spe5$spe5 <- ""
}
}
# spe6: 各教育階段學歷資料內容是否完整正確。-------------------------------------------------------------------
# 例如：
# 1.	各學歷階段「國別」非填入「本國」或者外交部網站之世界各國名稱一覽表的國家名稱（或者至少須足以辨識國家）。
# 2.	各學歷階段「學校」填入非學校名稱。
# 3.	各學歷階段「系所」填入非系所名稱。
# 4.	需有專科學歷，才能報考碩士研究所（若為逕讀碩士，副學士不得為N）。
# 5.	需有學士學歷，才能報考博士研究所（若為逕讀博士，學士不得為N或填逕讀博士）。
# 6.	學士學位欄位若填列「逕讀碩士」，應填列碩士學位（不應為N）。
# 7.	碩士學位欄位若填列「逕讀博士」，應填列博士學位（不應為N）。

flag_person <- drev_person_1

#博士學位畢業學校國別（一）
flag_person$err_ddegreen1 <- 0
flag_person$err_ddegreen1 <- if_else(grepl("博士", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("碩士", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("學士", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("副學士", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("大學", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("分校", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("學院", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("科大", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("學校", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("官校", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("預校", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("書院", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("專科", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("藝專", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("海專", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("工專", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("護專", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("家專", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("商專", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("行專", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("農專", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("體專", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("藥專", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("師專", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("醫專", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("語專", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("university", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("University", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("UNIVERSITY", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("college", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("College", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("COLLEGE", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("系", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("所", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("班$", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("不分科系", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("不分系", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("department", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("Department", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("DEPARTMENT", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("兼課", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("最高學歷", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("^Y$", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("^待查詢$", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("^無$", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("^外國$", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("^國立$", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("^歐洲$", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("^美洲$", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("^亞洲$", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("^非洲$", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("^大洋洲$", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("肄業", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)
flag_person$err_ddegreen1 <- if_else(grepl("學分班", flag_person$ddegreen1), 1, flag_person$err_ddegreen1)

#博士學位畢業學校（一）
flag_person$err_ddegreeu1 <- 1
flag_person$err_ddegreeu1 <- if_else(grepl("大學", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("分校", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("學院", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("師大", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("科大", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("學校", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("官校", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("預校", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("書院", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("體院", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("專科", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("藝專", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("海專", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("工專", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("護專", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("家專", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("商專", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("行專", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("農專", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("體專", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("師專", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("藥專", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("醫專", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("語專", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("士校", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("專校$", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("university", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("University", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("UNIVERSITY", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("Uni$", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("college", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("College", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("COLLEGE", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("Universidad", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("UNIVERSIDAD", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("Conservatory", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("CRD", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("ENM", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("CRC", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("EMMA", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("CRR", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("CNR", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("TheNewSchool", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("Hochschule", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("BergenSchoolofArchitecture", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("Universität", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("Institute$", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("StellenboschUni$", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("^TUDarmstadt$", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("^N$", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("^莫斯科柴可夫斯基$", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("音樂院$", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("音樂研究所$", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("大学$", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("研究所博士班$", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("^中興法商$", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("^待查詢$", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("本國", flag_person$ddegreeu1), 1, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("^日本國立岡山大學$", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("肄業", flag_person$ddegreeu1), 1, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("學分班", flag_person$ddegreeu1), 1, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("籌備處$", flag_person$ddegreeu1), 1, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("Academy", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("academy", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)
flag_person$err_ddegreeu1 <- if_else(grepl("ACADEMY", flag_person$ddegreeu1), 0, flag_person$err_ddegreeu1)

#博士學位畢業系所（一）
flag_person$err_ddegreeg1 <- 0
flag_person$err_ddegreeg1 <- if_else(grepl("^博士$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^碩士$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^學士$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^副學士$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^大學$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^分校$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^學院$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^科大$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^學校$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^官校$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^預校$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^書院$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^專科$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^藝專$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^海專$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^工專$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^護專$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^家專$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^商專$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^行專$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^農專$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^體專$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^藥專$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^師專$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^醫專$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^語專$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^university$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^University$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^UNIVERSITY$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^college$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^College$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^COLLEGE$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("兼課", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("最高學歷", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^Y$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^待查詢$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("^無$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("肄業", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("學分班", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(grepl("行政$", flag_person$ddegreeg1), 1, flag_person$err_ddegreeg1)

#博士學位畢業學校國別（二）
flag_person$err_ddegreen2 <- 0
flag_person$err_ddegreen2 <- if_else(grepl("博士", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("碩士", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("學士", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("副學士", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("大學", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("分校", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("學院", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("科大", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("學校", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("官校", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("預校", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("書院", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("專科", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("藝專", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("海專", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("工專", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("護專", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("家專", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("商專", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("行專", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("農專", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("體專", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("藥專", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("師專", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("醫專", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("語專", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("university", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("University", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("UNIVERSITY", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("college", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("College", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("COLLEGE", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("系", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("所", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("班$", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("不分科系", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("不分系", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("department", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("Department", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("DEPARTMENT", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("兼課", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("最高學歷", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("^Y$", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("^待查詢$", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("^無$", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("^外國$", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("^國立$", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("^歐洲$", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("^美洲$", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("^亞洲$", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("^非洲$", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("^大洋洲$", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("肄業", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)
flag_person$err_ddegreen2 <- if_else(grepl("學分班", flag_person$ddegreen2), 1, flag_person$err_ddegreen2)

#博士學位畢業學校（二）
flag_person$err_ddegreeu2 <- 1
flag_person$err_ddegreeu2 <- if_else(grepl("大學", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("分校", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("學院", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("師大", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("科大", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("學校", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("官校", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("預校", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("書院", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("體院", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("專科", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("藝專", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("海專", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("工專", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("護專", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("家專", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("商專", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("行專", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("農專", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("體專", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("師專", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("藥專", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("醫專", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("語專", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("士校", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("專校$", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("university", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("University", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("UNIVERSITY", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("Uni$", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("college", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("College", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("COLLEGE", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("Universidad", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("UNIVERSIDAD", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("Conservatory", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("CRD", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("ENM", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("CRC", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("EMMA", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("CRR", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("CNR", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("TheNewSchool", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("Hochschule", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("BergenSchoolofArchitecture", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("Universität", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("Institute$", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("StellenboschUni$", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("^TUDarmstadt$", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("^N$", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("^莫斯科柴可夫斯基$", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("音樂院$", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("音樂研究所$", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("大学$", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("研究所博士班$", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("^中興法商$", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("^待查詢$", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("本國", flag_person$ddegreeu2), 1, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("^日本國立岡山大學$", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("肄業", flag_person$ddegreeu2), 1, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("學分班", flag_person$ddegreeu2), 1, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("籌備處$", flag_person$ddegreeu2), 1, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("Academy", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("academy", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)
flag_person$err_ddegreeu2 <- if_else(grepl("ACADEMY", flag_person$ddegreeu2), 0, flag_person$err_ddegreeu2)

#博士學位畢業系所（二）
flag_person$err_ddegreeg2 <- 0
flag_person$err_ddegreeg2 <- if_else(grepl("^博士$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^碩士$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^學士$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^副學士$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^大學$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^分校$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^學院$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^科大$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^學校$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^官校$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^預校$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^書院$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^專科$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^藝專$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^海專$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^工專$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^護專$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^家專$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^商專$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^行專$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^農專$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^體專$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^藥專$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^師專$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^醫專$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^語專$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^university$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^University$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^UNIVERSITY$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^college$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^College$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^COLLEGE$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("兼課", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("最高學歷", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^Y$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^待查詢$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("^無$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("肄業", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("學分班", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(grepl("行政$", flag_person$ddegreeg2), 1, flag_person$err_ddegreeg2)

#碩士學位畢業學校國別（一）
flag_person$err_mdegreen1 <- 0
flag_person$err_mdegreen1 <- if_else(grepl("博士", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("碩士", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("學士", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("副學士", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("大學", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("分校", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("學院", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("科大", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("學校", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("官校", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("預校", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("書院", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("專科", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("藝專", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("海專", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("工專", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("護專", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("家專", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("商專", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("行專", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("農專", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("體專", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("藥專", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("師專", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("醫專", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("語專", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("university", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("University", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("UNIVERSITY", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("college", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("College", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("COLLEGE", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("系", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("所", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("班$", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("不分科系", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("不分系", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("department", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("Department", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("DEPARTMENT", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("^逕讀", flag_person$mdegreen1), 0, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("^逕讀碩士$", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("^兼課", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("^最高學歷", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("^Y$", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("^待查詢$", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("^無$", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("^外國$", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("^國立$", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("^歐洲$", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("^亞洲$", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("^美洲$", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("^非洲$", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("^大洋洲$", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("肄業", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)
flag_person$err_mdegreen1 <- if_else(grepl("學分班", flag_person$mdegreen1), 1, flag_person$err_mdegreen1)

#碩士學位畢業學校學校（一）
flag_person$err_mdegreeu1 <- 1
flag_person$err_mdegreeu1 <- if_else(grepl("大學", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("分校", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("學院", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("師院", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("師大", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("科大", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("學校", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("官校", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("預校", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("書院", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("體院", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("專科", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("藝專", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("海專", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("工專", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("護專", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("家專", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("商專", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("行專", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("農專", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("體專", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("師專", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("藥專", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("醫專", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("語專", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("士校", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("專校$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("逕讀", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("音樂院$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("音樂研究所$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("university", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("University", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("UNIVERSITY", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("Uni$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("UNIVERSIT", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("college", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("College", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("COLLEGE", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("Universidad$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("UNIVERSIDAD$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("Conservatory$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("CRD$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("ENM$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("CRC$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("EMMA$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("CRR$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("CNR$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("TheNewSchool$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("Hochschule$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("BergenSchoolofArchitecture$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("Universität$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("Institute$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("StellenboschUni$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^TUDarmstadt$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^大学$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^研究所博士班$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^中興法商$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^N$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^待查詢$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^無$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^離職$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^因故$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^NavalPostgraduateSchool$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^BiblicalInterpretationLondonSchoolofTheology$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^DallasBaptistUniv$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^NewYorkFilmAcademy$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^ConservatorioStatalediMilano“GiuseppeVerdi”Italia$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^衛理神學研究院$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^KnowledgeSystemInstitute$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^ColumbiaBiblicalSeminaryandSchoolofMissions$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^UnitecInstituteofTechnology$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^巴拉圭高等戰略研究院$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^本國$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^肄業$", flag_person$mdegreeu1), 1, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^學分班$", flag_person$mdegreeu1), 1, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^籌備處$", flag_person$mdegreeu1), 1, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^高雄餐旅$", flag_person$mdegreeu1), 1, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^universitéJeanMoulinLyon3$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^UniversityCollegeLondon$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^BirminghamUiversity$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^ColumbiaUniversity哥倫比亞大學MathematicsEducation$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^StevensInstituteofTechnology，NJ，USA$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^UNITEDSTATESSPORTSACADEMY$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^KnowledgeSystemsInstitute$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("DomusAcademy", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("DOMUSACADEMY", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)
flag_person$err_mdegreeu1 <- if_else(grepl("^UniversitéStendhalGrenobleIII$", flag_person$mdegreeu1), 0, flag_person$err_mdegreeu1)

#碩士學位畢業系所（一）
flag_person$err_mdegreeg1 <- 0
flag_person$err_mdegreeg1 <- if_else(grepl("^博士$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^碩士$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^學士$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^副學士$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^大學$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^分校$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^學院$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^科大$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^學校$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^官校$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^預校$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^書院$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^專科$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^藝專$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^海專$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^工專$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^護專$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^家專$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^商專$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^行專$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^農專$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^體專$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^藥專$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^師專$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^醫專$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^語專$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^university$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^University$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^UNIVERSITY$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^college$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^College$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^COLLEGE$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("兼課", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("最高學歷", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^Y$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^待查詢$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^無$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("肄業", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("學分班", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("行政$", flag_person$mdegreeg1), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^教育政策與行政$", flag_person$mdegreeg1) & (grepl("國立臺灣師範大學", flag_person$mdegreeu1) | grepl("國立台灣師範大學", flag_person$mdegreeu1)), 0, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^社會教育學系學校圖書行政$", flag_person$mdegreeg1) & (grepl("國立臺灣師範大學", flag_person$mdegreeu1) | grepl("國立台灣師範大學", flag_person$mdegreeu1)), 0, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(grepl("^工業教育學系技職教育行政$", flag_person$mdegreeg1) & (grepl("國立臺灣師範大學", flag_person$mdegreeu1) | grepl("國立台灣師範大學", flag_person$mdegreeu1)), 0, flag_person$err_mdegreeg1)

#碩士學位畢業學校國別（二）
flag_person$err_mdegreen2 <- 0
flag_person$err_mdegreen2 <- if_else(grepl("博士", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("碩士", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("學士", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("副學士", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("大學", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("分校", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("學院", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("科大", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("學校", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("官校", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("預校", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("書院", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("專科", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("藝專", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("海專", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("工專", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("護專", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("家專", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("商專", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("行專", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("農專", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("體專", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("藥專", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("師專", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("醫專", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("語專", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("university", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("University", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("UNIVERSITY", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("college", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("College", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("COLLEGE", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("系", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("所", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("班$", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("不分科系", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("不分系", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("department", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("Department", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("DEPARTMENT", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("^逕讀", flag_person$mdegreen2), 0, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("^逕讀碩士$", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("^兼課", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("^最高學歷", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("^Y$", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("^待查詢$", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("^無$", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("^外國$", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("^國立$", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("^歐洲$", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("^亞洲$", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("^美洲$", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("^非洲$", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("^大洋洲$", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("肄業", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)
flag_person$err_mdegreen2 <- if_else(grepl("學分班", flag_person$mdegreen2), 1, flag_person$err_mdegreen2)

#碩士學位畢業學校學校（二）
flag_person$err_mdegreeu2 <- 1
flag_person$err_mdegreeu2 <- if_else(grepl("大學", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("分校", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("學院", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("師院", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("師大", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("科大", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("學校", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("官校", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("預校", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("書院", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("體院", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("專科", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("藝專", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("海專", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("工專", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("護專", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("家專", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("商專", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("行專", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("農專", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("體專", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("師專", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("藥專", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("醫專", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("語專", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("士校", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("專校$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("逕讀$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("音樂院$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("音樂研究所$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("university$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("University$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("UNIVERSITY$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("Uni$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("college", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("College", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("COLLEGE", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("Universidad$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("UNIVERSIDAD$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("Conservatory$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("CRD$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("ENM$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("CRC$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("EMMA$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("CRR$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("CNR$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("TheNewSchool$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("Hochschule$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("BergenSchoolofArchitecture$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("Universität$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("Institute$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("StellenboschUni$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^TUDarmstadt$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^大学$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^研究所博士班$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^中興法商$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^N$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^待查詢$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^無$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^離職$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^因故$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^NavalPostgraduateSchool$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^BiblicalInterpretationLondonSchoolofTheology$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^DallasBaptistUniv$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^NewYorkFilmAcademy$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^ConservatorioStatalediMilano“GiuseppeVerdi”Italia$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^衛理神學研究院$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^KnowledgeSystemInstitute$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^ColumbiaBiblicalSeminaryandSchoolofMissions$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^UnitecInstituteofTechnology$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^巴拉圭高等戰略研究院$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^本國$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^肄業$", flag_person$mdegreeu2), 1, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^學分班$", flag_person$mdegreeu2), 1, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^籌備處$", flag_person$mdegreeu2), 1, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^高雄餐旅$", flag_person$mdegreeu2), 1, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^universitéJeanMoulinLyon3$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^UniversityCollegeLondon$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^BirminghamUiversity$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^ColumbiaUniversity哥倫比亞大學MathematicsEducation$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^StevensInstituteofTechnology，NJ，USA$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^UNITEDSTATESSPORTSACADEMY$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^KnowledgeSystemsInstitute$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("DomusAcademy", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("DOMUSACADEMY", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)
flag_person$err_mdegreeu2 <- if_else(grepl("^UniversitéStendhalGrenobleIII$", flag_person$mdegreeu2), 0, flag_person$err_mdegreeu2)

#碩士學位畢業系所（二）
flag_person$err_mdegreeg2 <- 0
flag_person$err_mdegreeg2 <- if_else(grepl("^博士$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^碩士$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^學士$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^副學士$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^大學$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^分校$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^學院$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^科大$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^學校$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^官校$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^預校$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^書院$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^專科$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^藝專$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^海專$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^工專$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^護專$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^家專$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^商專$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^行專$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^農專$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^體專$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^藥專$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^師專$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^醫專$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^語專$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^university$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^University$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^UNIVERSITY$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^college$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^College$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^COLLEGE$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("兼課", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("最高學歷", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^Y$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^待查詢$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^無$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("肄業", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("學分班", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("行政$", flag_person$mdegreeg2), 1, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^教育政策與行政$", flag_person$mdegreeg2) & (grepl("國立臺灣師範大學", flag_person$mdegreeu2) | grepl("國立台灣師範大學", flag_person$mdegreeu2)), 0, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^社會教育學系學校圖書行政$", flag_person$mdegreeg2) & (grepl("國立臺灣師範大學", flag_person$mdegreeu2) | grepl("國立台灣師範大學", flag_person$mdegreeu2)), 0, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(grepl("^工業教育學系技職教育行政$", flag_person$mdegreeg2) & (grepl("國立臺灣師範大學", flag_person$mdegreeu2) | grepl("國立台灣師範大學", flag_person$mdegreeu2)), 0, flag_person$err_mdegreeg2)

#學士學位畢業學校國別（一）
flag_person$err_bdegreen1 <- 0
flag_person$err_bdegreen1 <- if_else(grepl("博士", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("碩士", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("學士", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("副學士", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("大學", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("分校", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("學院", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("科大", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("學校", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("官校", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("預校", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("書院", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("專科", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("藝專", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("海專", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("工專", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("護專", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("家專", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("商專", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("行專", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("農專", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("體專", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("藥專", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("師專", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("醫專", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("語專", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("university", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("University", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("UNIVERSITY", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("college", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("College", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("COLLEGE", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("系", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("所", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("班$", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("不分科系", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("不分系", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("department", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("Department", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("DEPARTMENT", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("^逕讀博士$", flag_person$bdegreen1), 0, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("^逕讀碩士$", flag_person$bdegreen1), 0, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("^逕行修讀碩士$", flag_person$bdegreen1), 0, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("兼課", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("最高學歷", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("^Y$", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("^待查詢$", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("^無$", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("^外國$", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("^國立$", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("^歐洲$", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("^亞洲$", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("^美洲$", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("^非洲$", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("^大洋洲$", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("肄業", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)
flag_person$err_bdegreen1 <- if_else(grepl("學分班", flag_person$bdegreen1), 1, flag_person$err_bdegreen1)

#學士學位畢業學校（一）
flag_person$err_bdegreeu1 <- 1
flag_person$err_bdegreeu1 <- if_else(grepl("大學", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("分校", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("學院", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("師大", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("科大", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("教大", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("學校", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("官校", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("預校", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("書院", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("體院", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("師院", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("專科", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("藝專", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("海專", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("工專", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("護專", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("家專", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("商專", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("行專", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("農專", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("體專", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("師專", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("藥專", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("醫專", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("語專", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("士校", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("專校$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("逕讀", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("音樂院$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
#flag_person$err_bdegreeu1 <- if_else(grepl("音樂研究所$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("university", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("University", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("UNIVERSITY", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("Uni$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("college", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("College", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("COLLEGE", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("Universidad", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("UNIVERSIDAD", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("Conservatory", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("CRD", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("ENM", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("CRC", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("EMMA", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("CRR", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("CNR", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("TheNewSchool", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("Hochschule", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("BergenSchoolofArchitecture", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("Universität", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("StellenboschUni$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^TUDarmstadt$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("大学$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^中興法商$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^N$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^待查詢$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^無$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("離職", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("因故", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("VirginiaMilitaryInstitute", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^LISAA$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^IstitutoSecoli$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^輔大$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^UnivstersityOFDelaware$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^UinvofCentralOklahoma$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^赫拉德茨克拉洛韋$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^ConservatorioStatalediMilano“GiuseppeVerdi”Italia$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^陸軍官校專科班$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("中興法商", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("國立體院", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("逕獨碩士", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("UNISA", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("Univerity", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("山口県立大学", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("ColumbiaBiblicalSeminaryandSchoolofMissions", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("本國", flag_person$bdegreeu1), 1, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("肄業", flag_person$bdegreeu1), 1, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("學分班", flag_person$bdegreeu1), 1, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("籌備處$", flag_person$bdegreeu1), 1, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^教育學院$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^高雄餐旅$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^台灣體大$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^國立空大$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^台灣體院$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^CityandGuildsofLondonArtSchool$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^BirminghamUiversity$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^NewJerseyInstituteofTechnology$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^CollegeofEducationPotchefstroomSA$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^学校法人ホンダ学園ホンダテクニカルカレッジ関西$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^同等學力$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^政戰正規班$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^日本國立埼玉大學$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^日本國立埼玉大學教養學部$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^日本國立奈良教育大學（學院）$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^日本國立熊本大學$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("^ConservatoriostatalediMilano“GiuseppeVerdi”Italia$", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("研究所$", flag_person$bdegreeu1), 1, flag_person$err_bdegreeu1)
flag_person$err_bdegreeu1 <- if_else(grepl("逕行修讀", flag_person$bdegreeu1), 0, flag_person$err_bdegreeu1)

#學士學位畢業系所（一）
flag_person$err_bdegreeg1 <- 0
flag_person$err_bdegreeg1 <- if_else(grepl("^博士$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^碩士$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^學士$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^副學士$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^大學$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^分校$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^學院$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^科大$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^學校$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^官校$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^預校$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^書院$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^專科$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^藝專$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^海專$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^工專$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^護專$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^家專$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^商專$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^行專$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^農專$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^體專$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^藥專$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^師專$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^醫專$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^語專$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^university$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^University$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^UNIVERSITY$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^college$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^College$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^COLLEGE$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("兼課", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("最高學歷", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^Y$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^待查詢$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^無$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^逕獨碩士$", flag_person$bdegreeg1), 0, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("^逕行修讀碩士$", flag_person$bdegreeg1), 0, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("肄業", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("學分班", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("行政$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(grepl("研究所$", flag_person$bdegreeg1), 1, flag_person$err_bdegreeg1)

#學士學位畢業學校國別（二）
flag_person$err_bdegreen2 <- 0
flag_person$err_bdegreen2 <- if_else(grepl("博士", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("碩士", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("學士", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("副學士", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("大學", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("分校", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("學院", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("科大", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("學校", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("官校", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("預校", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("書院", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("專科", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("藝專", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("海專", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("工專", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("護專", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("家專", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("商專", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("行專", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("農專", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("體專", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("藥專", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("師專", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("醫專", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("語專", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("university", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("University", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("UNIVERSITY", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("college", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("College", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("COLLEGE", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("系", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("所", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("班$", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("不分科系", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("不分系", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("department", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("Department", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("DEPARTMENT", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("^逕讀博士$", flag_person$bdegreen2), 0, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("^逕讀碩士$", flag_person$bdegreen2), 0, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("兼課", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("最高學歷", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("^Y$", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("^待查詢$", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("^無$", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("^外國$", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("^國立$", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("^歐洲$", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("^亞洲$", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("^美洲$", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("^非洲$", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("^大洋洲$", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("肄業", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)
flag_person$err_bdegreen2 <- if_else(grepl("學分班", flag_person$bdegreen2), 1, flag_person$err_bdegreen2)

#學士學位畢業學校（二）
flag_person$err_bdegreeu2 <- 1
flag_person$err_bdegreeu2 <- if_else(grepl("大學", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("分校", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("學院", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("師大", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("科大", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("教大", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("學校", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("官校", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("預校", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("書院", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("體院", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("師院", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("專科", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("藝專", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("海專", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("工專", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("護專", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("家專", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("商專", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("行專", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("農專", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("體專", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("師專", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("藥專", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("醫專", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("語專", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("士校", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("專校$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("逕讀", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("音樂院$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
#flag_person$err_bdegreeu2 <- if_else(grepl("音樂研究所$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("university", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("University", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("UNIVERSITY", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("Uni$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("college", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("College", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("COLLEGE", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("Universidad", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("UNIVERSIDAD", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("Conservatory", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("CRD", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("ENM", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("CRC", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("EMMA", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("CRR", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("CNR", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("TheNewSchool", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("Hochschule", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("BergenSchoolofArchitecture", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("Universität", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("StellenboschUni$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^TUDarmstadt$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("大学$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^中興法商$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^N$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^待查詢$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^無$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("離職", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("因故", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("VirginiaMilitaryInstitute", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^LISAA$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^IstitutoSecoli$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^輔大$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^UnivstersityOFDelaware$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^UinvofCentralOklahoma$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^赫拉德茨克拉洛韋$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^ConservatorioStatalediMilano“GiuseppeVerdi”Italia$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^陸軍官校專科班$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("中興法商", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("國立體院", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("逕獨碩士", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("UNISA", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("Univerity", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("山口県立大学", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("ColumbiaBiblicalSeminaryandSchoolofMissions", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("本國", flag_person$bdegreeu2), 1, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("肄業", flag_person$bdegreeu2), 1, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("學分班", flag_person$bdegreeu2), 1, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("籌備處$", flag_person$bdegreeu2), 1, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^教育學院$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^高雄餐旅$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^台灣體大$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^國立空大$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^台灣體院$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^CityandGuildsofLondonArtSchool$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^BirminghamUiversity$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^NewJerseyInstituteofTechnology$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^CollegeofEducationPotchefstroomSA$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^学校法人ホンダ学園ホンダテクニカルカレッジ関西$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^同等學力$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^政戰正規班$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^日本國立埼玉大學$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^日本國立埼玉大學教養學部$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^日本國立奈良教育大學（學院）$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^日本國立熊本大學$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("^ConservatoriostatalediMilano“GiuseppeVerdi”Italia$", flag_person$bdegreeu2), 0, flag_person$err_bdegreeu2)
flag_person$err_bdegreeu2 <- if_else(grepl("研究所$", flag_person$bdegreeu2), 1, flag_person$err_bdegreeu2)

#學士學位畢業系所（二）
flag_person$err_bdegreeg2 <- 0
flag_person$err_bdegreeg2 <- if_else(grepl("^博士$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^碩士$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^學士$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^副學士$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^大學$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^分校$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^學院$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^科大$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^學校$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^官校$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^預校$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^書院$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^專科$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^藝專$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^海專$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^工專$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^護專$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^家專$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^商專$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^行專$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^農專$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^體專$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^藥專$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^師專$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^醫專$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^語專$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^university$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^University$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^UNIVERSITY$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^college$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^College$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^COLLEGE$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("兼課", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("最高學歷", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^Y$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^待查詢$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^無$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("^逕獨碩士$", flag_person$bdegreeg2), 0, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("肄業", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("學分班", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("行政$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(grepl("研究所$", flag_person$bdegreeg2), 1, flag_person$err_bdegreeg2)

#副學士學位畢業學校國別（一）
flag_person$err_adegreen1 <- 0
flag_person$err_adegreen1 <- if_else(grepl("博士", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("碩士", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("學士", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("副學士", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("大學", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("分校", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("學院", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("科大", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("學校", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("官校", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("預校", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("書院", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("專科", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("藝專", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("海專", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("工專", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("護專", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("家專", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("商專", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("行專", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("農專", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("體專", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("藥專", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("師專", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("醫專", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("語專", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("企專", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("university", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("University", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("UNIVERSITY", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("college", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("College", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("COLLEGE", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("系", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("所", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("班$", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("不分科系", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("不分系", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("department", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("Department", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("DEPARTMENT", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("兼課", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("最高學歷", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("逕讀", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("^Y$", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("^待查詢$", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("^無$", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("^外國$", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("^國立$", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("^歐洲$", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("^亞洲$", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("^同等學力$", flag_person$adegreen1), 0, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("肄業", flag_person$adegreen1), 1, flag_person$err_adegreen1)
flag_person$err_adegreen1 <- if_else(grepl("學分班", flag_person$adegreen1), 1, flag_person$err_adegreen1)

#副學士學位畢業學校（一）
flag_person$err_adegreeu1 <- 1
flag_person$err_adegreeu1 <- if_else(grepl("大學", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("分校", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("學院", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("師大", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("科大", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("學校", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("官校", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("預校", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("書院", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("體院", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("專科", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("藝專", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("海專", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("工專", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("護專", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("家專", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("商專", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("行專", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("農專", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("體專", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("藥專", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("師專", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("醫專", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("語專", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("企專", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("士校", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("專校$", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("逕讀", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("音樂院$", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
#flag_person$err_adegreeu1 <- if_else(grepl("音樂研究所$", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("university", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("University", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("UNIVERSITY", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("Uni$", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("college", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("College", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("COLLEGE", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("Universidad", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("UNIVERSIDAD", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("Conservatory", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("CRD", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("ENM", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("CRC", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("EMMA", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("CRR", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("CNR", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("TheNewSchool", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("Hochschule", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("BergenSchoolofArchitecture", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("Universität", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("^TUDarmstadt$", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("大学$", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("^N$", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("^待查詢$", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("職業學校", flag_person$adegreeu1), 1, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("職校", flag_person$adegreeu1), 1, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("高級", flag_person$adegreeu1), 1, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("高中", flag_person$adegreeu1), 1, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("高職", flag_person$adegreeu1), 1, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("本國", flag_person$adegreeu1), 1, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("肄業", flag_person$adegreeu1), 1, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("學分班", flag_person$adegreeu1), 1, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("^仁德高級醫事職業學校$", flag_person$adegreeu1), 0, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("籌備處$", flag_person$adegreeu1), 1, flag_person$err_adegreeu1)
flag_person$err_adegreeu1 <- if_else(grepl("研究所$", flag_person$adegreeu1), 1, flag_person$err_adegreeu1)

#副學士學位畢業系所（一）
flag_person$err_adegreeg1 <- 0
flag_person$err_adegreeg1 <- if_else(grepl("^博士$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^碩士$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^學士$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^副學士$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^大學$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^分校$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^學院$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^科大$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^學校$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^官校$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^預校$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^書院$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^專科$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
#陸軍官校有"專科班"
flag_person$err_adegreeg1 <- if_else((grepl("^專科$", flag_person$adegreeg1) | grepl("^專科班$", flag_person$adegreeg1)) & grepl("陸軍官校", flag_person$adegreeu1), 0, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^藝專$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^海專$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^工專$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^護專$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^家專$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^商專$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^行專$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^農專$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^體專$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^藥專$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^師專$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^醫專$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^語專$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^企專$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^university$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^University$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^UNIVERSITY$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^college$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^College$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^COLLEGE$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("兼課", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("最高學歷", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("逕讀", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^Y$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^待查詢$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^無$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("^同等學力$", flag_person$adegreeg1), 0, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("肄業", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("學分班", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("行政$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(grepl("研究所$", flag_person$adegreeg1), 1, flag_person$err_adegreeg1)

#副學士學位畢業學校國別（二）
flag_person$err_adegreen2 <- 0
flag_person$err_adegreen2 <- if_else(grepl("博士", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("碩士", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("學士", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("副學士", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("大學", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("分校", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("學院", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("科大", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("學校", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("官校", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("預校", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("書院", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("專科", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("藝專", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("海專", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("工專", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("護專", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("家專", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("商專", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("行專", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("農專", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("體專", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("藥專", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("師專", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("醫專", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("語專", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("企專", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("university", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("University", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("UNIVERSITY", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("college", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("College", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("COLLEGE", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("系", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("所", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("班$", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("不分科系", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("不分系", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("department", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("Department", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("DEPARTMENT", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("兼課", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("最高學歷", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("逕讀", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("^Y$", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("^待查詢$", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("^無$", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("^外國$", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("^國立$", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("^歐洲$", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("^亞洲$", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("^同等學力$", flag_person$adegreen2), 0, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("肄業", flag_person$adegreen2), 1, flag_person$err_adegreen2)
flag_person$err_adegreen2 <- if_else(grepl("學分班", flag_person$adegreen2), 1, flag_person$err_adegreen2)

#副學士學位畢業學校（二）
flag_person$err_adegreeu2 <- 1
flag_person$err_adegreeu2 <- if_else(grepl("大學", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("分校", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("學院", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("師大", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("科大", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("學校", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("官校", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("預校", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("書院", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("體院", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("專科", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("藝專", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("海專", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("工專", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("護專", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("家專", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("商專", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("行專", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("農專", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("體專", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("藥專", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("師專", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("醫專", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("語專", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("企專", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("士校", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("專校$", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("逕讀", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("音樂院$", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
#flag_person$err_adegreeu2 <- if_else(grepl("音樂研究所$", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("university", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("University", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("UNIVERSITY", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("Uni$", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("college", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("College", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("COLLEGE", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("Universidad", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("UNIVERSIDAD", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("Conservatory", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("CRD", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("ENM", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("CRC", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("EMMA", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("CRR", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("CNR", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("TheNewSchool", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("Hochschule", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("BergenSchoolofArchitecture", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("Universität", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("^TUDarmstadt$", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("大学$", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("^N$", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("^待查詢$", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("職業學校", flag_person$adegreeu2), 1, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("職校", flag_person$adegreeu2), 1, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("高級", flag_person$adegreeu2), 1, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("高中", flag_person$adegreeu2), 1, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("高職", flag_person$adegreeu2), 1, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("本國", flag_person$adegreeu2), 1, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("肄業", flag_person$adegreeu2), 1, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("學分班", flag_person$adegreeu2), 1, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("^仁德高級醫事職業學校$", flag_person$adegreeu2), 0, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("籌備處$", flag_person$adegreeu2), 1, flag_person$err_adegreeu2)
flag_person$err_adegreeu2 <- if_else(grepl("研究所$", flag_person$adegreeu2), 1, flag_person$err_adegreeu2)

#副學士學位畢業系所（二）
flag_person$err_adegreeg2 <- 0
flag_person$err_adegreeg2 <- if_else(grepl("^博士$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^碩士$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^學士$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^副學士$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^大學$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^分校$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^學院$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^科大$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^學校$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^官校$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^預校$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^書院$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^專科$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
#陸軍官校有"專科班"
flag_person$err_adegreeg2 <- if_else((grepl("^專科$", flag_person$adegreeg2) | grepl("^專科班$", flag_person$adegreeg2)) & grepl("陸軍官校", flag_person$adegreeu2), 0, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^藝專$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^海專$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^工專$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^護專$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^家專$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^商專$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^行專$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^農專$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^體專$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^藥專$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^師專$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^醫專$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^語專$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^企專$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^university$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^University$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^UNIVERSITY$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^college$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^College$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^COLLEGE$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("兼課", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("最高學歷", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("逕讀", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^Y$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^待查詢$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^無$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("^同等學力$", flag_person$adegreeg2), 0, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("肄業", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("學分班", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("行政$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(grepl("研究所$", flag_person$adegreeg2), 1, flag_person$err_adegreeg2)

#學校名稱與科系名稱相同之情形
flag_person$err_ddegreeg1 <- if_else(flag_person$ddegreeu1 == flag_person$ddegreeg1 & (flag_person$ddegreeu1 != "N" & flag_person$ddegreeu1 != "逕讀碩士" & flag_person$ddegreeu1 != "逕讀博士"), 1, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg2 <- if_else(flag_person$ddegreeu2 == flag_person$ddegreeg2 & (flag_person$ddegreeu2 != "N" & flag_person$ddegreeu2 != "逕讀碩士" & flag_person$ddegreeu2 != "逕讀博士"), 1, flag_person$err_ddegreeg2)
flag_person$err_mdegreeg1 <- if_else(flag_person$mdegreeu1 == flag_person$mdegreeg1 & (flag_person$mdegreeu1 != "N" & flag_person$mdegreeu1 != "逕讀碩士" & flag_person$mdegreeu1 != "逕讀博士"), 1, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg2 <- if_else(flag_person$mdegreeu2 == flag_person$mdegreeg2 & (flag_person$mdegreeu2 != "N" & flag_person$mdegreeu2 != "逕讀碩士" & flag_person$mdegreeu2 != "逕讀博士"), 1, flag_person$err_mdegreeg2)
flag_person$err_bdegreeg1 <- if_else(flag_person$bdegreeu1 == flag_person$bdegreeg1 & (flag_person$bdegreeu1 != "N" & flag_person$bdegreeu1 != "逕讀碩士" & flag_person$bdegreeu1 != "逕行修讀碩士" & flag_person$bdegreeu1 != "逕讀博士"), 1, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg2 <- if_else(flag_person$bdegreeu2 == flag_person$bdegreeg2 & (flag_person$bdegreeu2 != "N" & flag_person$bdegreeu2 != "逕讀碩士" & flag_person$bdegreeu2 != "逕讀博士"), 1, flag_person$err_bdegreeg2)
flag_person$err_adegreeg1 <- if_else(flag_person$adegreeu1 == flag_person$adegreeg1 & (flag_person$adegreeu1 != "N" & flag_person$adegreeu1 != "逕讀碩士" & flag_person$adegreeu1 != "逕讀博士"), 1, flag_person$err_adegreeg1)
flag_person$err_adegreeg2 <- if_else(flag_person$adegreeu2 == flag_person$adegreeg2 & (flag_person$adegreeu2 != "N" & flag_person$adegreeu2 != "逕讀碩士" & flag_person$adegreeu2 != "逕讀博士"), 1, flag_person$err_adegreeg2)

#軍校當時沒有區分系所
flag_person$err_ddegreeg1 <- if_else(flag_person$ddegreeu1 == flag_person$ddegreeg1 & flag_person$ddegreeg1 == "海軍軍官學校", 0, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg1 <- if_else(flag_person$ddegreeu1 == flag_person$ddegreeg1 & flag_person$ddegreeg1 == "空軍航空技術學院", 0, flag_person$err_ddegreeg1)
flag_person$err_ddegreeg2 <- if_else(flag_person$ddegreeu1 == flag_person$ddegreeg2 & flag_person$ddegreeg2 == "海軍軍官學校", 0, flag_person$err_ddegreeg2)
flag_person$err_ddegreeg2 <- if_else(flag_person$ddegreeu1 == flag_person$ddegreeg2 & flag_person$ddegreeg2 == "空軍航空技術學院", 0, flag_person$err_ddegreeg2)
flag_person$err_mdegreeg1 <- if_else(flag_person$mdegreeu1 == flag_person$mdegreeg1 & flag_person$mdegreeg1 == "海軍軍官學校", 0, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg1 <- if_else(flag_person$mdegreeu1 == flag_person$mdegreeg1 & flag_person$mdegreeg1 == "空軍航空技術學院", 0, flag_person$err_mdegreeg1)
flag_person$err_mdegreeg2 <- if_else(flag_person$mdegreeu1 == flag_person$mdegreeg2 & flag_person$mdegreeg2 == "海軍軍官學校", 0, flag_person$err_mdegreeg2)
flag_person$err_mdegreeg2 <- if_else(flag_person$mdegreeu1 == flag_person$mdegreeg2 & flag_person$mdegreeg2 == "空軍航空技術學院", 0, flag_person$err_mdegreeg2)
flag_person$err_bdegreeg1 <- if_else(flag_person$bdegreeu1 == flag_person$bdegreeg1 & flag_person$bdegreeg1 == "海軍軍官學校", 0, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg1 <- if_else(flag_person$bdegreeu1 == flag_person$bdegreeg1 & flag_person$bdegreeg1 == "空軍航空技術學院", 0, flag_person$err_bdegreeg1)
flag_person$err_bdegreeg2 <- if_else(flag_person$bdegreeu1 == flag_person$bdegreeg2 & flag_person$bdegreeg2 == "海軍軍官學校", 0, flag_person$err_bdegreeg2)
flag_person$err_bdegreeg2 <- if_else(flag_person$bdegreeu1 == flag_person$bdegreeg2 & flag_person$bdegreeg2 == "空軍航空技術學院", 0, flag_person$err_bdegreeg2)
flag_person$err_adegreeg1 <- if_else(flag_person$adegreeu1 == flag_person$adegreeg1 & flag_person$adegreeg1 == "海軍軍官學校", 0, flag_person$err_adegreeg1)
flag_person$err_adegreeg1 <- if_else(flag_person$adegreeu1 == flag_person$adegreeg1 & flag_person$adegreeg1 == "空軍航空技術學院", 0, flag_person$err_adegreeg1)
flag_person$err_adegreeg2 <- if_else(flag_person$adegreeu1 == flag_person$adegreeg2 & flag_person$adegreeg2 == "海軍軍官學校", 0, flag_person$err_adegreeg2)
flag_person$err_adegreeg2 <- if_else(flag_person$adegreeu1 == flag_person$adegreeg2 & flag_person$adegreeg2 == "空軍航空技術學院", 0, flag_person$err_adegreeg2)

#學士逕讀碩士，但副學士為N
flag_person$err_bdeade <- 0
flag_person$err_bdeade <- if_else(flag_person$bdegreeu1 == "逕讀碩士" & flag_person$adegreeu1 == "N", 1, flag_person$err_bdeade)
flag_person$err_bdeade <- if_else(flag_person$bdegreeu1 == "逕讀" & flag_person$adegreeu1 == "N", 1, flag_person$err_bdeade)
flag_person$err_bdeade <- if_else(grepl("讀", flag_person$bdegreeu1) & flag_person$adegreeu1 == "N", 1, flag_person$err_bdeade)

#碩士逕讀碩士，但學士為N
flag_person$err_bdeade2 <- 0
flag_person$err_bdeade2 <- if_else(flag_person$mdegreeu1 == "逕讀博士" & flag_person$bdegreeu1 == "N", 1, flag_person$err_bdeade2)
flag_person$err_bdeade2 <- if_else(flag_person$mdegreeu1 == "逕讀" & flag_person$bdegreeu1 == "N", 1, flag_person$err_bdeade2)

#碩士、學士逕讀博士，直接有博士
flag_person$err_bdeade3 <- 0
flag_person$err_bdeade3 <- if_else(flag_person$mdegreeu1 == "逕讀博士" & flag_person$bdegreeu1 == "逕讀博士", 1, flag_person$err_bdeade3)
flag_person$err_bdeade3 <- if_else(flag_person$mdegreeu1 == "逕讀" & flag_person$bdegreeu1 == "逕讀", 1, flag_person$err_bdeade3)

flag_person$err_bdeade <- if_else(flag_person$err_bdeade3 == 1 & flag_person$err_bdeade == 1, 0, flag_person$err_bdeade)

#有副學士，學士填逕讀碩士，但碩士為N
flag_person$err_bdeade6 <- 0
flag_person$err_bdeade6 <- if_else(flag_person$adegreeu1 != "N" & grepl("讀", flag_person$bdegreeu1) & flag_person$mdegreeu1 == "N", 1, flag_person$err_bdeade6)

#有學士，碩士填逕讀博士，但博士為N
flag_person$err_bdeade7 <- 0
flag_person$err_bdeade7 <- if_else(flag_person$bdegreeu1 != "N" & grepl("讀", flag_person$mdegreeu1) & flag_person$ddegreeu1 == "N", 1, flag_person$err_bdeade7)


#學校為國外學校，但國別卻填本國
flag_person$err_bdeade4 <- 0
#博士(一)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("A", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("E", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("I", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("O", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("U", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("a", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("e", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("i", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("o", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("u", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^英國", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^美國", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^加拿大", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^日本", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^韓國", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^菲律賓", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^南非", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^西班牙", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^法國", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^俄羅斯", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^德國", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^澳洲", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^紐西蘭", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^義大利", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^比利時", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^芬蘭", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^澳大利亞", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^泰國", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^美利堅合眾國", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^印尼", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^阿根廷", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^越南", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^香港", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="英國" & grepl("^彰化", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^澳大利亞", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("^荷蘭", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("紐約", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("州立", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("東京", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("波士頓", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("路易安納", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("關西", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("專修大學", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("匹茲堡", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("俄克拉荷馬", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("雪菲爾", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen1=="本國" & grepl("胡志明", flag_person$ddegreeu1), 1, flag_person$err_bdeade4)
#博士(二)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("A", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("E", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("I", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("O", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("U", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("a", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("e", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("i", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("o", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("u", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^英國", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^美國", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^加拿大", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^日本", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^韓國", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^菲律賓", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^南非", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^西班牙", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^法國", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^俄羅斯", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^德國", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^澳洲", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^紐西蘭", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^義大利", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^比利時", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^芬蘭", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^澳大利亞", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^泰國", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^美利堅合眾國", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^印尼", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^阿根廷", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^越南", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^香港", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="英國" & grepl("^彰化", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^澳大利亞", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("^荷蘭", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("紐約", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("州立", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("東京", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("波士頓", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("路易安納", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("關西", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("專修大學", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("匹茲堡", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("俄克拉荷馬", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("雪菲爾", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$ddegreen2=="本國" & grepl("胡志明", flag_person$ddegreeu2), 1, flag_person$err_bdeade4)
#碩士(一)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("A", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("E", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("I", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("O", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("U", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("a", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("e", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("i", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("o", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("u", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^英國", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^美國", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^加拿大", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^日本", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^韓國", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^菲律賓", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^南非", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^西班牙", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^法國", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^俄羅斯", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^德國", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^澳洲", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^紐西蘭", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^義大利", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^比利時", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^芬蘭", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^澳大利亞", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^泰國", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^美利堅合眾國", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^印尼", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^阿根廷", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^越南", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^香港", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="英國" & grepl("^彰化", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^澳大利亞", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("^荷蘭", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("紐約", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("州立", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("東京", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("波士頓", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("路易安納", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("關西", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("專修大學", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("匹茲堡", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("俄克拉荷馬", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("雪菲爾", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen1=="本國" & grepl("胡志明", flag_person$mdegreeu1), 1, flag_person$err_bdeade4)
#碩士(二)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("A", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("E", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("I", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("O", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("U", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("a", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("e", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("i", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("o", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("u", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^英國", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^美國", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^加拿大", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^日本", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^韓國", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^菲律賓", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^南非", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^西班牙", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^法國", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^俄羅斯", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^德國", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^澳洲", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^紐西蘭", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^義大利", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^比利時", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^芬蘭", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^澳大利亞", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^泰國", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^美利堅合眾國", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^印尼", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^阿根廷", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^越南", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^香港", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="英國" & grepl("^彰化", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^澳大利亞", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("^荷蘭", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("紐約", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("州立", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("東京", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("波士頓", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("路易安納", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("關西", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("專修大學", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("匹茲堡", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("俄克拉荷馬", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("雪菲爾", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$mdegreen2=="本國" & grepl("胡志明", flag_person$mdegreeu2), 1, flag_person$err_bdeade4)
#學士(一)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("A", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("E", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("I", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("O", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("U", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("a", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("e", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("i", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("o", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("u", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^英國", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^美國", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^加拿大", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^日本", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^韓國", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^菲律賓", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^南非", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^西班牙", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^法國", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^俄羅斯", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^德國", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^澳洲", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^紐西蘭", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^義大利", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^比利時", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^芬蘭", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^澳大利亞", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^泰國", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^美利堅合眾國", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^印尼", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^阿根廷", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^越南", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^香港", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="英國" & grepl("^彰化", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^澳大利亞", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("^荷蘭", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("紐約", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("州立", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("東京", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("波士頓", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("路易安納", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("關西", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("專修大學", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("匹茲堡", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("俄克拉荷馬", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("雪菲爾", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen1=="本國" & grepl("胡志明", flag_person$bdegreeu1), 1, flag_person$err_bdeade4)
#學士(二)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("A", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("E", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("I", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("O", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("U", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("a", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("e", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("i", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("o", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("u", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^英國", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^美國", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^加拿大", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^日本", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^韓國", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^菲律賓", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^南非", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^西班牙", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^法國", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^俄羅斯", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^德國", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^澳洲", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^紐西蘭", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^義大利", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^比利時", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^芬蘭", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^澳大利亞", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^泰國", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^美利堅合眾國", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^印尼", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^阿根廷", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^越南", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^香港", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="英國" & grepl("^彰化", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^澳大利亞", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("^荷蘭", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("紐約", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("州立", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("東京", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("波士頓", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("路易安納", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("關西", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("專修大學", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("匹茲堡", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("俄克拉荷馬", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("雪菲爾", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$bdegreen2=="本國" & grepl("胡志明", flag_person$bdegreeu2), 1, flag_person$err_bdeade4)
#副學士(一)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen1=="本國" & grepl("A", flag_person$adegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen1=="本國" & grepl("E", flag_person$adegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen1=="本國" & grepl("I", flag_person$adegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen1=="本國" & grepl("O", flag_person$adegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen1=="本國" & grepl("U", flag_person$adegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen1=="本國" & grepl("a", flag_person$adegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen1=="本國" & grepl("e", flag_person$adegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen1=="本國" & grepl("i", flag_person$adegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen1=="本國" & grepl("o", flag_person$adegreeu1), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen1=="本國" & grepl("u", flag_person$adegreeu1), 1, flag_person$err_bdeade4)
#副學士(二)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen2=="本國" & grepl("A", flag_person$adegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen2=="本國" & grepl("E", flag_person$adegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen2=="本國" & grepl("I", flag_person$adegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen2=="本國" & grepl("O", flag_person$adegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen2=="本國" & grepl("U", flag_person$adegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen2=="本國" & grepl("a", flag_person$adegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen2=="本國" & grepl("e", flag_person$adegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen2=="本國" & grepl("i", flag_person$adegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen2=="本國" & grepl("o", flag_person$adegreeu2), 1, flag_person$err_bdeade4)
flag_person$err_bdeade4 <- if_else(flag_person$adegreen2=="本國" & grepl("u", flag_person$adegreeu2), 1, flag_person$err_bdeade4)

#逕讀碩士國別填本國
flag_person$err_bdeade5 <- 0
flag_person$err_bdeade5 <- if_else(flag_person$bdegreeu1 == "逕讀碩士" & grepl("國", flag_person$bdegreen1), 1, flag_person$err_bdeade5)

flag_person$err_flag_sp6 <- flag_person$err_ddegreen1 + flag_person$err_ddegreeu1 + flag_person$err_ddegreeg1 + flag_person$err_ddegreen2 + flag_person$err_ddegreeu2 + flag_person$err_ddegreeg2 + flag_person$err_mdegreen1 + flag_person$err_mdegreeu1 + flag_person$err_mdegreeg1 + flag_person$err_mdegreen2 + flag_person$err_mdegreeu2 + flag_person$err_mdegreeg2 + flag_person$err_bdegreen1 + flag_person$err_bdegreeu1 + flag_person$err_bdegreeg1 + flag_person$err_bdegreen2 + flag_person$err_bdegreeu2 + flag_person$err_bdegreeg2 + flag_person$err_adegreen1 + flag_person$err_adegreeu1 + flag_person$err_adegreeg1 + flag_person$err_adegreen2 + flag_person$err_adegreeu2 + flag_person$err_adegreeg2 + flag_person$err_bdeade + flag_person$err_bdeade2 + flag_person$err_bdeade3 + flag_person$err_bdeade4 + flag_person$err_bdeade5+ flag_person$err_bdeade6+ flag_person$err_bdeade7

flag_person$err_flag <- 0
flag_person$err_flag <- if_else(flag_person$err_flag_sp6 != 0, 1, flag_person$err_flag)

#兼任或鐘點教師，次高學歷可不填

#最高學歷為碩士，但沒填學士學歷
flag_person$err_flag <- if_else(flag_person$source == 1 & (flag_person$emptype == "長期代課" | flag_person$emptype == "兼任" | flag_person$emptype == "鐘點教師") & flag_person$mdegreen1 != "" & (flag_person$bdegreen1 == "NA" | flag_person$bdegreen1 == "無法取得資料" | flag_person$bdegreen1 == "待查詢" ), 0, flag_person$err_flag)
#最高學歷為博士，但沒填學士學歷
flag_person$err_flag <- if_else(flag_person$source == 1 & (flag_person$emptype == "長期代課" | flag_person$emptype == "兼任" | flag_person$emptype == "鐘點教師") & flag_person$ddegreen1 != "" & (flag_person$bdegreen1 == "NA" | flag_person$bdegreen1 == "無法取得資料" | flag_person$bdegreen1 == "待查詢" ), 0, flag_person$err_flag)
#最高學歷為博士，但沒填碩士學歷
flag_person$err_flag <- if_else(flag_person$source == 1 & (flag_person$emptype == "長期代課" | flag_person$emptype == "兼任" | flag_person$emptype == "鐘點教師") & flag_person$ddegreen1 != "" & (flag_person$mdegreen1 == "NA" | flag_person$mdegreen1 == "無法取得資料" | flag_person$mdegreen1 == "待查詢" ), 0, flag_person$err_flag)

#加註
flag_person$name <- paste(flag_person$name, "（", sep = "")
flag_person$name <- if_else(flag_person$err_ddegreen1 != 0, paste(flag_person$name, "博士學位畢業學校國別（一）：", flag_person$ddegreen1, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_ddegreeu1 != 0, paste(flag_person$name, "博士學位畢業學校（一）：", flag_person$ddegreeu1, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_ddegreeg1 != 0, paste(flag_person$name, "博士學位畢業系所（一）：", flag_person$ddegreeg1, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_ddegreen2 != 0, paste(flag_person$name, "博士學位畢業學校國別（二）：", flag_person$ddegreen2, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_ddegreeu2 != 0, paste(flag_person$name, "博士學位畢業學校（二）：", flag_person$ddegreeu2, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_ddegreeg2 != 0, paste(flag_person$name, "博士學位畢業系所（二）：", flag_person$ddegreeg2, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_mdegreen1 != 0, paste(flag_person$name, "碩士學位畢業學校國別（一）：", flag_person$mdegreen1, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_mdegreeu1 != 0, paste(flag_person$name, "碩士學位畢業學校（一）：", flag_person$mdegreeu1, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_mdegreeg1 != 0, paste(flag_person$name, "碩士學位畢業系所（一）：", flag_person$mdegreeg1, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_mdegreen2 != 0, paste(flag_person$name, "碩士學位畢業學校國別（二）：", flag_person$mdegreen2, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_mdegreeu2 != 0, paste(flag_person$name, "碩士學位畢業學校（二）：", flag_person$mdegreeu2, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_mdegreeg2 != 0, paste(flag_person$name, "碩士學位畢業系所（二）：", flag_person$mdegreeg2, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_bdegreen1 != 0, paste(flag_person$name, "學士學位畢業學校國別（一）：", flag_person$bdegreen1, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_bdegreeu1 != 0, paste(flag_person$name, "學士學位畢業學校（一）：", flag_person$bdegreeu1, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_bdegreeg1 != 0, paste(flag_person$name, "學士學位畢業科系（一）：", flag_person$bdegreeg1, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_bdegreen2 != 0, paste(flag_person$name, "學士學位畢業學校國別（二）：", flag_person$bdegreen2, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_bdegreeu2 != 0, paste(flag_person$name, "學士學位畢業學校（二）：", flag_person$bdegreeu2, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_bdegreeg2 != 0, paste(flag_person$name, "學士學位畢業科系（二）：", flag_person$bdegreeg2, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_adegreen1 != 0, paste(flag_person$name, "副學士或專科畢業學校國別（一）：", flag_person$adegreen1, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_adegreeu1 != 0, paste(flag_person$name, "副學士或專科畢業學校（一）：", flag_person$adegreeu1, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_adegreeg1 != 0, paste(flag_person$name, "副學士或專科畢業科系（一）：", flag_person$adegreeg1, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_adegreen2 != 0, paste(flag_person$name, "副學士或專科畢業學校國別（二）：", flag_person$adegreen2, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_adegreeu2 != 0, paste(flag_person$name, "副學士或專科畢業學校（二）：", flag_person$adegreeu2, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_adegreeg2 != 0, paste(flag_person$name, "副學士或專科畢業科系（二）：", flag_person$adegreeg2, "；", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_bdeade != 0, paste(flag_person$name, "副學士或專科畢業學校國別（一）：", flag_person$adegreen1, "、副學士或專科畢業學校（一）：", flag_person$adegreeu1, "、副學士或專科畢業科系（一）：",  flag_person$adegreeg1, " (若逕讀碩士，副學士或專科畢業資訊應不為N)", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_bdeade2 != 0, paste(flag_person$name, "(若逕讀博士，學士畢業資訊應不為N)", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_bdeade3 != 0, paste(flag_person$name, "(若逕讀博士，學士或專科畢業資訊應不為逕讀博士)", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_bdeade4 != 0, paste(flag_person$name, "若於外國學校取得學位，其學位畢業學校國別不應為本國", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_bdeade5 != 0, paste(flag_person$name, "若為逕讀碩士，相關欄位請依欄位說明填寫", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_bdeade6 != 0, paste(flag_person$name, "若為逕讀碩士，應填列碩士學歷相關資訊", sep = ""), flag_person$name)
flag_person$name <- if_else(flag_person$err_bdeade7 != 0, paste(flag_person$name, "若為逕讀博士，應填列博士學歷相關資訊", sep = ""), flag_person$name)
flag_person$name <- paste(flag_person$name, "）", sep = "")
flag_person$name <- gsub("；）", replacement = "）", flag_person$name)
flag_person$name <- gsub("（）", replacement = "", flag_person$name)

# （請確認畢業科系正確名稱）
# （請確認畢業系所正確名稱）
# （請確認畢業學校正確名稱）
# （請依照*員『*士』學位畢業證書，填寫正確之學校名稱。）
# （請依照*員『*士』學位畢業證書，填寫正確之系所名稱。）
# （請依照*員『*士』學位畢業證書，填寫正確之科系名稱。）
# （請確認*員畢業學校正確名稱及拼字是否正確）
# （請依照上開人員碩士學位畢業證書，填寫正確之系所名稱全稱。）
# 
# 40學分班之文字
# （學分班非屬『學位授予法』規定之學位別，且依該法規定，須修業期滿、修滿應修學分並符畢業條件，始能獲頒學位。若謝員經確認未獲碩士學位，請於碩士學位畢業學校國別、畢業學校、畢業系所三欄填『N』）
# （請確認*員最高學歷，若*員最高學歷不為大專以上，「最高學歷是否為大專以上」及各級學歷資訊欄位請皆填「N」。）

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_spe6 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_spe6)[3 : length(colnames(flag_person_wide_spe6))]
flag_person_wide_spe6$spe6_r <- NA
for (i in temp){
  flag_person_wide_spe6$spe6_r <- paste(flag_person_wide_spe6$spe6_r, flag_person_wide_spe6[[i]], sep = " ")
}
flag_person_wide_spe6$spe6_r <- gsub("NA ", replacement="", flag_person_wide_spe6$spe6_r)
flag_person_wide_spe6$spe6_r <- gsub(" NA", replacement="", flag_person_wide_spe6$spe6_r)

#產生檢誤報告文字
spe6_temp <- flag_person_wide_spe6 %>%
  group_by(organization_id) %>%
  mutate(spe6_txt = paste(source, "之大學（學士）以上各教育階段學歷資料不完整或不正確：", spe6_r, sep = ""), "") %>%
  subset(select = c(organization_id, spe6_txt)) %>%
  distinct(organization_id, spe6_txt)

#根據organization_id，展開成寬資料(wide)
spe6 <- spe6_temp %>%
  dcast(organization_id ~ spe6_txt, value.var = "spe6_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(spe6)[2 : length(colnames(spe6))]
spe6$spe6 <- NA
for (i in temp){
  spe6$spe6 <- paste(spe6$spe6, spe6[[i]], sep = "； ")
}
spe6$spe6 <- gsub("NA； ", replacement="", spe6$spe6)
spe6$spe6 <- gsub("； NA", replacement="", spe6$spe6)

#產生檢誤報告文字
spe6 <- spe6 %>%
  subset(select = c(organization_id, spe6)) %>%
  distinct(organization_id, spe6)
}else{
#偵測spe6是否存在。若不存在，則產生NA行
if('spe6' %in% ls()){
  print("spe6")
}else{
  spe6 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  spe6$spe6 <- ""
}
}
# flag83: 離退教職員（工）資料表所列人員，不應填列為本次教員資料表或職員（工）資料表之專任或代理人員。 -------------------------------------------------------------------
flag_person <- drev_P_retire %>%
  rename(name = name.x, name_retire = name.y)

#離退教職員(工)資料表所列人員，不應填列為本次教員資料表或職員（工）資料表之專任或代理人員。(與本期資料比對)
#抓出:離退人員有出現在教員資料表、職員工資料表，且為專任或代理
flag_person$err_flag <- 0
flag_person$err_flag <- if_else(!is.na(flag_person$name_retire) & (flag_person$emptype == "專任" | flag_person$emptype == "代理"), 1, flag_person$err_flag)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag83 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag83)[3 : length(colnames(flag_person_wide_flag83))]
flag_person_wide_flag83$flag83_r <- NA
for (i in temp){
  flag_person_wide_flag83$flag83_r <- paste(flag_person_wide_flag83$flag83_r, flag_person_wide_flag83[[i]], sep = " ")
}
flag_person_wide_flag83$flag83_r <- gsub("NA ", replacement="", flag_person_wide_flag83$flag83_r)
flag_person_wide_flag83$flag83_r <- gsub(" NA", replacement="", flag_person_wide_flag83$flag83_r)

#產生檢誤報告文字
flag83_temp <- flag_person_wide_flag83 %>%
  group_by(organization_id) %>%
  mutate(flag83_txt = paste(source, "：", flag83_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag83_txt)) %>%
  distinct(organization_id, flag83_txt)

#根據organization_id，展開成寬資料(wide)
flag83 <- flag83_temp %>%
  dcast(organization_id ~ flag83_txt, value.var = "flag83_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag83)[2 : length(colnames(flag83))]
flag83$flag83 <- NA
for (i in temp){
  flag83$flag83 <- paste(flag83$flag83, flag83[[i]], sep = "； ")
}
flag83$flag83 <- gsub("NA； ", replacement="", flag83$flag83)
flag83$flag83 <- gsub("； NA", replacement="", flag83$flag83)

#產生檢誤報告文字
flag83 <- flag83 %>%
  subset(select = c(organization_id, flag83)) %>%
  distinct(organization_id, flag83) %>%
  mutate(flag83 = paste(flag83, "（請確認上述人員是否退休、退伍或因故離職，若是，則不需填列至本次教員資料表或職員（工）資料表）", sep = ""))
}else{
#偵測flag83是否存在。若不存在，則產生NA行
if('flag83' %in% ls()){
  print("flag83")
}else{
  flag83 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag83$flag83 <- ""
}
}
# flag92: 教員/職員（工）資料表及離退教職員(工)資料表，同一身分識別碼所對應的姓名不一致，請確認各該身分識別碼所屬正確人員。 -------------------------------------------------------------------
flag_person <- drev_P_retire %>%
  rename(name = name.x, name_retire = name.y)

#本次離退教職員(工)資料表所列人員，若有出現本次教員資料表或職員（工）資料表(已存在錯誤情形)，姓名不一致
flag_person$err_flag <- 0
flag_person$err_flag <- if_else(flag_person$name != flag_person$name_retire, 1, flag_person$err_flag)

#加註
flag_person$name <- paste(flag_person$name, "/", flag_person$name_retire, sep = "")
flag_person$name <- gsub("；）", replacement = "）", flag_person$name)
flag_person$name <- gsub("（）", replacement = "", flag_person$name)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag92 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag92)[3 : length(colnames(flag_person_wide_flag92))]
flag_person_wide_flag92$flag92_r <- NA
for (i in temp){
  flag_person_wide_flag92$flag92_r <- paste(flag_person_wide_flag92$flag92_r, flag_person_wide_flag92[[i]], sep = " ")
}
flag_person_wide_flag92$flag92_r <- gsub("NA ", replacement="", flag_person_wide_flag92$flag92_r)
flag_person_wide_flag92$flag92_r <- gsub(" NA", replacement="", flag_person_wide_flag92$flag92_r)

#產生檢誤報告文字
flag92_temp <- flag_person_wide_flag92 %>%
  group_by(organization_id) %>%
  mutate(flag92_txt = paste("請確認：", flag92_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag92_txt)) %>%
  distinct(organization_id, flag92_txt)

#根據organization_id，展開成寬資料(wide)
flag92 <- flag92_temp %>%
  dcast(organization_id ~ flag92_txt, value.var = "flag92_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag92)[2 : length(colnames(flag92))]
flag92$flag92 <- NA
for (i in temp){
  flag92$flag92 <- paste(flag92$flag92, flag92[[i]], sep = "； ")
}
flag92$flag92 <- gsub("NA； ", replacement="", flag92$flag92)
flag92$flag92 <- gsub("； NA", replacement="", flag92$flag92)

#產生檢誤報告文字
flag92 <- flag92 %>%
  subset(select = c(organization_id, flag92)) %>%
  distinct(organization_id, flag92) %>%
  mutate(flag92 = paste(flag92, "（教員/職員（工）資料表及離退教職員(工)資料表，同一身分識別碼所對應的姓名不一致，請確認各該身分識別碼所屬正確人員。）", sep = ""))
}else{
#偵測flag92是否存在。若不存在，則產生NA行
if('flag92' %in% ls()){
  print("flag92")
}else{
  flag92 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag92$flag92 <- ""
}
}
# flag84: 離退教職員（工）資料表所列人員，應為上一學年（期）專任教職員（工）。 -------------------------------------------------------------------
flag_person <- drev_P_retire_pre_inner %>%
  rename(name = name.x, name_retire = name.y) %>%
  left_join(edu_name2, by = c("organization_id"))

#填寫在「離退教職員(工)資料表」之人員，聘任類別需為「專任」。(與上一期資料比對)
flag_person$err_flag <- 0
flag_person$err_flag <- if_else(flag_person$emptype != "專任", 1, flag_person$err_flag)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag84 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag84)[3 : length(colnames(flag_person_wide_flag84))]
flag_person_wide_flag84$flag84_r <- NA
for (i in temp){
  flag_person_wide_flag84$flag84_r <- paste(flag_person_wide_flag84$flag84_r, flag_person_wide_flag84[[i]], sep = " ")
}
flag_person_wide_flag84$flag84_r <- gsub("NA ", replacement="", flag_person_wide_flag84$flag84_r)
flag_person_wide_flag84$flag84_r <- gsub(" NA", replacement="", flag_person_wide_flag84$flag84_r)

#產生檢誤報告文字
flag84_temp <- flag_person_wide_flag84 %>%
  group_by(organization_id) %>%
  mutate(flag84_txt = paste(source, "：", flag84_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag84_txt)) %>%
  distinct(organization_id, flag84_txt)

#根據organization_id，展開成寬資料(wide)
flag84 <- flag84_temp %>%
  dcast(organization_id ~ flag84_txt, value.var = "flag84_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag84)[2 : length(colnames(flag84))]
flag84$flag84 <- NA
for (i in temp){
  flag84$flag84 <- paste(flag84$flag84, flag84[[i]], sep = "； ")
}
flag84$flag84 <- gsub("NA； ", replacement="", flag84$flag84)
flag84$flag84 <- gsub("； NA", replacement="", flag84$flag84)

#產生檢誤報告文字
flag84 <- flag84 %>%
  subset(select = c(organization_id, flag84)) %>%
  distinct(organization_id, flag84) %>%
  mutate(flag84 = paste(flag84, "（查貴校上一學年所填資料，上述人員聘任類別非屬『專任』。依欄位說明，非專任教職員(工)之退休或離職者，不須填列離退教職員(工)資料表，請務必再確認。）", sep = ""))
}else{
#偵測flag84是否存在。若不存在，則產生NA行
if('flag84' %in% ls()){
  print("flag84")
}else{
  flag84 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag84$flag84 <- ""
}
}
# flag85: 離退教職員（工）資料表中，離職退休情況註記為「退休」人員之年齡偏小。 -------------------------------------------------------------------
flag_person <- drev_P_retire_pre_inner %>%
  rename(name = name.x, name_retire = name.y) %>%
  left_join(edu_name2, by = c("organization_id"))

#離職退休情況為「退休」之人員年齡低於42歲。(與上一期資料比對)
#年齡
#創設變項出生年月日：birthy birthm birthd
flag_person$birthy <- ""
flag_person$birthm <- ""
flag_person$birthd <- ""

flag_person$birthy <- if_else(nchar(flag_person$birthdate) == 6, substr(flag_person$birthdate, 1, 2), flag_person$birthy)
flag_person$birthm <- if_else(nchar(flag_person$birthdate) == 6, substr(flag_person$birthdate, 3, 4), flag_person$birthm)
flag_person$birthd <- if_else(nchar(flag_person$birthdate) == 6, substr(flag_person$birthdate, 5, 6), flag_person$birthd)
flag_person$birthy <- if_else(nchar(flag_person$birthdate) == 7, substr(flag_person$birthdate, 1, 3), flag_person$birthy)
flag_person$birthm <- if_else(nchar(flag_person$birthdate) == 7, substr(flag_person$birthdate, 4, 5), flag_person$birthm)
flag_person$birthd <- if_else(nchar(flag_person$birthdate) == 7, substr(flag_person$birthdate, 6, 7), flag_person$birthd)

flag_person$birthy <- as.numeric(flag_person$birthy)
flag_person$birthm <- as.numeric(flag_person$birthm)
flag_person$birthd <- as.numeric(flag_person$birthd)

flag_person$survey_year <- 2022

#創設變項年齡（以年為單位）：age
flag_person$age <- 0
flag_person$age <- if_else(flag_person$survey_year %% 4 != 0, ((flag_person$survey_year-1911) + 9/12 + 30/365) - (flag_person$birthy + (flag_person$birthm/12) + (flag_person$birthd/365)), flag_person$age)
flag_person$age <- if_else(flag_person$survey_year %% 4 == 0, ((flag_person$survey_year-1911) + 9/12 + 30/366) - (flag_person$birthy + (flag_person$birthm/12) + (flag_person$birthd/366)), flag_person$age)


flag_person$err_flag <- 0
flag_person$err_flag <- if_else(flag_person$resitu == "R" & flag_person$age < 42, 1, flag_person$err_flag)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag85 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag85)[3 : length(colnames(flag_person_wide_flag85))]
flag_person_wide_flag85$flag85_r <- NA
for (i in temp){
  flag_person_wide_flag85$flag85_r <- paste(flag_person_wide_flag85$flag85_r, flag_person_wide_flag85[[i]], sep = " ")
}
flag_person_wide_flag85$flag85_r <- gsub("NA ", replacement="", flag_person_wide_flag85$flag85_r)
flag_person_wide_flag85$flag85_r <- gsub(" NA", replacement="", flag_person_wide_flag85$flag85_r)

#產生檢誤報告文字
flag85_temp <- flag_person_wide_flag85 %>%
  group_by(organization_id) %>%
  mutate(flag85_txt = paste(source, "：", flag85_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag85_txt)) %>%
  distinct(organization_id, flag85_txt)

#根據organization_id，展開成寬資料(wide)
flag85 <- flag85_temp %>%
  dcast(organization_id ~ flag85_txt, value.var = "flag85_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag85)[2 : length(colnames(flag85))]
flag85$flag85 <- NA
for (i in temp){
  flag85$flag85 <- paste(flag85$flag85, flag85[[i]], sep = "； ")
}
flag85$flag85 <- gsub("NA； ", replacement="", flag85$flag85)
flag85$flag85 <- gsub("； NA", replacement="", flag85$flag85)

#產生檢誤報告文字
flag85 <- flag85 %>%
  subset(select = c(organization_id, flag85)) %>%
  distinct(organization_id, flag85) %>%
  mutate(flag85 = paste(flag85, "（該員年齡似低於最低法定退休年齡，敬請再協助確認）", sep = ""))
}else{
#偵測flag85是否存在。若不存在，則產生NA行
if('flag85' %in% ls()){
  print("flag85")
}else{
  flag85 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag85$flag85 <- ""
}
}
# flag93: 離退教職員（工）資料表所列人員，應為上一學年（期）之教員或職員（工）。 -------------------------------------------------------------------
flag_person <- drev_P_retire_pre_right %>%
  rename(name = name.x, name_retire = name.y) %>%
  left_join(edu_name2, by = c("organization_id"))

#填寫在「離退教職員(工)資料表」之人員，聘任類別需為「專任」。(與上一期資料比對)
#抓出:離退人員在上一期聘任類別非專任(專任的人員才能填到離退表)
flag_person$err_flag <- 0
flag_person$err_flag <- if_else(is.na(flag_person$pre), 1, flag_person$err_flag)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ flag_person$name_retire,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag93 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, source, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id + source ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag93)[3 : length(colnames(flag_person_wide_flag93))]
flag_person_wide_flag93$flag93_r <- NA
for (i in temp){
  flag_person_wide_flag93$flag93_r <- paste(flag_person_wide_flag93$flag93_r, flag_person_wide_flag93[[i]], sep = " ")
}
flag_person_wide_flag93$flag93_r <- gsub("NA ", replacement="", flag_person_wide_flag93$flag93_r)
flag_person_wide_flag93$flag93_r <- gsub(" NA", replacement="", flag_person_wide_flag93$flag93_r)

#產生檢誤報告文字
flag93_temp <- flag_person_wide_flag93 %>%
  group_by(organization_id) %>%
  mutate(flag93_txt = paste("離退教職員(工)資料表：", flag93_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag93_txt)) %>%
  distinct(organization_id, flag93_txt)

#根據organization_id，展開成寬資料(wide)
flag93 <- flag93_temp %>%
  dcast(organization_id ~ flag93_txt, value.var = "flag93_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag93)[2 : length(colnames(flag93))]
flag93$flag93 <- NA
for (i in temp){
  flag93$flag93 <- paste(flag93$flag93, flag93[[i]], sep = "； ")
}
flag93$flag93 <- gsub("NA； ", replacement="", flag93$flag93)
flag93$flag93 <- gsub("； NA", replacement="", flag93$flag93)

#產生檢誤報告文字
flag93 <- flag93 %>%
  subset(select = c(organization_id, flag93)) %>%
  distinct(organization_id, flag93) %>%
  mutate(flag93 = paste(flag93, "（查貴校上一學年所填資料，上述人員聘任類別非屬『專任』。依欄位說明，非專任教職員(工)之退休或離職者，不須填列離退教職員(工)資料表，請務必再確認。）", sep = ""))
}else{
#偵測flag93是否存在。若不存在，則產生NA行
if('flag93' %in% ls()){
  print("flag93")
}else{
  flag93 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag93$flag93 <- ""
}
}
# flag86: 上一學年（期）專任教職員（工）若未於本學年（期）教職員（工）資料表內，則應填列於離退教職員（工）資料表。 -------------------------------------------------------------------
flag_person <- drev_P_retire_merge_pre %>%
  rename(name = name.x, name_retire = name, edu_name2 = edu_name2.x)

#抓出:有出現在上一期資料，但在本次填報已被刪除，但退休資料表沒有出現的專任人員
#只出現在上一期(pre = 1 & now = NA)   且為專任    但本次離退表卻沒有出現(空白)(retire = NA)
flag_person$err_flag <- 0
flag_person$err_flag <- if_else(flag_person$pre == 1 & is.na(flag_person$now) & is.na(flag_person$retire) & flag_person$emptype.y == "專任", 1, flag_person$err_flag)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ flag_person$name.y,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag86 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag86)[3 : length(colnames(flag_person_wide_flag86))]
flag_person_wide_flag86$flag86_r <- NA
for (i in temp){
  flag_person_wide_flag86$flag86_r <- paste(flag_person_wide_flag86$flag86_r, flag_person_wide_flag86[[i]], sep = " ")
}
flag_person_wide_flag86$flag86_r <- gsub("NA ", replacement="", flag_person_wide_flag86$flag86_r)
flag_person_wide_flag86$flag86_r <- gsub(" NA", replacement="", flag_person_wide_flag86$flag86_r)

#產生檢誤報告文字
flag86_temp <- flag_person_wide_flag86 %>%
  group_by(organization_id) %>%
  mutate(flag86_txt = paste("姓名：", flag86_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag86_txt)) %>%
  distinct(organization_id, flag86_txt)

#根據organization_id，展開成寬資料(wide)
flag86 <- flag86_temp %>%
  dcast(organization_id ~ flag86_txt, value.var = "flag86_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag86)[2 : length(colnames(flag86))]
flag86$flag86 <- NA
for (i in temp){
  flag86$flag86 <- paste(flag86$flag86, flag86[[i]], sep = "； ")
}
flag86$flag86 <- gsub("NA； ", replacement="", flag86$flag86)
flag86$flag86 <- gsub("； NA", replacement="", flag86$flag86)

#產生檢誤報告文字
flag86 <- flag86 %>%
  subset(select = c(organization_id, flag86)) %>%
  distinct(organization_id, flag86) %>%
  mutate(flag86 = paste(flag86, "（經比對貴校上一學年所填資料，上述人員並未出現於本學期的教員資料表或職員(工)資料表，請確認渠等是否於本學期(111年8月1日至9月30日)退休或因故離職，若於本學期退休或因故離職，應於離退教職員(工)資料表填寫資料。如非於本學期退休或因故離職，或已介聘、調至他校，請來電告知。）", sep = ""))
}else{
#偵測flag86是否存在。若不存在，則產生NA行
if('flag86' %in% ls()){
  print("flag86")
}else{
  flag86 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag86$flag86 <- ""
}
}




# flag91: 本學期離退教職員（工）資料表比對上一學年（期）教員/職員（工）資料表，同一身分識別碼所對應的姓名不同。 -------------------------------------------------------------------
flag_person <- drev_P_retire_merge_pre %>%
  rename(name = name.x, name_pre = name.y, name_retire = name, edu_name2 = edu_name2.x)

#本學期離退教職員（工）資料表比對上一學年（期）教員/職員（工）資料表，同一身分識別碼所對應的姓名不同。
flag_person$err_flag <- 0
flag_person$err_flag <- if_else(flag_person$pre == 1 & is.na(flag_person$now) & flag_person$retire == 1 & flag_person$name_pre != flag_person$name_retire, 1, flag_person$err_flag)

#加註
flag_person$name <- paste(flag_person$name_pre, "/", flag_person$name_retire, sep = "")
flag_person$name <- gsub("；）", replacement = "）", flag_person$name)
flag_person$name <- gsub("（）", replacement = "", flag_person$name)

#呈現姓名
flag_person$err_flag_txt <- ""
flag_person$err_flag_txt <- case_when(
  flag_person$err_flag == 1 ~ flag_person$name,
  TRUE ~ flag_person$err_flag_txt
)

if (dim(flag_person %>% subset(err_flag == 1))[1] != 0){
#根據organization_id + source，展開成寬資料(wide)
flag_person_wide_flag91 <- flag_person %>%
  subset(select = c(organization_id, idnumber, err_flag_txt, edu_name2, err_flag)) %>%
  subset(err_flag == 1) %>%
  dcast(organization_id ~ err_flag_txt, value.var = "err_flag_txt")

#合併所有name
temp <- colnames(flag_person_wide_flag91)[3 : length(colnames(flag_person_wide_flag91))]
flag_person_wide_flag91$flag91_r <- NA
for (i in temp){
  flag_person_wide_flag91$flag91_r <- paste(flag_person_wide_flag91$flag91_r, flag_person_wide_flag91[[i]], sep = " ")
}
flag_person_wide_flag91$flag91_r <- gsub("NA ", replacement="", flag_person_wide_flag91$flag91_r)
flag_person_wide_flag91$flag91_r <- gsub(" NA", replacement="", flag_person_wide_flag91$flag91_r)

#產生檢誤報告文字
flag91_temp <- flag_person_wide_flag91 %>%
  group_by(organization_id) %>%
  mutate(flag91_txt = paste("請確認：", flag91_r, sep = ""), "") %>%
  subset(select = c(organization_id, flag91_txt)) %>%
  distinct(organization_id, flag91_txt)

#根據organization_id，展開成寬資料(wide)
flag91 <- flag91_temp %>%
  dcast(organization_id ~ flag91_txt, value.var = "flag91_txt")

#合併教員資料表及職員(工)資料表報告文字
temp <- colnames(flag91)[2 : length(colnames(flag91))]
flag91$flag91 <- NA
for (i in temp){
  flag91$flag91 <- paste(flag91$flag91, flag91[[i]], sep = "； ")
}
flag91$flag91 <- gsub("NA； ", replacement="", flag91$flag91)
flag91$flag91 <- gsub("； NA", replacement="", flag91$flag91)

#產生檢誤報告文字
flag91 <- flag91 %>%
  subset(select = c(organization_id, flag91)) %>%
  distinct(organization_id, flag91) %>%
  mutate(flag91 = paste(flag91, "（離退人員於上一期資料填報姓名不相同。如已更名，請來電告知）", sep = ""))
}else{
#偵測flag91是否存在。若不存在，則產生NA行
if('flag91' %in% ls()){
  print("flag91")
}else{
  flag91 <- drev_person_1 %>%
    distinct(organization_id, .keep_all = TRUE) %>%
    subset(select = c(organization_id))
  flag91$flag91 <- ""
}
}


# 教務 -----------------------------------------------------------


# 建立合併列印檔 -------------------------------------------------------------------
temp <- c("flag2", "flag3", "flag6", "flag7", "flag8", "flag9", "flag15", "flag16", "flag18", "flag19", "flag20", "flag24", "flag39", "flag45", "flag47", "flag48", "flag49", "flag50", "flag51", "flag52", "flag57", "flag59", "flag62", "flag64", "flag80", "flag82", "flag83", "flag84", "flag85", "flag86", "flag89", "flag90", "flag91", "flag92", "flag93", "flag94", "sp3", "sp5", "sp6")
check02 <- merge(x = edu_name2, y = flag1, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag2, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag3, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag6, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag7, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag8, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag9, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag15, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag16, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag18, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag19, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag20, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag24, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag39, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag45, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag47, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag48, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag49, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag50, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag51, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag52, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag57, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag59, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag62, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag64, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag80, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag82, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag83, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag84, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag85, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag86, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag89, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag90, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag91, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag92, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag93, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = flag94, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = spe3, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = spe5, by = c("organization_id"), all.x = TRUE, all.y = TRUE)
check02 <- merge(x = check02, y = spe6, by = c("organization_id"), all.x = TRUE, all.y = TRUE)

#autochecking_personnel 第168行改成      dim(teacher)[1] == 0 & dim(staff)[1] == 0
#autochecking_personnel 第182, 257行改成 readxl :: read_excel("C:\\edhr-111t1\\work\\edhr-111t1-check_print-人事.xlsx")

# 計畫端個案處理 -------------------------------------------------------------------
#國立中央大學附屬中壢高中(030305)
  #謝窘木 陳觀柏 曾庭芝 景泰秀 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "030305", "", check02$flag86)

#國立北科大附屬桃園農工(030403)
  #同等學力
  check02$spe6 <- if_else(check02$spe6 == "職員(工)資料表之大學（學士）以上各教育階段學歷資料不完整或不正確：歐奕廷（副學士或專科畢業學校（一）：同等學歷）" & check02$organization_id == "030403", "", check02$spe6)
  #林美豪 李後翰 蔡宗奮 郭金山 翁菁穗 林殷兆 許日進 呂竺昇 吳聖傑 李國釗 呂阿祝 黃華彩 蘇佳恩 房增財 黃婉婷 陳坤松  皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "030403", "", check02$flag86)

#市立龍潭高中(033302)
  #邱憶茹 吳瑞芬 黃秀娟 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "033302", "", check02$flag86)

#市立武陵高中(033306)
  #黃政揚 林雅惠 康金花 李志成 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "033306", "", check02$flag86)

#市立楊梅高中(033316)
  #陳建旺 楊孝媛 楊凱婷 王慧真 陳宏宗 張瑋秦 江美珠 黃鈴芳 黃建堉 陳逸帆 張藝馨 陳素修 羅筱恩 袁素卿 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "033316", "", check02$flag86)

#市立陽明高中(033325)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "033325", "", check02$flag80)
  #黎玉蘭 CABRERA MARLOWE CAALAMA 郭育華 朱秀容 顏中山 黃淑敏 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "033325", "", check02$flag86)

#市立中壢家商(033408)
  #陳楙開 李昱旼 黃麗芬 邱玥紜 許美倫 何淑貞 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "033408", "", check02$flag86)

#市立南崁高中(034306)
  #吳素玉 彭冠學 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "034306", "", check02$flag86)

#市立壽山高中(034314)
  #原住民姓名無誤
  check02$flag6 <- if_else(check02$flag6 == "職員(工)資料表需修改姓名處：纚固.妮卡兒" & check02$organization_id == "034314", "", check02$flag6)
  #確實沒有主（會）計室主管
  check02$flag1 <- if_else(check02$flag1 == "尚待增補之學校主管：主（會）計室主管（請確認是否填報完整名單，倘貴校###主任尚未到職，請來電告知）" & check02$organization_id == "034314", "", check02$flag1)
    #徐宗盛 苗天霞 簡瑞葶 李惠珍 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "034314", "", check02$flag86)

#市立觀音高中(034332)
  #曾麗文 劉禎芸 莊竣翔 黃順斌 張鳳玲 林玉瓊 吳宜珍 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "034332", "", check02$flag86)

#市立新屋高級中等學校(034335)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "034335", "", check02$flag80)
  #黃國峰 賀春暉 陳癸月 黃嬿瑄 蔡慧茵 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "034335", "", check02$flag86)

#市立永豐高中(034347)
  #代理教師連續聘任沒有中斷
  #check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "034347", "", check02$flag80)
  #杜冠毅 陳麗婷 林美秋 皆非本學期退休或因故離職人員
  #check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "034347", "", check02$flag86)

#市立大園國際高中(034399)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "034399", "", check02$flag80)
  #蕭辰桓 莊孟淳 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "034399", "", check02$flag86)

#國立彰化女中(070301)
  #楊明祥 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "070301", "", check02$flag86)

#國立員林高中(070304)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "070304", "", check02$flag80)

#國立彰化高中(070307)
  #呂健嘉 林胤國 林沛玲 呂興忠 柯鐙鐊 陳冠佑 施柏宇 邱建凱 陳品妤 皆非本學期退休或因故離職人員
  #check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "070307", "", check02$flag86)

#國立鹿港高中(070316)
  #職員(工)資料表：蔡奉育54歲，但學校工作總年資有38年（約16歲開始工作）這是正確的
  check02$flag39 <- if_else(check02$flag39 == "請確認該員之「本校到職日期」、「本校任職需扣除之年資」、「本校到職前學校服務總年資」，職員(工)資料表：蔡奉育54歲，但學校工作總年資有38年（約16歲開始工作）" & check02$organization_id == "070316", "", check02$flag39)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "070316", "", check02$flag80)
  #廖逵宸 林政興 李蕙菁 許笠 張雅如 陳佳麟 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "070316", "", check02$flag86)

#國立溪湖高中(070319)
  #沒有設置學程主任
  check02$flag3 <- if_else(check02$flag3 == "請學校確認是否設置學程主任" & check02$organization_id == "070319", "", check02$flag3)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "070319", "", check02$flag80)
  #林意珊 卓宏祺 洪維謙 吳旭薇 劉珮如 楊家因 林宗賢 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "070319", "", check02$flag86)

#國立彰師附工(070401)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "070401", "", check02$flag80)
  #蘇俊榮 周照棠 柯姿羽 陳凱勛 梁文瀞 洪武詮 林益瑋 吳武峰 施雄敏 林浚彬 林侑辰 梁祐祥 柯東佑 高駿雲 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "070401", "", check02$flag86)

#國立永靖高工(070402)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "070402", "", check02$flag80)
  #林欣慧 廖惠娜 李秋蘭 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "070402", "", check02$flag86)

#國立二林工商(070403)
  #無設置學程主任
  check02$flag3 <- if_else(check02$flag3 == "請學校確認是否設置學程主任" & check02$organization_id == "070403", "", check02$flag3)
  #林恒毅 黃舜牧 王子殷 黃梓彥 張玲菁 黃焜地 葉紋彬 鄭伊均 陳映軒 廖如茵 徐文振 呂宛儒 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "070403", "", check02$flag86)

#國立秀水高工(070405)
  #梁奕魁 陳淵源 施忠良 陳文瑋 林黃莉雯 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "070405", "", check02$flag86)

#國立彰化高商(070406)
  #盧諝程 黃碧雲 吳志信 何政衛 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "070406", "", check02$flag86)

#國立員林農工(070408)
  #同等學力
  check02$spe6 <- if_else(check02$spe6 == "職員(工)資料表之大學（學士）以上各教育階段學歷資料不完整或不正確：王如亮（副學士或專科畢業學校（一）：同等學力）" & check02$organization_id == "070408", "", check02$spe6)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "070408", "", check02$flag80)
  #燕美黎 蘇盈嘉 葉俊賢 梁煌坤 王怡詠 宋崇銘 黃俊偉 游麗珍 賴宥如 黃小屏 許振維 張永承 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "070408", "", check02$flag86)

#國立員林崇實高工(070409)
  #職員(工)資料表：王麗粧55歲，但學校工作總年資有40年（約15歲開始工作）這是正確的
  check02$flag39 <- if_else(check02$flag39 == "請確認該員之「本校到職日期」、「本校任職需扣除之年資」、「本校到職前學校服務總年資」，職員(工)資料表：王麗粧55歲，但學校工作總年資有40年（約15歲開始工作）" & check02$organization_id == "070409", "", check02$flag39)
  #魏延斌 許意山 洪俊林 蘇忠祥 陳玥彤 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "070409", "", check02$flag86)

#國立員林家商(070410)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "070410", "", check02$flag80)
  #黃竫茵 許庭瑜 張淑芳 蓋琪君 黃美鳳 陳定宏 孫雅玲 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "070410", "", check02$flag86)

#國立北斗家商(070415)
  #張耕濃 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "070415", "", check02$flag86)

#縣立彰化藝術高中(074308)
  #職員(工)資料表：賈慶貞57歲，但學校工作總年資有41年（約16歲開始工作）這是正確的
  check02$flag39 <- if_else(check02$flag39 == "請確認該員之「本校到職日期」、「本校任職需扣除之年資」、「本校到職前學校服務總年資」，職員(工)資料表：賈慶貞57歲，但學校工作總年資有41年（約16歲開始工作）" & check02$organization_id == "074308", "", check02$flag39)
  #李宥婕 李東祐 歐陽沐靈 鄭綜聖 陳清彬 尤建勝 蕭麗娜 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "074308", "", check02$flag86)

#縣立田中高中(074328)
  #江祥場 代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "074328", "", check02$flag80)
  #廖采頻 程惠珍 方效國 禤建國 鄭凱元 李柏萱 蕭慧儀 洪村展 艾翹楚 楊子鐳 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "074328", "", check02$flag86)

#縣立成功高中(074339)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "074339", "", check02$flag80)
  #李孟珍 林俊樂 張明仁 沈麗琴 鄭純琪 方美滿 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "074339", "", check02$flag86)

#國立中興高中(080305)
  #職員(工)資料表：李美紅58歲，但學校工作總年資有41年（約17歲開始工作） 黃鳳微64歲，但學校工作總年資有48年（約16歲開始工作）這是正確的
  check02$flag39 <- if_else(check02$flag39 == "請確認該員之「本校到職日期」、「本校任職需扣除之年資」、「本校到職前學校服務總年資」，職員(工)資料表：李美紅58歲，但學校工作總年資有41年（約17歲開始工作） 黃鳳微64歲，但學校工作總年資有48年（約16歲開始工作）" & check02$organization_id == "080305", "", check02$flag39)
  #翁廷豪 任以真 黃碧霞 紀松佐 鍾昊宏 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "080305", "", check02$flag86)

#國立竹山高中(080307)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "080307", "", check02$flag80)

#國立暨大附中(080308)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "080308", "", check02$flag80)
  #張正彥 吳毓蕙 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "080308", "", check02$flag86)

#國立埔里高工(080403)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "080403", "", check02$flag80)
  #安蘭台 陳郁文 孫兆霞 王淑珊 林秀賢 江雅如 魏廷卉 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "080403", "", check02$flag86)

#國立南投高商(080404)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "080404", "", check02$flag80)
  #陳志修 余玉鳳 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "080404", "", check02$flag86)

#國立草屯商工(080406)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "080406", "", check02$flag80)
  #姚美慧 廖大銓 黃國軒 劉沛劼 黃方伯 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "080406", "", check02$flag86)

#國立水里商工(080410)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "080410", "", check02$flag80)
  #古欣卉 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "080410", "", check02$flag86)

#縣立旭光高中(084309)
  #張玲莉 蔣計芬 黃素珍 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "084309", "", check02$flag86)

#國立臺東女中(140302)
  #李兆起 羅忠雄 陳淑惠 吳佳純 黃淑芳 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "140302", "", check02$flag86)

#國立臺東高中(140303)
  #廖珠伶 黃玲娜 宗明珊 蔡文貴 張勝發 羅青松 吳麗香 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "140303", "", check02$flag86)

#國立關山工商(140404)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "140404", "", check02$flag80)
  #黃延泰 張仁輔 莊佩璋 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "140404", "", check02$flag86)

#國立臺東高商(140405)
  #丁英桔 楊芷筠 林弘毅 曾振華 古梅花 李立方 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "140405", "", check02$flag86)

#縣立蘭嶼高中(144322)
  #確實沒有圖書館主管
  check02$flag1 <- if_else(check02$flag1 == "尚待增補之學校主管：圖書館主管（請確認是否填報完整名單，倘貴校###主任尚未到職，請來電告知）" & check02$organization_id == "144322", "", check02$flag1)
  #劉俞伶 傅桂 劉康文 鍾蘭芬 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "144322", "", check02$flag86)

#國立基隆女中(170301)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "170301", "", check02$flag80)
  #田派霖 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "170301", "", check02$flag86)

#市立安樂高中(173306)
  #高麗玲 游文毅 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "173306", "", check02$flag86)

#市立八斗高中(173314)
  #陳璿文 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "173314", "", check02$flag86)

#國立海洋大學附屬基隆海事(170403)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "170403", "", check02$flag80)
  #華志德 徐秀珠 蘇世昌 劉俊妤 劉姵妏 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "170403", "", check02$flag86)

#國立基隆商工(170404)
  #陳景堅 宋元傑 盧韻如 余沂家 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "170404", "", check02$flag86)

#市立中崙高中(313302)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "313302", "", check02$flag80)
  #陳瓊珍 陳清吉 呂育娟 上學年聘任類別填錯 應為"兼任"人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "313302", "", check02$flag86)

#市立松山高中(323301)
  #江心瑜 朱玉華 羅世焜 莊明煌 黃金雄 陶韻如 李鎂英 王筱惠 陳惠華 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "323301", "", check02$flag86)

#市立永春高中(323302)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "323302", "", check02$flag80)
  #許詠智 易台玉 上學年聘任類別填錯 應為"兼任"人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "323302", "", check02$flag86)

#市立松山家商(323401)
  #劉勇成 潘廷敏 陳彥傑 林麗華 邱雅萍 葉芝鱗 劉安妮 楊世昌 蕭雅馨 李玟頤 姚智翊 趙忠明 非本學期退休或因故離職人員
  #check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "323401", "", check02$flag86)

#市立松山工農(323402)
  #吳兆芳 尤筱潔 林怡君 陳盈如 莊昀鑫 林安琪 陳忠裕 曾家盈 江玉珍 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "323402", "", check02$flag86)

#國立師大附中(330301)
  #陳采妍 朱佩瑾 莊錦毓 林秀芬 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "330301", "", check02$flag86)

#市立和平高中(333301)
  #職員(工)資料表：潘麗卿54歲，但學校工作總年資有37年（約17歲開始工作）這是正確的
  check02$flag39 <- if_else(check02$flag39 == "請確認該員之「本校到職日期」、「本校任職需扣除之年資」、「本校到職前學校服務總年資」，職員(工)資料表：潘麗卿54歲，但學校工作總年資有37年（約17歲開始工作）" & check02$organization_id == "333301", "", check02$flag39)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "333301", "", check02$flag80)
  #林炤美 李財友 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "333301", "", check02$flag86)

#臺北市芳和實驗中學(333304)
  #該校行政單位較特別 原則上有填列各處室主管
  check02$flag1 <- if_else(check02$flag1 != "" & check02$organization_id == "333304", "", check02$flag1)
  #馮思妤 張顥瀚 宋湘媛 黃正宗 黎慧欣 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "333304", "", check02$flag86)

#市立大安高工(333401)
  #趙晋海 高添貴 孔祥璿 陳揚翔 潘虹竹 田育台 陳亮君 林俊岳 何元皓 黃柏文 楊運祥 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "333401", "", check02$flag86)

#市立中山女中(343301)
  #吳季勳 廖登山 徐志豪 陳秀鳳 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "343301", "", check02$flag86)

#市立大同高中(343302)
  #談得聖 莊靜怡 曾如玉 劉育伶 趙彥凱 莊智鈞 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "343302", "", check02$flag86)

#市立大直高中(343303)
  #張藝馨 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "343303", "", check02$flag86)

#市立建國中學(353301)
  #賴啟林 林素華 林純妙 朱軒樑 黃儷慧 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "353301", "", check02$flag86)

#市立北一女中(353303)
  #廖珈琪 上一學年聘任類別應為"約聘雇"
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "353303", "", check02$flag86)

#市立明倫高中(363301)
  #吳蕙惠 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "363301", "", check02$flag86)

#市立成淵高中(363302)
  #沈碧麗16歲開始工作無誤
  check02$flag39 <- if_else(check02$flag39 == "請確認該員之「本校到職日期」、「本校任職需扣除之年資」、「本校到職前學校服務總年資」，職員(工)資料表：沈碧麗61歲，但學校工作總年資有45年（約16歲開始工作）" & check02$organization_id == "363302", "", check02$flag39)
  #連詠順 林佳音 陳婉萍 吳淑絹 游璧僑 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "363302", "", check02$flag86)

#市立華江高中(373301)
  #林怡君 黃佳茵 羅媛馨 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "373301", "", check02$flag86)

#國立政大附中(380301)
  #盧碧英 謝欣恬 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "380301", "", check02$flag86)

#市立景美女中(383301)
  #確實沒有主（會）計室主管
  check02$flag1 <- if_else(check02$flag1 == "尚待增補之學校主管：主（會）計室主管（請確認是否填報完整名單，倘貴校###主任尚未到職，請來電告知）" & check02$organization_id == "383301", "", check02$flag1)
  #扣除年資不為零的人數確實偏高
  check02$flag64 <- if_else(check02$flag64 == "扣除年資不為零的人數偏高，請再依欄位說明確認。" & check02$organization_id == "383301", "", check02$flag64)
  #楊永祿 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "383301", "", check02$flag86)

#市立萬芳高中(383302)
  #教官 拿到退伍令 故退休
  check02$flag85 <- if_else(check02$flag85 == "教員資料表：周建賢（該員年齡似低於最低法定退休年齡，敬請再協助確認）" & check02$organization_id == "383302", "", check02$flag85)
  #陳裕隆 黃一庭 劉貞利 劉佳怡 許雅筑 陳淑伶 蘇苑瑜 陳怡廷 錢欽嫄 吳金蘭 施富祥 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "383302", "", check02$flag86)

#市立數位實驗高中(383303)
  #校長 教務處主管 學輔處主管 總務處主管 人事室主管 主（會）計室主管 皆為外校借調入 人事室沒有這些人員的人事資料 暫先讓學校通過
  check02$flag1 <- if_else(check02$flag1 == "尚待增補之學校主管：校長 教務處主管 學務處主管 總務處主管 輔導室主管 圖書館主管 實習處主管 人事室主管 主（會）計室主管（請確認是否填報完整名單，倘貴校###主任尚未到職，請來電告知）" & check02$organization_id == "383303", "", check02$flag1)
  #3位職員皆為約聘僱 暫先讓學校通過
  check02$flag18 <- if_else(check02$flag18 == "職員(工)資料表專任人員人數偏低，請再協助確認實際聘任情況，或請確認是否填報完整職員(工)名單資料。" & check02$organization_id == "383303", "", check02$flag18)

#市立木柵高工(383401)
  #陳憲章 陳秀娥 黃奕榤 鄭芷昀 蔡美季 王秀琪 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "383401", "", check02$flag86)

#市立南港高中(393301)
  #原住民姓名無誤
  check02$flag6 <- if_else(check02$flag6 == "教員資料表需修改姓名處：以柏．亞告" & check02$organization_id == "393301", "", check02$flag6)
  #林柏祥 林德明 黃苓毓 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "393301", "", check02$flag86)

#市立育成高中(393302)
  #確實沒有主（會）計室主管
  check02$flag1 <- if_else(check02$flag1 == "尚待增補之學校主管：主（會）計室主管（請確認是否填報完整名單，倘貴校###主任尚未到職，請來電告知）" & check02$organization_id == "393302", "", check02$flag1)
  #陳昱蓁 李展源 曹念儒 張政豪 張芳瑜 蔣漢旗 洪清好 林姿綿 鍾一良 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "393302", "", check02$flag86)

#市立南港高工(393401)
  #揭朝平 王上林 王錦輝 江長遠 謝荺琪 陳宗暘 何岳珈 柯炳州 吳素惠 陳泰羽 李瑄 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "393401", "", check02$flag86)

#市立內湖高中(403301)
  #原住民姓名無誤
  check02$flag6 <- if_else(check02$flag6 == "教員資料表需修改姓名處：古拉斯‧達那哈 烏巴赫‧尤紀" & check02$organization_id == "403301", "", check02$flag6)

#市立麗山高中(403302)
  #陳逸弘 邱曉淇 柯明樹 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "403302", "", check02$flag86)

#市立南湖高中(403303)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "403303", "", check02$flag80)
  #王翊軒 董家莒 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "403303", "", check02$flag86)

#市立內湖高工(403401)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "403401", "", check02$flag80)
  #錢柄匡 丁宇沛 楊秀鳳 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "403401", "", check02$flag86)

#市立陽明高中(413301)
  #林柏佑 邱嘉芳 王文玲 陳至義 羅聖涵 吳蓓欣 饒秀娟 簡淑玲 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "413301", "", check02$flag86)

#市立百齡高中(413302)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "413302", "", check02$flag80)
  #王廉今 陳依君 李宜蓁 蔡麗惠 段建強 陳玟瑾 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "413302", "", check02$flag86)

#市立士林高商(413401)
  #黃玉婷16歲開始工作無誤
  check02$flag39 <- if_else(check02$flag39 == "請確認該員之「本校到職日期」、「本校任職需扣除之年資」、「本校到職前學校服務總年資」，職員(工)資料表：黃玉婷46歲，但學校工作總年資有30年（約16歲開始工作）" & check02$organization_id == "413401", "", check02$flag39)
  #彭柏鈞 俞相榕 范筠軒 曾薇潔 黃佳淦 蔣明峰 蔣德馨 涂嘉妤 詹孟菁 施乃華 孫振華 陳美芳 姜金桂 非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "413401", "", check02$flag86)

#市立復興高中(423301)
  #閔懿 張黎明 蔡莉莉 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "423301", "", check02$flag86)

#市立中正高中(423302)
  #林憶君 王傳忠 林家豪 陳梵煦 林于雅 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "423302", "", check02$flag86)

#國立馬祖高中(720301)
  #沒有設置學程主任
  check02$flag3 <- if_else(check02$flag3 == "請學校確認是否設置學程主任" & check02$organization_id == "720301", "", check02$flag3)
  #原住民姓名無誤
  check02$flag6 <- if_else(check02$flag6 != "" & check02$organization_id == "720301", "", check02$flag6)
  #代理教師連續聘任沒有中斷
  check02$flag80 <- if_else(check02$flag80 != "" & check02$organization_id == "720301", "", check02$flag80)
  #陳怡文 陳宜道 皆非本學期退休或因故離職人員
  check02$flag86 <- if_else(check02$flag86 != "" & check02$organization_id == "720301", "", check02$flag86)



check02$err_flag <- 0

temp <- c("flag1", "flag2", "flag3", "flag6", "flag7", "flag8", "flag9", "flag15", "flag16", "flag18", "flag19", "flag20", "flag24", "flag39", "flag45", "flag47", "flag48", "flag49", "flag50", "flag51", "flag52", "flag57", "flag59", "flag62", "flag64", "flag80", "flag82", "flag83", "flag84", "flag85", "flag86", "flag89", "flag90", "flag91", "flag92", "flag93", "flag94", "spe3", "spe5", "spe6")
for (i in temp){
  check02[[i]] <- if_else(is.na(check02[[i]]), "", check02[[i]])
  check02$err_flag <- if_else(nchar(check02[[i]]) != 0, 1, check02$err_flag)
}

#刪除無錯誤的學校
check02 <- check02 %>%
  subset(err_flag != 0)

if (dim(check02 %>% subset(err_flag != 0))[1] != 0){
#標誌出無錯誤的處室
check02$err_flag_P <- 0
check02$err_flag_Ps <- 0
for (i in temp){
  check02$err_flag_P <- if_else(check02[[i]] == "", 1, check02$err_flag_P)
  check02$err_flag_Ps <- if_else(check02[[i]] != "", 1 + check02$err_flag_Ps, check02$err_flag_Ps)
}

check02$err_flag_Ps <- check02$err_flag_Ps %>% as.character()

check02$flag_P_txt <- if_else(
  check02$err_flag_P == 0, "貴處室提供的資料，沒有檢查出需要修正之處。謝謝貴處室協助完成填報工作，請等待其他處室重新上傳資料，如果處室間資料比對有誤，系統會再發信通知。謝謝！",
  paste0("經本計畫複檢，仍發現共有",  check02$err_flag_Ps,  "個可能需要修正之處，懇請貴處室協助增補，尚祈見諒！修正後的檔案需重新完成整個填報流程。如有疑問，請與本計畫人員聯繫，謝謝！")
  )

for (i in temp){
  for (j in 1:dim(check02)[1]){
    check02[[i]][j] <- if_else(check02[[i]][j] == "", "通過", check02[[i]][j])
  }
}

check02 <- check02 %>%
  subset(select = -c(err_flag, err_flag_P, err_flag_Ps))
openxlsx :: write.xlsx(check02, file = "C:\\edhr-111t1\\work\\edhr-111t1-check_print-人事.xlsx", rowNames = FALSE, overwrite = TRUE)
}else{
openxlsx :: write.xlsx(check02, file = "C:\\edhr-111t1\\work\\edhr-111t1-check_print-人事.xlsx", rowNames = FALSE, overwrite = TRUE)
}

#####自動化檢誤#####
#若全部學校皆未上傳(教員資料表及職員工資料表皆無資料)，以下皆不執行
if(dim(teacher)[1] == 0 & dim(staff)[1] == 0)
{
  print("所有學校皆尚未上傳資料")
}else
{
  #####自動化信件通知 - 每小時通知<本次上傳學校名單、需補正學校名單、三階檢通過學校名單>#####
  #初次執行需建立pre_list_agree和pre_correct_list兩個xlsx檔，用if else來做
  #若xlsx檔不存在，建立檔案
  if(!file.exists("C:/edhr-111t1/dta/edhr_111t1-202210/pre_list_agree_人事.xlsx"))
  {
    #建立pre_list_agree.xlsx
    pre_list_agree <- list_agree
    openxlsx :: write.xlsx(pre_list_agree, file = "C:/edhr-111t1/dta/edhr_111t1-202210/pre_list_agree_人事.xlsx", rowNames = FALSE, overwrite = TRUE)
    #建立pre_correct_list.xlsx
    correct_list <- readxl :: read_excel("C:\\edhr-111t1\\work\\edhr-111t1-check_print-人事.xlsx") %>% #本次需補正學校
      subset(select = c(organization_id, edu_name2))
    correct_list$edu_name2 <- paste(correct_list$edu_name2, "(", correct_list$organization_id, ")", sep = "")
    correct_list <- correct_list %>%
      mutate(pre_correct = 1)
    #以下是為了解決無法合併的問題
    if(dim(correct_list)[1] == 0)
    {
      correct_list[1, 1 : 2] = "0"
      colnames(correct_list) <- c("organization_id", "edu_name2", "pre_correct")
    }else
    {
      correct_list = correct_list
    }
    openxlsx :: write.xlsx(correct_list, file = "C:/edhr-111t1/dta/edhr_111t1-202210/pre_correct_list_人事.xlsx", rowNames = FALSE, overwrite = TRUE)
  }else
  {
    print("pre_list_agree和pre_correct_list兩個xlsx檔存在，繼續執行")
  }
  
  #若xlsx檔存在，執行
  if(file.exists("C:/edhr-111t1/dta/edhr_111t1-202210/pre_list_agree_人事.xlsx"))
  {
    #讀取上次名單
    pre_list_agree <- readxl :: read_excel("C:/edhr-111t1/dta/edhr_111t1-202210/pre_list_agree_人事.xlsx")
    pre_list_agree$organization_id <- as.character(pre_list_agree$organization_id)
    pre_correct_list <- readxl :: read_excel("C:/edhr-111t1/dta/edhr_111t1-202210/pre_correct_list_人事.xlsx")
    pre_correct_list <- mutate(pre_correct_list, pre_correct = 1)
    #以下是為了解決pre_list_agree無法合併的問題(若出現此問題只會發生在list_agree為空 且 pre_list_agree為空)
    if(dim(pre_list_agree)[1] == 0 & dim(list_agree)[1] == 0)
    {
      pre_list_agree <- list_agree
    }else
    {
      pre_list_agree = pre_list_agree
    }
    
    #本次上傳 - 本次出現但上次沒出現
    organization <- dbGetQuery(edhr, "SELECT a.id as organization_id, b.name
  FROM
  (SELECT id, max(year) as year
  FROM [plat5_edhr].[dbo].[organization_details]
  group by id) a LEFT JOIN (SELECT id, year, name
							FROM [plat5_edhr].[dbo].[organization_details]) b ON a.id = b.id AND a.year = b.year
  where (substring(a.id, 4, 1) = '1' OR substring(a.id, 4, 1) = '2' OR substring(a.id, 4, 1) = '3' OR substring(a.id, 4, 1) = '4' OR substring(a.id, 4, 1) = 'B' OR substring(a.id, 4, 1) = 'C' OR substring(a.id, 4, 1) = 'F' OR substring(a.id, 4, 1) = 'G')")
    compare_list <- left_join(list_agree, pre_list_agree, by = c("organization_id")) %>%
      subset(is.na(agree.y))
    compare_list <- merge(x = compare_list, y = organization, by = "organization_id", all.x = TRUE)
    #以下是為了解決compare_list為0
    if(dim(compare_list)[1] == 0)
    {
      compare_list[1, 1 : 4] = 0
    }else
    {
      compare_list = compare_list
    }
    compare_list$name <- paste(compare_list$name, "(", compare_list$organization_id, ")", sep = "")
    
    #本次上傳 - 本次出現且上次出現且出現在上次需修正名單(compare_correct_list的意思是在這次上傳期間未處理上次未通過的學校)
    compare_correct_list <- left_join(list_agree, pre_correct_list, by = c("organization_id")) %>%
      subset(pre_correct == 1)
    #以下是為了解決compare_correct_list為0
    if(dim(compare_correct_list)[1] == 0)
    {
      compare_correct_list[1, 1 : 4] = 0
    }else
    {
      compare_correct_list = compare_correct_list
    }
    
    #compare_correct_list$edu_name2 <- paste(compare_correct_list$edu_name2, "(", compare_correct_list$organization_id, ")", sep = "")
    #另存'本次已上傳名單"，以便於與下次名單比對
    pre_list_agree <- list_agree
    openxlsx :: write.xlsx(pre_list_agree, file = "C:/edhr-111t1/dta/edhr_111t1-202210/pre_list_agree_人事.xlsx", rowNames = FALSE, overwrite = TRUE)
    
    correct_list <- readxl :: read_excel("C:\\edhr-111t1\\work\\edhr-111t1-check_print-人事.xlsx") %>% #本次需補正學校
      subset(select = c(organization_id, edu_name2))
    correct_list_c <- correct_list %>%
      subset(select = c(organization_id))
    correct_list$edu_name2 <- paste(correct_list$edu_name2, "(", correct_list$organization_id, ")", sep = "")

    #處理correct_list為tibble的問題
    if(dim(correct_list)[1] == 0){
      correct_list <- data.frame(
      organization_id = c(""), 
      edu_name2 = c("")
    )
    correct_list <- correct_list[-1, ]
    }else{
      correct_list <- correct_list
    }
    #將correct_list_c 變數的data type改為char
    if(is.character(correct_list_c$organization_id)){
      correct_list_c <- correct_list_c
    }else{
    correct_list_c <- correct_list_c %>% mutate(across(organization_id, as.character))
    }
      correct_list <- left_join(compare_list, correct_list, by = c("organization_id")) %>%
      subset(select = c(organization_id, edu_name2)) %>%
      subset(!is.na(edu_name2))
    correct_list_2 <- left_join(correct_list_c, compare_correct_list, by = c("organization_id")) %>%
      subset(select = c(organization_id, edu_name2, pre_correct)) %>%
      subset(pre_correct == 1)
    correct_list <- bind_rows(correct_list, correct_list_2)
    correct <- apply(as.data.frame(correct_list$edu_name2), 2, paste, collapse = ", ")
    
    #用stata將學校三階檢未通過改為通過之處理
    #出現在上次需修正名單(pre_correct_list) 且未出現在本次需修正名單(correct_list) 且出現在compare_correct_list，則從compare_correct_list刪除
    #也就是我不要pre_correct_list == 1 & is.na(correct_list) & compare_correct_list == 1
    compare_correct_list <- compare_correct_list %>%
      mutate(compare_correct_list = 1)
    pre_correct_list_c <- pre_correct_list %>%
      mutate(pre_correct_list = 1)
    correct_list_c <- correct_list
    #以下是為了解決correct_list_c無法合併的問題
    if(dim(correct_list_c)[1] == 0)
    {
      correct_list_c[1, 1 : 2] = 0
      colnames(correct_list_c) <- c("organization_id", "edu_name2", "pre_correct")
    }else
    {
      correct_list_c = correct_list_c
    }
    correct_list_c <- correct_list_c %>%
      mutate(correct_list = 1)
    compare_correct_list <- merge(x = compare_correct_list, y = pre_correct_list_c, by = "organization_id", all = TRUE)
    compare_correct_list <- merge(x = compare_correct_list, y = correct_list_c, by = "organization_id", all = TRUE)
    compare_correct_list <- compare_correct_list %>%
      subset(compare_correct_list != 1 | pre_correct != 1 | !is.na(correct_list)) #By De Morgan' s Laws, (A交集B交集C)的補集合 = A補集合或B補集合或C補集合
    compare_correct_list <- compare_correct_list %>%
      subset(select = c(organization_id, agree, edu_name2.x, pre_correct.x)) %>%
      rename(edu_name2 = edu_name2.x, pre_correct = pre_correct.x)
    
    #以下是為了解決無法合併的問題
    if(dim(correct_list)[1] == 0)
    {
      correct_list[1, 1 : 2] = 0
      colnames(correct_list) <- c("organization_id", "edu_name2", "pre_correct")
    }else
    {
      correct_list = correct_list
    }
    
    #以下是為了解決correct開頭為","
    str_corr <- str_locate(correct, ",")[ ,1]
    
    if(is.na(str_corr))
    {
      str_corr = " "
    }else
    {
      str_corr = str_corr
    }
    
    if(str_corr == 1)
    {
      correct = substr(correct, start = 2, stop = nchar(correct))  
    }else
    {
      correct = correct
    }
    
    #建立信件內容會用到的學校名單
    now <- apply(as.data.frame(compare_list$name), 2, paste, collapse = ", ") #本次上傳學校
    #以下是為了解決now為0(0)
    if(now == "0(0)")
    {
      now = ""
    }else
    {
      now = now
    }
    now <- paste(now, apply(as.data.frame(compare_correct_list$edu_name2), 2, paste, collapse = ", "), sep = ", ")
    
    #以下是為了解決now開頭為","
    str_now <- str_locate(now, ",")[ ,1]
    
    if(is.na(str_now))
    {
      str_now = " "
    }else
    {
      str_now = str_now
    }
    
    if(str_now == 1)
    {
      now = substr(now, start = 2, stop = nchar(now))  
    }else
    {
      now = now
    }
    #以下是為了解決now為0
    if(now == "0")
    {
      now = ""
    }else
    {
      now = now
    }
    
    #以下是為了解決now為 0
    if(now == " 0")
    {
      now = ""
    }else
    {
      now = now
    }
    
    #以下是為了解決now結尾為", 0"
    str_now <- str_locate(now, ", 0")[ ,1]
    
    if(!is.na(str_now) & now != "")
    {
      now = substr(now, start = 1, stop = nchar(now) - 3)
    }else
    {
      now = now
    }
    
    #以下是為了解決now結尾為", NA"
    if(is.na(str_locate(now, ", NA")[ ,1])){
      now = now
    }else if(str_locate(now, ", NA")[ ,1] == nchar(now) - 3){
      now = substr(now, start = 1, stop = nchar(now) - 4)
    }else{
      now = now
    }
    
    #另存'本次需修正名單"，以便於與下次名單比對
    openxlsx :: write.xlsx(correct_list, file = "C:/edhr-111t1/dta/edhr_111t1-202210/pre_correct_list_人事.xlsx", rowNames = FALSE, overwrite = TRUE)
    
    clear_list <- left_join(compare_list, correct_list, by = c("organization_id")) %>% #本次三階檢通過學校
      subset(select = c(organization_id, edu_name2, name)) %>%
      subset(is.na(edu_name2)) 
    clear <-apply(as.data.frame(clear_list$name), 2, paste, collapse = ", ")
    clear_correct_list <- merge(x = correct_list, y = pre_correct_list, by = c("organization_id"), all = TRUE) %>%
      subset(is.na(edu_name2.x)) %>%
      subset(select = c(organization_id, edu_name2.y))
    clear_correct_list <- merge(x = clear_correct_list , y = pre_correct_list, by = c("organization_id"), all.x = TRUE) %>%  #clear_correct_list: 沒出現在correct_list 且出現在pre_correct_list，可能為(1)本次通過且上次未通過 或(2)本次被退件且上次未通過，需排除(2)，也就是clear_correct_list的名單若也出現在pre_correct_list，需排除
      subset(is.na(edu_name2)) %>%
      subset(select = c(organization_id, edu_name2.y))
    #以下是為了解決clear_correct_list為0
    if(dim(clear_correct_list)[1] == 0)
    {
      clear_correct_list[1, 1 : 2] = 0
    }else
    {
      clear_correct_list = clear_correct_list
    }
    
    clear_correct_list$edu_name2.y <- substr(clear_correct_list$edu_name2.y, start = 1, stop = str_locate(clear_correct_list$edu_name2.y, pattern = "\\(")[1, 1] - 1)
    clear_correct_list$edu_name2.y <- paste(clear_correct_list$edu_name2.y, "(", clear_correct_list$organization_id, ")", sep = "")
    clear_2 <-apply(as.data.frame(clear_correct_list$edu_name2.y), 2, paste, collapse = ", ")
    clear <- paste(clear, clear_2, sep = ",")
    
    #以下是為了解決clear開頭為","
    if(str_locate(clear, ",")[ ,1] == 1)
    {
      clear = substr(clear, start = 2, stop = nchar(clear))  
    }else
    {
      clear = clear
    }
    
    #以下是為了解決"0(0)"
    if(now == "0(0)")
    {
      now = ""
    }else
    {
      now = now
    }
    
    #以下是為了解決now開頭為" "
    # if (is.na(now)) {
    #   now = now
    # } else if (str_locate(now, " ")[ ,1] == 1 & now != "") {
    #   now = substr(now, start = 2, stop = nchar(now))  
    # } else {
    #   now = now
    # }
    
    #以下是為了解決now中間出現NA
    now <- gsub(", NA", "", now)
    
    if(correct == "0(0)")
    {
      correct = ""
    }else
    {
      correct = correct
    }
    
    #以下是為了解決"NA(0)"
    if(clear == "NA(0)")
    {
      clear = ""
    }else
    {
      clear = clear
    }
    
    #以下是為了解決clear出現,NA(0)
    clear <- gsub(",NA\\(0\\)", "", clear)
    
    #以下是為了解決clear為"0(0)"
    if(clear == "0(0)")
    {
      clear = ""
    }else
    {
      clear = clear
    }
    
    #以下是為了解決now為" NA"
    now <- gsub(" NA", "", now)
    
    #寄信
    #先判斷check_print檔案是否使用中(以"是否可更改檔案名稱"來判斷 若可更改 代表未使用中)
    checkprint_filename <- "C:/edhr-111t1/work/edhr-111t1-check_print-人事.xls"
    checkprint_filename_2 <- substr(checkprint_filename, start = 1, stop = str_locate(checkprint_filename, ".xls")[ ,1] - 1)  
    
    if(file.rename(from = checkprint_filename, to = paste(checkprint_filename_2, "2.xls", sep = "")) == TRUE)
    {
      if(nchar(now) == 0)
      {
        file.rename(from = paste(checkprint_filename_2, "2.xls", sep = ""), to = checkprint_filename)
        
        paste(format(time_now, format = "%Y/%m/%d %H:%M"), " 本次無學校上傳", sep = "")
      }else
      {
        #存入xlsx，自動開啟
        #建立檔案名稱
        correct_filename_year <- substr(title, start = str_locate(title, "學年度")[ ,1] - 3, stop = str_locate(title, "學年度")[ ,1] - 1)
        if(substr(title, start = str_locate(title, "學期")[ ,1] - 1, stop = str_locate(title, "學期")[ ,1] - 1) == "上")
        {
          correct_filename_sem <- "1"
        }else{
          correct_filename_sem <- "2"
        }
        correct_filename_name <- substr(title, start = str_locate(title, "（")[ ,1] + 1, stop = str_locate(title, "）")[ ,1] - 1)
        correct_filename <- paste(correct_filename_year, correct_filename_sem, correct_filename_name, "_上傳名單", sep = "")
        
        #建立fileopen.bat
        write.table(paste("start C:\\autochecking\\",correct_filename, ".xlsx", sep = ""), file = "C:\\autochecking\\fileopen.bat", append = FALSE, quote = FALSE, col.names = FALSE, row.names = FALSE, fileEncoding = "BIG5")
        
        if(!file.exists(paste("C:\\autochecking\\",correct_filename, ".xlsx", sep = "")))
        {
          #如果檔案不存在
          # 建立 Excel 活頁簿
          wb <- createWorkbook()
          
          # 設定框線樣式
          options("openxlsx.borderColour" = "#4F80BD")
          options("openxlsx.borderStyle" = "thin")
          
          # 設定 Excel 活頁簿預設字型
          modifyBaseFont(wb, fontSize = 20, fontName = "Arial")
          
          # 新增工作表
          addWorksheet(wb, sheetName = "上傳名單", gridLines = FALSE)
          
          # 建立上傳學校名單表格
          body <- data.frame(matrix(0, 1, 4))
          colnames(body) <- c("上傳時間", "本次上傳學校", "本次需補正學校", "本次三階檢通過學校")
          body[1, ] <- c(format(time_now, format = "%Y/%m/%d %H:%M"), now, correct, clear)
          
          # 建立樣式
          headSty <- createStyle(fontSize = 22, fgFill="#DCE6F1", halign="center", border = "TopBottomLeftRight", wrapText = TRUE)
          
          # 將學校名單表格寫入
          txtSty <- createStyle(halign="left", valign = "center", border = "TopBottomLeftRight", wrapText = TRUE)
          writeData(wb, 1, x = body, startCol = "A", startRow=1, borders="rows", headerStyle = headSty)
          addStyle(wb, sheet = 1, style = txtSty, cols = 1:4, rows = 2:(dim(body)[1]+1), gridExpand = TRUE)
          
          # 設定欄寬
          setColWidths(wb, 1, cols=1, widths = 16)
          setColWidths(wb, 1, cols=2:5, widths = 20)
          
          # 儲存 Excel 活頁簿
          saveWorkbook(wb, paste("C:\\autochecking\\",correct_filename, ".xlsx", sep = ""), overwrite = TRUE)
          
          # excel檔開啟30秒後自動關閉
          time_a <- Sys.time()
          a <- as.numeric(format(time_a, format = "%M")) * 60 + as.numeric(format(time_a, format = "%S"))      
          shell.exec("C:\\autochecking\\fileopen.bat")
          
          b <- a
          while (b - a < 30)
          {
            time_b <- Sys.time()
            b <- as.numeric(format(time_b, format = "%M")) * 60 + as.numeric(format(time_b, format = "%S"))
          }
          
          if(!file.exists("C:\\autochecking\\fileclose.bat"))
          {
            write.table(paste("taskkill /FI \"WINDOWTITLE eq ", correct_filename, "*\"", sep = ""), file = "C:\\autochecking\\fileclose.bat", append = FALSE, quote = FALSE, col.names = FALSE, row.names = FALSE)
          }else
          {
            print("建立fileclose.bat檔案")
          }
          
          shell.exec("C:\\autochecking\\fileclose.bat")
        }else{
          #如果檔案存在
          # 建立 Excel 活頁簿
          wb <- createWorkbook()
          
          # 設定框線樣式
          options("openxlsx.borderColour" = "#4F80BD")
          options("openxlsx.borderStyle" = "thin")
          
          # 設定 Excel 活頁簿預設字型
          modifyBaseFont(wb, fontSize = 20, fontName = "Arial")
          
          # 新增工作表
          addWorksheet(wb, sheetName = "上傳名單", gridLines = FALSE)
          
          # 建立上傳學校名單表格
          body <- readxl :: read_excel(paste("C:\\autochecking\\",correct_filename, ".xlsx", sep = ""))
          body <- rbind(c("0", "0", "0", "0"), body)
          colnames(body) <- c("上傳時間", "本次上傳學校", "本次需補正學校", "本次三階檢通過學校")
          body[1, ] <- as.list(c(format(time_now, format = "%Y/%m/%d %H:%M"), now, correct, clear))
          
          # 建立樣式
          headSty <- createStyle(fontSize = 22, fgFill="#DCE6F1", halign="center", border = "TopBottomLeftRight", wrapText = TRUE)
          
          # 將學校名單表格寫入
          txtSty <- createStyle(halign="left", valign = "center", border = "TopBottomLeftRight", wrapText = TRUE)
          writeData(wb, 1, x = body, startCol = "A", startRow=1, borders="rows", headerStyle = headSty)
          addStyle(wb, sheet = 1, style = txtSty, cols = 1:4, rows = 2:(dim(body)[1]+1), gridExpand = TRUE)
          
          # 設定欄寬
          setColWidths(wb, 1, cols=1, widths = 16)
          setColWidths(wb, 1, cols=2:5, widths = 20)
          
          # 儲存 Excel 活頁簿
          saveWorkbook(wb, paste("C:\\autochecking\\",correct_filename, ".xlsx", sep = ""), overwrite = TRUE)
          
          # excel檔開啟30秒後自動關閉
          time_a <- Sys.time()
          a <- as.numeric(format(time_a, format = "%M")) * 60 + as.numeric(format(time_a, format = "%S"))      
          shell.exec("C:\\autochecking\\fileopen.bat")
          
          b <- a
          while (b - a < 30)
          {
            time_b <- Sys.time()
            b <- as.numeric(format(time_b, format = "%M")) * 60 + as.numeric(format(time_b, format = "%S"))
          }
          
          write.table(paste("taskkill /FI \"WINDOWTITLE eq ", correct_filename, "*\"", sep = ""), file = "C:\\autochecking\\fileclose.bat", append = FALSE, quote = FALSE, col.names = FALSE, row.names = FALSE, fileEncoding = "BIG5")
          
          shell.exec("C:\\autochecking\\fileclose.bat")
          
          file.rename(from = "C:\\edhr-111t1\\work\\edhr-111t1-check_print-人事2.xls", to = "C:\\edhr-111t1\\work\\edhr-111t1-check_print-人事.xls")
        }
      }
    }else
    {
      #建立errortext_fileopen.bat
      if(!file.exists("C:\\autochecking\\errortext_fileopen.bat"))
      {
        write.table("start C:\\autochecking\\errortext_fileopen.xlsx", file = "C:\\autochecking\\errortext_fileopen.bat", append = FALSE, quote = FALSE, col.names = FALSE, row.names = FALSE)
        print("建立errortext_fileopen.bat檔案")
      }else
      {
        print("errortext_fileopen.bat檔案存在")
      }
      
      #如果檔案不存在
      # 建立 Excel 活頁簿
      wb <- createWorkbook()
      
      # 設定框線樣式
      options("openxlsx.borderColour" = "#4F80BD")
      options("openxlsx.borderStyle" = "thin")
      
      # 設定 Excel 活頁簿預設字型
      modifyBaseFont(wb, fontSize = 20, fontName = "Arial")
      
      # 新增工作表
      addWorksheet(wb, sheetName = "上傳名單", gridLines = FALSE)
      
      # 建立上傳學校名單表格
      body <- "本次自動化檢核未執行，請盡速關閉檢核報告word檔，自動化檢核方可繼續執行。閱讀完畢請關閉此檔案。"
      
      # 建立樣式
      headSty <- createStyle(fontSize = 22, fgFill="#DCE6F1", halign="center", border = "TopBottomLeftRight", wrapText = TRUE)
      
      # 將學校名單表格寫入
      txtSty <- createStyle(halign="left", valign = "center", border = "TopBottomLeftRight", wrapText = TRUE)
      writeData(wb, 1, x = body, startCol = "A", startRow=1, borders="rows", headerStyle = headSty)
      addStyle(wb, sheet = 1, style = txtSty, cols = 1:5, rows = 1, gridExpand = TRUE)
      mergeCells(wb, sheet = 1, cols = 1:5, rows = 1:1)
      
      # 設定欄寬
      setColWidths(wb, 1, cols=1, widths = 16)
      setColWidths(wb, 1, cols=2:5, widths = 20)
      
      # 儲存 Excel 活頁簿
      saveWorkbook(wb, "C:\\autochecking\\errortext_fileopen.xlsx", overwrite = TRUE)
      shell.exec("C:\\autochecking\\errortext_fileopen.bat")
    }
  }else
  {
    print("pre_list_agree和pre_correct_list兩個xlsx檔不存在")
  }
}