#
#-------------------------------------------------------------------------------
# Env setup
#-------------------------------------------------------------------------------
#

source("~/amana/.Rprofile")

StartTime <- Sys.time() 

## Activate python venv
# use_virtualenv(virtualenv = .PYTHON_VENV_DIR, required = TRUE)    ## CPython
# use_condaenv("gmkt", required = TRUE)                           ## Anaconda
# reticulate::source_python(paste0(.MAIN_DIR, '/.codes/tools_others.py'))


## Source Utils.

source(paste0(.MAIN_DIR, '.rcodes/pv_agents/etls.R'))
source(paste0(.MAIN_DIR, '.rcodes/pv_agents/tools_duplicates.R'))
reticulate::source_python(paste0(.MAIN_DIR, '.rcodes/pv_agents/tools_others.py'))


#
#-------------------------------------------------------------------------------
# Process
#-------------------------------------------------------------------------------
#

## Report month
report_month <- .monthFirstDate(read_excel(paste0(.MAIN_DIR, 'Inputs.xlsm'), sheet = "Update") %>% pull(MOIS) %>% max())


## Clean Data
.preprocessSysRAW(month_ = report_month)


## Update com DB
reticulate::source_python(paste0(.MAIN_DIR, 'update_agent_db.py'))


## Variations vs mois précédents
.globalVariations(month_ = report_month)
tmp <- .loadTotals(month_ = report_month)
global_real <- tmp$real
global_obj <- tmp$obj
rm(tmp);gc()

#
#---------------------------------------------
#
#  Duplicates
#
#---------------------------------------------
#

dup_agent <- .confirmedDuplicates(month_ = report_month, info_ = "Agent", threshold = 90)
dup_agent %>% select(RAW) %>% distinct() %>% nrow() - nrow(dup_agent)

if(nrow(dup_agent) > 0) {
  process_dup <- TRUE
} else {
  process_dup <- FALSE
}

#
#---------------------------------------------
#
#  Objectifs
#
#---------------------------------------------
#

# obj_agents <- .loadObj(month_ = report_month, dup_agent_list = list(process = TRUE, dup_df = dup_agent))
obj_agents <- .loadObj(month_ = report_month, dup_agent_list = list(process = process_dup, dup_df = dup_agent))
obj_agents <- .alterArchi(obj_agents) %>%
  filter(OBJ_NBRE_ENVOIS > 0 | OBJ_FRAIS_ENVOIS > 0 | OBJ_MONTANT_RECHARGES > 0)

glimpse(obj_agents)
any(is.na(obj_agents))
obj_agents %>% select(AGENT, AGENCE, COORDONATEUR, PAYS_COORDONATEUR) %>% distinct() %>% nrow() - nrow(obj_agents)

agg_obj_agents <- obj_agents %>%
  group_by(AGENT, AGENCE, COORDONATEUR, PAYS_COORDONATEUR) %>%
  summarise(
    .groups = "drop",
    OBJ_NBRE_ENVOIS = sum(OBJ_NBRE_ENVOIS),
    OBJ_FRAIS_ENVOIS = sum(OBJ_FRAIS_ENVOIS),
    OBJ_MONTANT_RECHARGES = sum(OBJ_MONTANT_RECHARGES)
  ) %>%
  ungroup() %>%
  filter(OBJ_NBRE_ENVOIS > 0 | OBJ_FRAIS_ENVOIS > 0 | OBJ_MONTANT_RECHARGES > 0)


obj_check_msg <- paste0(
  "Chargement des objectifs de ", format(report_month, "%b-%Y"), ":\n",
  "Nombre d'envois            | ", if_else(sum(obj_agents$OBJ_NBRE_ENVOIS, na.rm = TRUE) == global_obj$nbre_envois, "OK", "NOK"), "\n",
  "Frais d'envois             | ", if_else(sum(obj_agents$OBJ_FRAIS_ENVOIS, na.rm = TRUE) == global_obj$frais_envois, "OK", "NOK"), "\n",
  "Montant des recharges      | ", if_else(sum(obj_agents$OBJ_MONTANT_RECHARGES, na.rm = TRUE) == global_obj$montant_recharges, "OK", "NOK"), 
  "\n\n"
)

cat(obj_check_msg)

prettyNum(round(sum(obj_agents$OBJ_NBRE_ENVOIS),0), big.mark = " ")
prettyNum(round(sum(obj_agents$OBJ_FRAIS_ENVOIS),0), big.mark = " ")
prettyNum(round(sum(obj_agents$OBJ_MONTANT_RECHARGES),0), big.mark = " ")

#
#---------------------------------------------
#
#  Realisations
#
#---------------------------------------------
#


real_list <- .loadReals(month_ = report_month, dup_agent_list = list(process = process_dup, dup_df = dup_agent))

## Envois Mois M
envois_agents <- real_list$envois
envois_agents <- .alterArchi(envois_agents) %>% filter(NBRE_ENVOIS > 0 | FRAIS_ENVOIS > 0)

glimpse(envois_agents)
any(is.na(envois_agents))
envois_agents %>% select(AGENT, PDV, AGENCE) %>% distinct() %>% nrow() - nrow(envois_agents)

agg_envois_agents <- envois_agents %>%
  group_by(AGENT, AGENCE, COORDONATEUR, PAYS_COORDONATEUR) %>%
  summarise(
    .groups = "drop",
    NBRE_ENVOIS = sum(NBRE_ENVOIS),
    FRAIS_ENVOIS = sum(FRAIS_ENVOIS)
  ) %>%
  ungroup() %>%
  filter(NBRE_ENVOIS > 0 | FRAIS_ENVOIS > 0)


## Envois Mois M-1
report_month_mm1 <- .monthFirstDate(.monthFirstDate(report_month) - 1)
envois_agents_mm1 <- .loadReals(month_ = report_month_mm1, dup_agent_list = list(process = process_dup, dup_df = dup_agent))$envois
envois_agents_mm1 <- .alterArchi(envois_agents_mm1) %>%
  rename(NBRE_ENVOIS_MM1 = NBRE_ENVOIS, FRAIS_ENVOIS_MM1 = FRAIS_ENVOIS) %>%
  filter(NBRE_ENVOIS_MM1 > 0 | FRAIS_ENVOIS_MM1 > 0)

glimpse(envois_agents_mm1)
any(is.na(envois_agents_mm1))
envois_agents_mm1 %>% select(AGENT, PDV, AGENCE) %>% distinct() %>% nrow() - nrow(envois_agents_mm1)

agg_envois_agents_mm1 <- envois_agents_mm1 %>%
  group_by(AGENT, AGENCE, COORDONATEUR, PAYS_COORDONATEUR) %>%
  summarise(
    .groups = "drop",
    NBRE_ENVOIS_MM1 = sum(NBRE_ENVOIS_MM1),
    FRAIS_ENVOIS_MM1 = sum(FRAIS_ENVOIS_MM1)
  ) %>%
  ungroup() %>%
  filter(NBRE_ENVOIS_MM1 > 0 | FRAIS_ENVOIS_MM1 > 0)



## Recharges
recharges_agents <- real_list$recharges
recharges_agents <- .splitRecharges(month_ = report_month, recharges_df = recharges_agents)
recharges_agents <- .alterArchi(recharges_agents) %>% filter(MONTANT_RECHARGES > 0)

glimpse(recharges_agents)
any(is.na(recharges_agents))
recharges_agents %>% select(AGENT, PDV, AGENCE) %>% distinct() %>% nrow() - nrow(recharges_agents)


agg_recharges_agents <- recharges_agents %>%
  group_by(AGENT, AGENCE, COORDONATEUR, PAYS_COORDONATEUR) %>%
  summarise(
    .groups = "drop",
    MONTANT_RECHARGES_AGENT = sum(MONTANT_RECHARGES_AGENT),
    MONTANT_RECHARGES_COM = sum(MONTANT_RECHARGES_COM),
    MONTANT_RECHARGES_DISTRI = sum(MONTANT_RECHARGES_DISTRI),
    MONTANT_RECHARGES = sum(MONTANT_RECHARGES)
  ) %>%
  ungroup() %>%
  filter(MONTANT_RECHARGES > 0)



## Retraits
retraits_agents <- real_list$retraits

agence_coord <- envois_agents %>%
  select(AGENCE, COORDINATION, PAYS) %>%
  distinct() %>% 
  mutate(RANK = 1) %>% 
  bind_rows(recharges_agents %>% select(AGENCE, COORDINATION, PAYS) %>% distinct() %>% mutate(RANK = 2)) %>% 
  group_by(AGENCE, COORDINATION, PAYS) %>%
  mutate(RANK2 = row_number(RANK)) %>%
  ungroup() %>%
  filter(RANK2 == 1) %>%
  select(-c(RANK, RANK2))


retraits_agents <- retraits_agents %>%
  left_join(agence_coord, by = c("AGENCE" = "AGENCE", "PAYS" = "PAYS")) %>%
  replace_na(replace = list(COORDINATION = "_._"))

retraits_agents <- .alterArchi(retraits_agents) %>% filter(NBRE_RETRAITS > 0)

glimpse(retraits_agents)
any(is.na(retraits_agents))
retraits_agents %>% select(AGENT, PDV, AGENCE) %>% distinct() %>% nrow() - nrow(retraits_agents)


agg_retraits_agents <- retraits_agents %>%
  group_by(AGENT, AGENCE, COORDONATEUR, PAYS_COORDONATEUR) %>%
  summarise(
    .groups = "drop",
    NBRE_RETRAITS = sum(NBRE_RETRAITS)
  ) %>%
  ungroup() %>%
  filter(NBRE_RETRAITS > 0)


real_check_msg <- paste0(
  "Chargement des réalisations de ", format(report_month, "%b-%Y"), ":\n",
  "Nombre d'envois            | ", if_else(sum(envois_agents$NBRE_ENVOIS, na.rm = TRUE) == global_real$nbre_envois, "OK", "NOK"), "\n",
  "Frais d'envois             | ", if_else(sum(envois_agents$FRAIS_ENVOIS, na.rm = TRUE) == global_real$frais_envois, "OK", "NOK"), "\n",
  "Nombre de retraits         | ", if_else(sum(retraits_agents$NBRE_RETRAITS, na.rm = TRUE) == global_real$nbre_retraits, "OK", "NOK"), "\n",
  "Montant des recharges      | ", if_else(sum(recharges_agents$MONTANT_RECHARGES , na.rm = TRUE) == global_real$montant_recharges, "OK", "NOK"), 
  "\n\n"
)

cat(real_check_msg)



## Architecture
jointure_ <- c(
  'AGENT' = 'AGENT',
  'AGENCE' = 'AGENCE',
  'COORDONATEUR' = 'COORDONATEUR',
  'PAYS_COORDONATEUR' = 'PAYS_COORDONATEUR'
)

cols_ <- c(
  "AGENT",
  "PDV",
  "AGENCE",
  "COORDONATEUR",
  "PAYS_COORDONATEUR"
)

old_arch <- obj_agents %>%
  filter(OBJ_NBRE_ENVOIS > 0 | OBJ_FRAIS_ENVOIS > 0 | OBJ_MONTANT_RECHARGES > 0) %>% 
  select(AGENT, PDVs, AGENCE, COORDONATEUR, PAYS_COORDONATEUR) %>%
  distinct() %>%
  separate_rows(PDVs, sep = "; ") %>%
  rename(PDV = PDVs) %>%
  mutate(
    # PDV = sapply(PDV, .manageSpetialCar),
    PDV = str_to_upper(string = PDV),
    PDV = str_squish(string = PDV)
  ) %>% 
  filter(!is.na(PDV))

any(is.na(old_arch))


global_ref <- old_arch %>% select(all_of(cols_)) %>% 
  bind_rows(envois_agents %>% select(all_of(cols_)) %>% distinct()) %>%
  bind_rows(envois_agents_mm1 %>% select(all_of(cols_)) %>% distinct()) %>%
  bind_rows(retraits_agents %>% select(all_of(cols_)) %>% distinct()) %>%
  bind_rows(retraits_agents %>% select(all_of(cols_)) %>% distinct()) %>%
  distinct() %>%
  arrange(PAYS_COORDONATEUR, COORDONATEUR, AGENCE, AGENT, PDV) %>%
  group_by(AGENT, AGENCE, COORDONATEUR, PAYS_COORDONATEUR) %>%
  summarise(
    .groups = "drop",
    PDVs = as.character(knitr::combine_words(PDV, sep = "; ", and = ""))
  ) %>%
  ungroup() %>%
  select(AGENT, PDVs, AGENCE, COORDONATEUR, PAYS_COORDONATEUR) %>%
  mutate(
    PDVs = str_replace(string = PDVs, pattern = "; _._", replacement = ""),
    PDVs = str_replace(string = PDVs, pattern = "^_._; ", replacement = ""),
    PDVs = str_replace(string = PDVs, pattern = "_._", replacement = "")
  ) %>%
  arrange(PAYS_COORDONATEUR, COORDONATEUR, AGENCE, AGENT)




global_ref <- old_arch %>% select(all_of(cols_)) %>% 
  bind_rows(envois_agents %>% select(all_of(cols_)) %>% distinct()) %>%
  bind_rows(envois_agents_mm1 %>% select(all_of(cols_)) %>% distinct()) %>%
  bind_rows(retraits_agents %>% select(all_of(cols_)) %>% distinct()) %>%
  bind_rows(retraits_agents %>% select(all_of(cols_)) %>% distinct()) %>%
  distinct() %>%
  arrange(PAYS_COORDONATEUR, COORDONATEUR, AGENCE, AGENT, PDV) %>%
  group_by(AGENT, AGENCE, COORDONATEUR, PAYS_COORDONATEUR) %>%
  mutate(rank = row_number(PDV)) %>%
  ungroup() %>%
  filter(rank == 1) %>%
  mutate(MAT = "-") %>% 
  select(
    MAT,
    AGENT,
    PDV,
    AGENCE,
    COORD = COORDONATEUR,
    PAYS = PAYS_COORDONATEUR
  ) %>%
  mutate(
    AGENT = str_squish(string = AGENT),
    PDV = str_squish(string = PDV),
    AGENCE = str_squish(string = AGENCE),
    COORD = str_squish(string = COORD),
    PAYS = str_squish(string = PAYS)
  ) %>%
  distinct() %>%
  arrange(PAYS, COORD, AGENCE, PDV) %>%
  rowid_to_column(var = 'tmp') %>% 
  mutate(MAT = paste0('M2023', str_pad(tmp, 3, pad = "0"))) %>%
  select(-c(tmp)) %>%
  # Coord
  group_by(COORD) %>%
  mutate(rank_coord = row_number(AGENT)) %>%
  ungroup() %>%
  # Agence
  group_by(AGENCE) %>%
  mutate(rank_agence = row_number(AGENT)) %>%
  ungroup() %>%
  mutate(CATEGORIE = case_when(
    rank_coord == 1 ~ "Coordonateur",
    rank_coord == 2 & str_detect(string = str_to_lower(string = COORD), pattern = "niamey") ~ "Coordonateur adjoint",
    rank_agence == 1 ~ "Chef d'agence",
    TRUE ~ "Agent"
  )) %>%
  select(-c(rank_agence, rank_coord)) %>%
  mutate(SALAIRE = case_when(
    CATEGORIE == "Coordonateur" ~ 550000.0,
    CATEGORIE == "Coordonateur adjoint" ~ 250000.0,
    CATEGORIE == "Chef d'agence" ~ 150000.0,
    CATEGORIE == "Agent" ~ 75000.0,
    TRUE ~ 0.0
  ))



hire_date <- function(item) {
  dates_ <- seq(from=as.Date("2021-01-01"), to=as.Date("2022-12-31"), by="1 day")
  return(sample(dates_, 1))
}

departure_date <- function(item) {
  dates_ <- seq(from=as.Date("2021-01-01"), to=as.Date("2022-12-31"), by="1 day")
  return(sample(dates_, 1))
}



employees <- global_ref %>%
  mutate(
    DATE_EMBAUCHE = sapply(AGENT, hire_date),
    DATE_EMBAUCHE = as.Date(DATE_EMBAUCHE, origin = "1970-01-01")
  ) %>%
  mutate(BONUS = 0)


coordonateurs <- employees %>%
  filter(CATEGORIE == "Coordonateur") %>%
  select(MAT, EMPLOYEE=AGENT, CATEGORIE, DATE_EMBAUCHE, SALAIRE, BONUS, COORD)


coordonateurs_adjoints <- employees %>%
  filter(CATEGORIE == "Coordonateur adjoint") %>%
  select(MAT, EMPLOYEE=AGENT, CATEGORIE, DATE_EMBAUCHE, SALAIRE, BONUS, COORD)

chefs_agences <- employees %>%
  filter(CATEGORIE == "Chef d'agence") %>%
  select(MAT, EMPLOYEE=AGENT, CATEGORIE, DATE_EMBAUCHE, SALAIRE, BONUS, AGENCE)


agents <- employees %>% 
  filter(CATEGORIE == "Agent") %>%
  select(MAT, EMPLOYEE=AGENT, CATEGORIE, DATE_EMBAUCHE, SALAIRE, BONUS, PDV)


#----------------


pays <- employees %>% 
  select(PAYS) %>%
  distinct() %>%
  arrange(PAYS)

coords <- employees %>% 
  select(
    COORD,
    PAYS
  ) %>%
  distinct() %>%
  arrange(PAYS, COORD)


agences <- employees %>% 
  select(
    AGENCE,
    COORD,
    PAYS
  ) %>%
  distinct() %>%
  arrange(PAYS, COORD, AGENCE)


pdvs <- employees %>% 
  select(
    PDV,
    AGENCE,
    COORD,
    PAYS
  ) %>%
  distinct() %>%
  arrange(PAYS, COORD, AGENCE, PDV)


#----------------

status <- read_excel('C:/TravauxR/pyCharm/local/rh/data/rh_conf.xlsx', sheet="Status")

levels <- read_excel('C:/TravauxR/pyCharm/local/rh/data/rh_conf.xlsx', sheet="Level")

#----------------


fake_db_path <- 'C:/TravauxR/pyCharm/local/rh/data/fake_db.xlsx'
wb <- createWorkbook(fake_db_path)

## Status
sheet_name <- "Status"
addWorksheet(wb, sheet_name)
df <- status

writeDataTable(
  wb,sheet=sheet_name, x=df, 
  startCol = 1, startRow = 1, xy = NULL,
  colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleLight9",
  tableName = NULL, headerStyle = NULL, withFilter = TRUE,
  keepNA = FALSE, sep = ", ", stack = FALSE, firstColumn = FALSE,
  lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE
)


## Levels
sheet_name <- "Levels"
addWorksheet(wb, sheet_name)
df <- levels

writeDataTable(
  wb,sheet=sheet_name, x=df, 
  startCol = 1, startRow = 1, xy = NULL,
  colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleLight9",
  tableName = NULL, headerStyle = NULL, withFilter = TRUE,
  keepNA = FALSE, sep = ", ", stack = FALSE, firstColumn = FALSE,
  lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE
)


## Pays
sheet_name <- "Pays"
addWorksheet(wb, sheet_name)
df <- pays

writeDataTable(
  wb,sheet=sheet_name, x=df, 
  startCol = 1, startRow = 1, xy = NULL,
  colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleLight9",
  tableName = NULL, headerStyle = NULL, withFilter = TRUE,
  keepNA = FALSE, sep = ", ", stack = FALSE, firstColumn = FALSE,
  lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE
)


## Coordinations
sheet_name <- "Coordinations"
addWorksheet(wb, sheet_name)
df <- coords

writeDataTable(
  wb,sheet=sheet_name, x=df, 
  startCol = 1, startRow = 1, xy = NULL,
  colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleLight9",
  tableName = NULL, headerStyle = NULL, withFilter = TRUE,
  keepNA = FALSE, sep = ", ", stack = FALSE, firstColumn = FALSE,
  lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE
)


## Agences
sheet_name <- "Agences"
addWorksheet(wb, sheet_name)
df <- agences

writeDataTable(
  wb,sheet=sheet_name, x=df, 
  startCol = 1, startRow = 1, xy = NULL,
  colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleLight9",
  tableName = NULL, headerStyle = NULL, withFilter = TRUE,
  keepNA = FALSE, sep = ", ", stack = FALSE, firstColumn = FALSE,
  lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE
)


## PDVs
sheet_name <- "PDVs"
addWorksheet(wb, sheet_name)
df <- pdvs

writeDataTable(
  wb,sheet=sheet_name, x=df, 
  startCol = 1, startRow = 1, xy = NULL,
  colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleLight9",
  tableName = NULL, headerStyle = NULL, withFilter = TRUE,
  keepNA = FALSE, sep = ", ", stack = FALSE, firstColumn = FALSE,
  lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE
)


#----------------


## Coordonateurs
sheet_name <- "Coordonateurs"
addWorksheet(wb, sheet_name)
df <- coordonateurs

writeDataTable(
  wb,sheet=sheet_name, x=df, 
  startCol = 1, startRow = 1, xy = NULL,
  colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleLight9",
  tableName = NULL, headerStyle = NULL, withFilter = TRUE,
  keepNA = FALSE, sep = ", ", stack = FALSE, firstColumn = FALSE,
  lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE
)


## Coordonateurs Adjoints
sheet_name <- "Coordonateurs Adjoints"
addWorksheet(wb, sheet_name)
df <- coordonateurs_adjoints

writeDataTable(
  wb,sheet=sheet_name, x=df, 
  startCol = 1, startRow = 1, xy = NULL,
  colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleLight9",
  tableName = NULL, headerStyle = NULL, withFilter = TRUE,
  keepNA = FALSE, sep = ", ", stack = FALSE, firstColumn = FALSE,
  lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE
)


## Chefs d'agence
sheet_name <- "Chefs d'agence"
addWorksheet(wb, sheet_name)
df <- chefs_agences

writeDataTable(
  wb,sheet=sheet_name, x=df, 
  startCol = 1, startRow = 1, xy = NULL,
  colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleLight9",
  tableName = NULL, headerStyle = NULL, withFilter = TRUE,
  keepNA = FALSE, sep = ", ", stack = FALSE, firstColumn = FALSE,
  lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE
)



## Agents
sheet_name <- "Agents"
addWorksheet(wb, sheet_name)
df <- agents

writeDataTable(
  wb,sheet=sheet_name, x=df, 
  startCol = 1, startRow = 1, xy = NULL,
  colNames = TRUE, rowNames = FALSE, tableStyle = "TableStyleLight9",
  tableName = NULL, headerStyle = NULL, withFilter = TRUE,
  keepNA = FALSE, sep = ", ", stack = FALSE, firstColumn = FALSE,
  lastColumn = FALSE, bandedRows = TRUE, bandedCols = FALSE
)


saveWorkbook(wb, fake_db_path,overwrite = TRUE)





#
#---------------------------------------------
#


#
#---------------------------------------------
#

#
# END ##############################################################################################################
#
EndTime <- Sys.time() 
Duree <- difftime(EndTime,StartTime,unit="mins")
Duree
#
#rm(list=ls())
gc()
#
#











