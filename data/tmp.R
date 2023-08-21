



file_path <- 'C:/TravauxR/pyCharm/local/rh/data/pv_ref_202307.xlsx'
df <- read_excel(file_path, sheet = "Archi")


pays <- df %>% 
  select(PAYS = PAYS_COORDONATEUR) %>%
  distinct() %>%
  arrange(PAYS)

coords <- df %>% 
  select(
    COORD = COORDONATEUR,
    PAYS = PAYS_COORDONATEUR
  ) %>%
  distinct() %>%
  arrange(PAYS, COORD)


agences <- df %>% 
  select(
    AGENCE,
    COORD = COORDONATEUR,
    PAYS = PAYS_COORDONATEUR
  ) %>%
  distinct() %>%
  arrange(PAYS, COORD, AGENCE)


pdvs <- df %>% 
  mutate(PDV = str_squish(string = PDV)) %>% 
  select(
    PDV,
    AGENCE,
    COORD = COORDONATEUR,
    PAYS = PAYS_COORDONATEUR
  ) %>%
  distinct() %>%
  arrange(PAYS, COORD, AGENCE, PDV)


#----------------

employees <- df %>% 
  select(
    AGENT,
    PDV,
    AGENCE,
    COORD = COORDONATEUR,
    PAYS = PAYS_COORDONATEUR
  ) %>%
  distinct() %>%
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
  )) %>%
  rename(EMPLOYEE = AGENT) %>%
  rowid_to_column(var = 'tmp') %>% 
  mutate(MAT = paste0('M2023', str_pad(tmp, 3, pad = "0"))) %>%
  select(-c(tmp))



coordonateurs <- employees %>%
  filter(CATEGORIE == "Coordonateur") %>%
  select(MAT, EMPLOYEE, CATEGORIE, SALAIRE, COORD)


coordonateurs_adjoints <- employees %>%
  filter(CATEGORIE == "Coordonateur adjoint") %>%
  select(MAT, EMPLOYEE, CATEGORIE, SALAIRE, COORD)

chefs_agences <- employees %>%
  filter(CATEGORIE == "Chef d'agence") %>%
  select(MAT, EMPLOYEE, CATEGORIE, SALAIRE, AGENCE)


agents <- employees %>% 
  filter(CATEGORIE == "Agent") %>%
  select(MAT, EMPLOYEE, CATEGORIE, SALAIRE, PDV)


#----------------

fake_db_path <- 'C:/TravauxR/pyCharm/local/rh/data/fake_db.xlsx'
wb <- createWorkbook(fake_db_path)



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





#----------------

