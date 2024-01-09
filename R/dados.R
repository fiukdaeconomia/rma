# Configurações Iniciais ----

## Pacotes ----
library(dplyr)
library(flextable)
library(formattable)
library(glue)
library(here)
library(janitor)
library(lubridate)
library(officer)
library(readr)
library(readxl)
library(scales)
library(purrr)
library(stringr)
library(tidyverse)
library(writexl)


## Projeto ----

name <- 'RMA CP'
vers <- '0.0.1'

## Diretórios ----

raw     <- here('data', 'raw')
treated <- here('data', 'treated')
xlsf    <- here('planilhas',
                'own')

# Dados ----

## Carregar Dados ----

fs <- read_excel(here(xlsf,
                      glue('{name} - {vers}.xlsx')),
                 sheet = 'Dados',
                 range = cell_cols('C:J')) |>
  na.omit() |>
  row_to_names(row_number = 1) |>
  clean_names() |>
  mutate(data          = as.Date(as.numeric(data),
                                 origin = "1899-12-30")) |> 
  mutate(across(c(demonstrativo,
                  empresa,
                  categoria,
                  classe,
                  item,
                  descricao),
                ~as.factor(.x))) |> 
  mutate(valor = as.numeric(valor)) |> 
  mutate(item = as.character(item),
         item = case_when(item == 'Caixa e Equivalentes de Caixa' ~
                            'Disponibilidades',
                            TRUE ~ item),
         item = case_when(item == 'Aplicações de Liquidez não Imediata' ~
                            'Aplicações Financeiras',
                          TRUE ~ item),
         item = case_when(item == 'Depósitos Restitutíveis e Valores Vinculados' ~
                            'Depósitos Restituíveis',
                          TRUE ~ item))

## Armazemar Dados ----

### csv ----

write.csv(fs,
          here(treated,
               glue('dados - {name}.csv')),
          row.names = FALSE)

### txt ----

write.table(fs,
            here(treated,
                 glue('dados - {name}.txt')),
            row.names = FALSE)

### xlsx ----

write_xlsx(fs,
           here(treated,
                glue('dados - {name}.xlsx')),
           col_names = TRUE)
