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
library(readxl)
library(scales)
library(purrr)
library(stringr)
library(tidyverse)

## Projeto ----

name <- 'Zuquetti e Marzola'

## Diretórios ----

raw     <- here('data', 'raw')
treated <- here('data', 'treated')
output  <- here('output')

## Sources ----

map(here('R',
         c('r2c colors.R',
           'dados.R')),
    source)

# Dados ----

## Carregar Dados ----

is <- fs |>
  filter(demonstrativo == 'Demonstração dos Resultados do Exercício')

bs <- fs |>
  filter(demonstrativo == 'Balanço Patrimonial')

is_real <-  fs %>%
  filter(demonstrativo == 'Demonstração dos Resultados do Exercício') %>%
  mutate(item = as.character(item),
         item = case_when(item == 'Despesas Financeirasa' ~
                            'Despesas Financeiras',
                          TRUE ~ item)) |>
  filter(!item %in% c('Despesas Financeiras',
                      'Receitas Financeiras',
                      'IR e CSLL',
                      'IR e CSLL Diferidos',
                      'Resultado dos Não Controladores',
                      'Resultado Atribuído a SCPs'),
         valor != 0,
         data >= as.Date('2017-01-01', '%Y-%m-%d')) %>%
  mutate(data = year(data)) %>%
  group_by(data, classe) %>%
  summarise(valor = sum(valor))

is_full <- is_real

### Fluxos de Caixa ----

fc_real <-  fs %>%
  filter(demonstrativo == 'Demonstrações dos Fluxos de Caixa',
         valor         != 0,
         data >= as.Date('2019-01-01', '%Y-%m-%d')) %>%
  mutate(data = year(data)) %>%
  mutate(categoria = as.character(categoria)) |>
  mutate(categoria = case_when(categoria == 'Ajuste de Itens Não-Caixa' |
                                 categoria == 'Lucr Líquido do Exercício' ~
                                 'Fluxo de Caixa das Atividades Operacionais',
                               categoria == 'Fluxo de Caixa das Atividades de Investimentos' ~
                                 'Fluxo de Caixa de Investimentos',
                               categoria == 'Fluxo de Caixa das Atividades de Financiamento' ~
                                 'Fluxo de Caixa de Financiamento',
                               TRUE ~ categoria)) %>%
  mutate(categoria = case_when(categoria == 'Fluxo de Caixa das Atividades Operacionais' ~
                                 'Fluxo de Caixa Operacional',
                               TRUE ~ categoria)) %>%
  group_by(data, categoria) %>%
  summarise(valor = sum(valor)) %>%
  rename('classe' = categoria)

### Depreciação e Amortização ----


#### Realizada CF ----

dep_real <- fs %>%
  filter(demonstrativo == 'Demonstrações dos Fluxos de Caixa',
         categoria     == 'Ajuste de Itens Não-Caixa',
         item          == 'Depreciação e Amortização') %>%
  mutate(data = year(data)) %>%
  group_by(data) %>%
  summarise(valor = sum(valor)) %>%
  mutate(classe = 'Depreciação e Amortização')


dep_full <- dep_real

# Tabela EBITDA ----

res_op <- is_full %>%
  group_by(data) %>%
  summarise(valor = sum(valor)) %>%
  mutate(classe = 'Resultado Operacional')

ebitda <- dep_full %>%
  bind_rows(res_op) %>%
  filter(data >= 2019,
         data <= 2022) %>%
  mutate(data = paste0(data, 'Y')) %>%
  pivot_wider(id_cols = classe,
              names_from = 'data',
              values_from = 'valor') %>%
  arrange(factor(classe,
                 levels = c('Resultado Operacional',
                            'Depreciação e Amortização'))) %>%
  adorn_totals(where = 'col') %>%
  adorn_totals(where = 'row', name = 'EBITDA') %>%
  select(classe, Total, `2019Y`, `2020Y`, `2021Y`) %>%
  mutate(across(-classe, ~ round(.x * 1e-0,0))) %>%
  mutate(across(-classe, ~ as.character(accounting(.x,
                                                   digits = 0L)))) %>%
  mutate(across(-classe, ~ case_when(.x == '0.0' ~ '-',
                                     TRUE ~ .x)))

## Flextable ----

flex_ebitda <- ebitda %>%
  flextable() %>%
  
  border_remove() %>%
  
  # CABEÇALHO
  set_header_labels(classe    = 'Item') %>%
  hline(i = 1,  j = 1:ncol(ebitda),
        part = 'header', fp_border(width = 2, color = r2c_amarelo_1)) %>%
  bold(i = 1,  j = 1, part = 'header', bold = TRUE) %>%
  fontsize(i = 1, j = 1, part = 'header', size = 12) %>%
  italic(i = 1, j = 1, part = 'header', italic = TRUE) %>%
  font(part = 'header', fontname = 'Cambria') %>%
  color(part = 'header', color = r2c_azul_1) %>%
  align(j = 1,  part = 'header',   align = 'left') %>%
  align(j = 2:ncol(ebitda),  part = 'header',   align = 'center') %>%
  valign(j = 2, part = 'header',   valign = 'bottom') %>%
  padding(part = 'header', padding.bottom = .5) %>%
  
  # LINHAS
  hline(i = nrow(ebitda) - 1 ,  j = 1:ncol(ebitda), part = 'body', fp_border(width = 1,
                                                                             color = r2c_azul_1)) %>%
  hline(i = nrow(ebitda), j = 1:ncol(ebitda), part = 'body', fp_border(width = 1,
                                                                       color = r2c_azul_1)) %>%
  vline(i = 1:nrow(ebitda), j = 1:2, part = 'body',
        fp_border(width = 1.2, color = r2c_azul_1)) %>%
  vline(i = 1, j = 1:2, part = 'header',
        fp_border(width = 1.2, color = r2c_azul_1)) %>%
  
  # NEGRITOS
  bold(i = nrow(ebitda) ,  j = 1:ncol(ebitda), bold = TRUE) %>%
  
  # MARGENS
  padding(i = 1:(nrow(ebitda) - 1) , j = 1, padding.left = 15) %>%
  
  # ALINHAMENTOS
  align(j = 1,  part = 'body',   align = 'left') %>%
  align(j = 2:ncol(ebitda),  part = 'body',   align = 'right') %>%
  width(j = 1,  width = 8.0) %>%
  width(j = 2:ncol(ebitda),  width = 1.5) %>%
  
  # FONTE
  font(part = 'body', fontname = 'Cambria') %>%
  fontsize(part = 'body', size = 11)

#### Armazenar Tabela ----

save_as_image(flex_ebitda,
              here(output,
                   paste0(name,
                          ' - Tabela EBITDA.png')),
              zoom = 1)
