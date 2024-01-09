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

cf <- fs |>
  filter(demonstrativo == 'Demonstrações dos Fluxos de Caixa')

# Cascata Histórica ----

dates <- seq.Date(from = as.Date('2019-12-31', '%Y-%m-%d'),
                  to   = as.Date('2022-12-31', '%Y-%m-%d'),
                  by = 'year')

for (date in seq_along(dates)) {
  tcf <- cf %>%
    filter(data == as.Date(dates[date], '%Y-%m-%d')) %>%
    mutate(categoria = as.character(categoria),
           item      = as.character(item)) |>
    mutate(categoria = case_when(categoria == 'Lucro Líquido do Exercício' |
                                   categoria == 'Ajuste de Itens Não-Caixa' ~
                                   'Fluxo de Caixa das Atividades Operacionais',
                                 TRUE ~ categoria)) %>%
    group_by(categoria) %>%
    summarise(valor = sum(valor)) %>%
    arrange(factor(categoria, levels = c(
      'Fluxo de Caixa das Atividades Operacionais',
      'Fluxo de Caixa das Atividades de Investimentos',
      'Fluxo de Caixa das Atividades de Financiamento'
    ))) %>%
    adorn_totals(name = 'Variação do Caixa', where = 'row', fill = 'Variação do Caixa') %>%
    select(categoria, valor) %>%
    mutate(valor = case_when(valor == 0 ~ 1e-9,
                             TRUE ~ valor)) %>%
    mutate(categoria = as.factor(categoria),
           id   = seq_along(categoria),
           type = case_when(valor >= 0 ~ 'in',
                            valor <  0 ~ 'out'),
           type = case_when(categoria == 'Variação do Caixa' ~
                              'net',
                            TRUE ~ type)) %>%
    mutate(end = cumsum(valor),
           end = case_when(categoria == 'Variação do Caixa' ~ 0,
                           TRUE ~ end)) %>%
    mutate(start = lag(end, 1),
           start = case_when(is.na(start) == TRUE ~ 0,
                             TRUE ~ start)) %>%
    mutate(categoria = fct_reorder(categoria, id)) %>%
    arrange(-row_number()) %>%
    mutate(id = seq_along(id)) %>%
    mutate(categoria = fct_reorder(categoria, id))
  
  ttcf <- tcf %>%
    mutate(valor = format(round(valor * 1e-3, 1)), scientific = FALSE, big.mark = ',') %>%
    select(categoria, valor)
  print(dates[date])
  print(ttcf)
  
  gcf <- ggplot(tcf) +
    aes(categoria, fill = type) +
    geom_rect(aes(x = categoria, xmin = id - .45, xmax = id + 0.45, ymin = end, ymax = start)) +
    geom_hline(yintercept = 0, color = 'red', linetype = 'dashed', size = .6) +
    coord_flip() +
    scale_fill_manual(values = c('in'  = r2c_pistache_1,
                                 'out' = r2c_vermelho_1,
                                 'net' = r2c_blue_2)) +
    theme_minimal() +
    theme(panel.background = element_rect(fill     =  'transparent',
                                          colour   =  'transparent'),
          axis.line.y          =   element_blank(),
          axis.line.x          =   element_blank(),
          panel.grid.major     =   element_blank(),
          panel.grid.minor     =   element_blank(),
          legend.title         =   element_blank(),
          legend.text          =   element_text(size = 10),
          legend.position      =   c("none"),
          legend.key.size      =   unit(1,"cm"),
          axis.text.x          =   element_blank(),
          axis.text.y          =   element_blank(),
          axis.title.x         =   element_blank(),
          axis.title.y         =   element_blank(),
          rect                 =   element_rect(fill = "transparent"),
          plot.margin          =   unit(c(.1,.1,.1,.1),"cm"))
  
  
  ### Armazenar o Gráfico ----
  
  ggsave(filename =   here(output,
                           paste0(name,
                                  ' - Variacao de Caixa - ',
                                  year(as.Date(dates[date])),
                                  '.png')),
         plot     =  gcf,
         device   =  'png',
         scale    =  1,
         width    =  3,
         height   =  3.5,
         units    =  c('cm'),
         dpi      =  500,
         bg       = 'transparent')
}

# Histórico do Nível de Caixa----

hcf <- cf %>%
  filter(data >= as.Date('2019-01-01', '%Y-%m-%d')) %>%
  mutate(data = year(data)) %>%
  group_by(data) %>%
  summarise(valor = sum(valor))

hcf[1,2] <- 1245867

hcf <- hcf %>%
  mutate(caixa = cumsum(valor))

gcf <- hcf %>%
  ggplot() +
  aes(x = data, y = caixa) +
  geom_bar(position = 'stack', stat = 'identity', fill = r2c_azul_1) +
  geom_hline(yintercept = 0) +
  scale_x_continuous(breaks = pretty_breaks(n = 4)) +
  theme_minimal() +
  theme(panel.background = element_rect(fill     =  'transparent',
                                        colour   =  'transparent'),
        axis.line.y          =   element_blank(),
        axis.line.x          =   element_blank(),
        panel.grid.major     =   element_blank(),
        panel.grid.minor     =   element_blank(),
        legend.title         =   element_blank(),
        legend.text          =   element_text(size = 10),
        legend.position      =   c("none"),
        legend.key.size      =   unit(1,"cm"),
        axis.text.x          =   element_text(size = 8, color = 'black',
                                              face = 'italic'),
        axis.text.y          =   element_blank(),
        axis.title.x         =   element_blank(),
        axis.title.y         =   element_blank(),
        rect                 =   element_rect(fill = "transparent"),
        plot.margin          =   unit(c(.1,.1,.1,.1),"cm"))

## Armazenar o Gráfico ----

ggsave(filename =   here(output,
                         paste0(name,
                                ' - Historico nivel de caixa ',
                                '.png')),
       plot     =  gcf,
       device   =  'png',
       scale    =  1,
       width    =  13.1,
       height   =  3.8,
       units    =  c('cm'),
       dpi      =  500,
       bg       = 'transparent')

# Variação Acumulada do Caixa ----

tcf <- cf %>%
  mutate(categoria = as.character(categoria),
         item      = as.character(item)) |>
  filter(data >= as.Date('2019-01-01', '%Y-%m-%d')) %>%
  mutate(categoria = case_when(categoria == 'Lucro Líquido do Exercício' |
                                 categoria == 'Ajuste de Itens Não-Caixa' ~
                                 'Fluxo de Caixa das Atividades Operacionais',
                               TRUE ~ categoria)) %>%
  group_by(categoria) %>%
  summarise(valor = sum(valor)) %>%
  arrange(factor(categoria, levels = c(
    'Fluxo de Caixa das Atividades Operacionais',
    'Fluxo de Caixa das Atividades de Investimentos',
    'Fluxo de Caixa das Atividades de Financiamento'
  ))) %>%
  adorn_totals(name = 'Variação do Caixa', where = 'row', fill = 'Variação do Caixa') %>%
  select(categoria, valor) %>%
  mutate(valor = case_when(valor == 0 ~ 1e-9,
                           TRUE ~ valor)) %>%
  mutate(categoria = as.factor(categoria),
         id   = seq_along(categoria),
         type = case_when(valor >= 0 ~ 'in',
                          valor <  0 ~ 'out'),
         type = case_when(categoria == 'Variação do Caixa' ~
                            'net',
                          TRUE ~ type)) %>%
  mutate(end = cumsum(valor),
         end = case_when(categoria == 'Variação do Caixa' ~ 0,
                         TRUE ~ end)) %>%
  mutate(start = lag(end, 1),
         start = case_when(is.na(start) == TRUE ~ 0,
                           TRUE ~ start)) %>%
  mutate(categoria = fct_reorder(categoria, id)) %>%
  mutate(id = seq_along(id)) %>%
  mutate(categoria = fct_reorder(categoria, id))

ttcf <- tcf %>%
  mutate(valor = format(round(valor * 1e-3, 1)), scientific = FALSE, big.mark = ',') %>%
  select(categoria, valor)
print(ttcf)

gcf <- ggplot(tcf) +
  aes(categoria, fill = type) +
  geom_rect(aes(x = categoria, xmin = id - .45, xmax = id + 0.45, ymin = end, ymax = start)) +
  geom_hline(yintercept = 0, color = 'red', linetype = 'dashed', size = .6) +
  scale_fill_manual(values = c('in'  = r2c_pistache_1,
                               'out' = r2c_vermelho_1,
                               'net' = r2c_blue_2)) +
  theme_minimal() +
  theme(panel.background = element_rect(fill     =  'transparent',
                                        colour   =  'transparent'),
        axis.line.y          =   element_blank(),
        axis.line.x          =   element_blank(),
        panel.grid.major     =   element_blank(),
        panel.grid.minor     =   element_blank(),
        legend.title         =   element_blank(),
        legend.text          =   element_text(size = 10),
        legend.position      =   c("none"),
        legend.key.size      =   unit(1,"cm"),
        axis.text.x          =   element_blank(),
        axis.text.y          =  element_blank(),
        axis.title.x         =   element_blank(),
        axis.title.y         =   element_blank(),
        rect                 =   element_rect(fill = "transparent"),
        plot.margin          =   unit(c(.1,.1,.1,.1),"cm"))


### Armazenar o Gráfico ----

ggsave(filename =   here(output,
                         paste0(name,
                                ' - FC Acumulado',
                                '.png')),
       plot     =  gcf,
       device   =  'png',
       scale    =  1,
       width    =  12.5,
       height   =  4.5,
       units    =  c('cm'),
       dpi      =  500,
       bg       = 'transparent')

## Mini FC Acumulado ----

gcf <- ggplot(tcf) +
  aes(categoria, fill = type) +
  geom_rect(aes(x = categoria, xmin = id - .45, xmax = id + 0.45, ymin = end, ymax = start)) +
  geom_hline(yintercept = 0, color = 'red', linetype = 'dashed', size = .6) +
  scale_fill_manual(values = c('in'  = r2c_pistache_1,
                               'out' = r2c_vermelho_1,
                               'net' = r2c_blue_2)) +
  coord_flip() +
  theme_minimal() +
  theme(panel.background = element_rect(fill     =  'transparent',
                                        colour   =  'transparent'),
        axis.line.y          =   element_blank(),
        axis.line.x          =   element_blank(),
        panel.grid.major     =   element_blank(),
        panel.grid.minor     =   element_blank(),
        legend.title         =   element_blank(),
        legend.text          =   element_text(size = 10),
        legend.position      =   c("none"),
        legend.key.size      =   unit(1,"cm"),
        axis.text.x          =   element_blank(),
        axis.text.y          =  element_blank(),
        axis.title.x         =   element_blank(),
        axis.title.y         =   element_blank(),
        rect                 =   element_rect(fill = "transparent"),
        plot.margin          =   unit(c(.1,.1,.1,.1),"cm"))


ggsave(filename =   here(output,
                         paste0(name,
                                ' - mini FC Acumulado',
                                '.png')),
       plot     =  gcf,
       device   =  'png',
       scale    =  1,
       width    =  3,
       height   =  3.5,
       units    =  c('cm'),
       dpi      =  500,
       bg       = 'transparent')

# Fluxo de Caixa Operacional ----

tcf <- cf %>%
  filter(data >= as.Date('2017-12-31', '%Y-%m-%d'),
         valor != 0) %>%
  mutate(categoria = as.character(categoria),
         item = as.character(item),
         descricao = as.character(descricao),
         categoria = case_when(categoria == 'Lucro Líquido do Exercício' |
                                 categoria == 'Ajuste de Itens Não-Caixa' ~
                                 'Fluxo de Caixa das Atividades Operacionais',
                               TRUE ~ categoria)) %>%
  filter(categoria == 'Fluxo de Caixa das Atividades Operacionais') %>%
  mutate(descricao = case_when(descricao == 'Lucro Líquido do Exercício'  |
                               descricao == 'Equivalência Patrimonial'    |
                               descricao == 'Lucros não Realizados nos Estoques' |
                               descricao == 'Variação Cambial' |
                               descricao == 'Conversão de Moeda' |
                               descricao == 'Apropriação de despesas de juros sobre capital próprio' |
                               descricao == 'Ganho no ativo imobilizado, intangível e investimentos' |
                               descricao == 'Outros itens não monetários incluídos no resultado' |
                               descricao == 'Ganhos na compra vantajosa' |
                               descricao == 'Receita de Juros sobre Capital Próprio' |
                               descricao == 'Reversões passivo tributário, cíveis e trabalhista' |
                               descricao == 'Receita de Incentivos Fiscais' |
                               descricao == 'IR e CSLL' |
                               descricao == 'Provisões de clientes inadimplentes' |
                               descricao == 'Ganhos com Derivativos' |
                               descricao == 'Provisão de Juros Ativos' |
                               descricao == 'Provisão de Juros Passivos' |
                               descricao == 'Provisões para Redução de Valor Recuperável' |
                               descricao == 'Ressarcimento de Despesas Comerciais' |
                               descricao == 'Recueração crédito extemporâneo' |
                               descricao == 'Passivo Tributário' |
                               descricao == 'Saldo de de empresas encerradas em 2020' ~
                                 'Ajuste de Itens não caixa',
                               TRUE ~ descricao)) %>%
  group_by(descricao) %>%
  summarise(valor = sum(valor)) %>%
  mutate(descricao = case_when(abs(valor) < 200000 ~ 'Outros',
                               TRUE ~ descricao)) %>%
  group_by(descricao) %>%
  summarise(valor = sum(valor)) %>%
  arrange(desc(valor)) %>%
  arrange(factor(descricao, levels = c(
    'Lucro Líquido',
    'Ajuste de Itens Não-Caixa'
  ))) %>%
  adorn_totals(name = 'Fluxo de Caixa Operacional', where = 'row', fill = 'Variação do Caixa') %>%
  select(descricao, valor) %>%
  mutate(valor = case_when(valor == 0 ~ 1e-9,
                           TRUE ~ valor)) %>%
  mutate(descricao = as.factor(descricao),
         id   = seq_along(descricao),
         type = case_when(valor >= 0 ~ 'in',
                          valor <  0 ~ 'out'),
         type = case_when(descricao == 'Fluxo de Caixa Operacional' ~
                            'net',
                          TRUE ~ type)) %>%
  mutate(end = cumsum(valor),
         end = case_when(descricao == 'Fluxo de Caixa Operacional' ~ 0,
                         TRUE ~ end)) %>%
  mutate(start = lag(end, 1),
         start = case_when(is.na(start) == TRUE ~ 0,
                           TRUE ~ start)) %>%
  mutate(descricao = fct_reorder(descricao, id)) %>%
  mutate(id = seq_along(id)) %>%
  mutate(descricao = fct_reorder(descricao, id))

ttcf <- tcf %>%
  mutate(valor = format(round(valor * 1e-3, 1)), scientific = FALSE, big.mark = ',') %>%
  select(descricao, valor)
print(ttcf)

gcf <- ggplot(tcf) +
  aes(descricao, fill = type) +
  geom_rect(aes(x = descricao, xmin = id - .45, xmax = id + 0.45, ymin = end, ymax = start)) +
  geom_hline(yintercept = 0, color = 'red', linetype = 'dashed', size = .6) +
  scale_fill_manual(values = c('in'  = r2c_pistache_1,
                               'out' = r2c_vermelho_1,
                               'net' = r2c_blue_2)) +
  theme_minimal() +
  theme(panel.background = element_rect(fill     =  'transparent',
                                        colour   =  'transparent'),
        axis.line.y          =   element_blank(),
        axis.line.x          =   element_blank(),
        panel.grid.major     =   element_blank(),
        panel.grid.minor     =   element_blank(),
        legend.title         =   element_blank(),
        legend.text          =   element_text(size = 10),
        legend.position      =   c("none"),
        legend.key.size      =   unit(1,"cm"),
        axis.text.x          =   element_blank(),
        axis.text.y          =  element_blank(),
        axis.title.x         =   element_blank(),
        axis.title.y         =   element_blank(),
        rect                 =   element_rect(fill = "transparent"),
        plot.margin          =   unit(c(.1,.1,.1,.1),"cm"))

## Armazenar o Gráfico ----

ggsave(filename =   here(output,
                         paste0(name,
                                ' - FCO Acumulado',
                                '.png')),
       plot     =  gcf,
       device   =  'png',
       scale    =  1,
       width    =  20.0,
       height   =  3.4,
       units    =  c('cm'),
       dpi      =  500,
       bg       = 'transparent')

# Fluxo de Caixa de Investimentos ----

depr <- fs %>%
  filter(demonstrativo == 'Demonstrações dos Fluxos de Caixa',
         categoria     == 'Ajuste de Itens Não-Caixa',
         item          == 'Depreciação e Amortização') %>%
  mutate(data = year(data)) %>%
  group_by(data) %>%
  summarise(valor = sum(valor)) %>%
  mutate(classe = 'Depreciação e Amortização')

deprec <- sum(depr$valor)

tcf <- cf %>%
  mutate(descricao = as.character(descricao)) |>
  filter(data >= as.Date('2017-12-31', '%Y-%m-%d'),
         valor != 0) %>%
  filter(classe == 'Fluxo de Caixa das Atividades de Investimentos') %>%
  group_by(descricao) %>%
  summarise(valor = sum(valor)) %>%
  mutate(descricao = case_when(abs(valor) < 50000 ~ 'Outros',
                               TRUE ~ descricao)) %>%
  group_by(descricao) %>%
  summarise(valor = sum(valor)) %>%
  arrange(desc(valor)) %>%
  adorn_totals(name = 'Fluxo de Caixa de Investimentos',
               where = 'row',
               fill = 'Variação do Caixa') %>%
  select(descricao, valor) %>%
  mutate(valor = case_when(valor == 0 ~ 1e-9,
                           TRUE ~ valor)) %>%
  mutate(descricao = as.factor(descricao),
         id   = seq_along(descricao),
         type = case_when(valor >= 0 ~ 'in',
                          valor <  0 ~ 'out'),
         type = case_when(descricao == 'Fluxo de Caixa de Investimentos' ~
                            'net',
                          TRUE ~ type)) %>%
  mutate(end = cumsum(valor),
         end = case_when(descricao == 'Fluxo de Caixa de Investimentos' ~ 0,
                         TRUE ~ end)) %>%
  mutate(start = lag(end, 1),
         start = case_when(is.na(start) == TRUE ~ 0,
                           TRUE ~ start)) %>%
  mutate(descricao = fct_reorder(descricao, id)) %>%
  mutate(id = seq_along(id)) %>%
  mutate(descricao = fct_reorder(descricao, id))

ttcf <- tcf %>%
  mutate(valor = format(round(valor * 1e-3, 1)), scientific = FALSE, big.mark = ',') %>%
  select(descricao, valor)
print(ttcf)

gcf <- ggplot(tcf) +
  aes(descricao, fill = type) +
  geom_rect(aes(x = descricao, xmin = id - .45, xmax = id + 0.45, ymin = end, ymax = start)) +
  geom_hline(yintercept = 0, color = 'red', linetype = 'dashed', size = .6) +
  scale_fill_manual(values = c('in'  = r2c_pistache_1,
                               'out' = r2c_vermelho_1,
                               'net' = r2c_blue_2)) +
  theme_minimal() +
  theme(panel.background = element_rect(fill     =  'transparent',
                                        colour   =  'transparent'),
        axis.line.y          =   element_blank(),
        axis.line.x          =   element_blank(),
        panel.grid.major     =   element_blank(),
        panel.grid.minor     =   element_blank(),
        legend.title         =   element_blank(),
        legend.text          =   element_text(size = 10),
        legend.position      =   c("none"),
        legend.key.size      =   unit(1,"cm"),
        axis.text.x          =   element_blank(),
        axis.text.y          =  element_blank(),
        axis.title.x         =   element_blank(),
        axis.title.y         =   element_blank(),
        rect                 =   element_rect(fill = "transparent"),
        plot.margin          =   unit(c(.1,.1,.1,.1),"cm"))

## Armazenar o Gráfico ----

ggsave(filename =   here(output,
                         paste0(name,
                                ' - FCI Acumulado',
                                '.png')),
       plot     =  gcf,
       device   =  'png',
       scale    =  1,
       width    =  10.0,
       height   =  3.0,
       units    =  c('cm'),
       dpi      =  500,
       bg       = 'transparent')

# Fluxo de Caixa de Financiamentos ----

tcf <- cf %>%
  mutate(classe = as.character(classe),
         descricao = as.character(descricao)) |>
  filter(data >= as.Date('2017-12-31', '%Y-%m-%d'),
         valor != 0) %>%
  filter(classe == 'Fluxo de Caixa das Atividades de Financiamento') %>%
  mutate(descricao = case_when(abs(valor) < 10000 ~ 'Outros',
                               TRUE ~ descricao)) %>%
  group_by(descricao) %>%
  summarise(valor = sum(valor)) %>%
  group_by(descricao) %>%
  summarise(valor = sum(valor)) %>%
  arrange(desc(valor)) %>%
  adorn_totals(name = 'Fluxo de Caixa das Atividades de Financiamentos', where = 'row', fill = 'Variação do Caixa') %>%
  select(descricao, valor) %>%
  mutate(valor = case_when(valor == 0 ~ 1e-9,
                           TRUE ~ valor)) %>%
  mutate(descricao = as.factor(descricao),
         id   = seq_along(descricao),
         type = case_when(valor >= 0 ~ 'in',
                          valor <  0 ~ 'out'),
         type = case_when(descricao == 'Fluxo de Caixa das Atividades de Financiamentos' ~
                            'net',
                          TRUE ~ type)) %>%
  mutate(end = cumsum(valor),
         end = case_when(descricao == 'Fluxo de Caixa das Atividades de Financiamentos' ~ 0,
                         TRUE ~ end)) %>%
  mutate(start = lag(end, 1),
         start = case_when(is.na(start) == TRUE ~ 0,
                           TRUE ~ start)) %>%
  mutate(descricao = fct_reorder(descricao, id)) %>%
  mutate(id = seq_along(id)) %>%
  mutate(descricao = fct_reorder(descricao, id))

ttcf <- tcf %>%
  mutate(valor = format(round(valor * 1e-3, 1)), scientific = FALSE, big.mark = ',') %>%
  select(descricao, valor)
print(ttcf)

gcf <- ggplot(tcf) +
  aes(descricao, fill = type) +
  geom_rect(aes(x = descricao, xmin = id - .45, xmax = id + 0.45, ymin = end, ymax = start)) +
  geom_hline(yintercept = 0, color = 'red', linetype = 'dashed', size = .6) +
  scale_fill_manual(values = c('in'  = r2c_pistache_1,
                               'out' = r2c_vermelho_1,
                               'net' = r2c_blue_2)) +
  theme_minimal() +
  theme(panel.background = element_rect(fill     =  'transparent',
                                        colour   =  'transparent'),
        axis.line.y          =   element_blank(),
        axis.line.x          =   element_blank(),
        panel.grid.major     =   element_blank(),
        panel.grid.minor     =   element_blank(),
        legend.title         =   element_blank(),
        legend.text          =   element_text(size = 10),
        legend.position      =   c("none"),
        legend.key.size      =   unit(1,"cm"),
        axis.text.x          =   element_blank(),
        axis.text.y          =  element_blank(),
        axis.title.x         =   element_blank(),
        axis.title.y         =   element_blank(),
        rect                 =   element_rect(fill = "transparent"),
        plot.margin          =   unit(c(.1,.1,.1,.1),"cm"))

## Armazenar o Gráfico ----

ggsave(filename =   here(output,
                         paste0(name,
                                ' - FCF Acumulado ',
                                '.png')),
       plot     =  gcf,
       device   =  'png',
       scale    =  1,
       width    =  10.0,
       height   =  3.0,
       units    =  c('cm'),
       dpi      =  500,
       bg       = 'transparent')

