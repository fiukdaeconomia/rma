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

name <- 'Fluxo de Caixa CP'
nick <- 'cpt'
vers <- '2.0.0'

## Diretórios ----

raw     <- here('data', 'raw')
treated <- here('data', 'treated')
xlsf    <- here('planilhas',
                'own')

source(here('R',
            'r2c colors.R'))

# Dados ----

## Carregar Dados ----

fs <- read_excel(here(xlsf,
                      glue('{name} - {vers}.xlsx')),
                 sheet = 'Dados',
                 range = cell_cols('C:K')) |>
  na.omit() |>
  row_to_names(row_number = 1) |>
  clean_names() |>
  mutate(data          = as.Date(as.numeric(data),
                                 origin = "1899-12-30")) |> 
  mutate(valor = as.numeric(valor)) 
 


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

# Análise ----

cx0 <- 26800

dados <- fs |> 
          filter(frequencia == 'Diária') |> 
          select(data,
                 categoria,
                 item,
                 valor) |> 
          pivot_wider(names_from = 'item',
                      values_from = 'valor') |> 
          clean_names() |> 
          mutate(liberacao_vinculada = case_when(is.na(liberacao_vinculada) == TRUE ~ 0,
                                                 TRUE ~ liberacao_vinculada),
                 liberacao_recebiveis = case_when(is.na(liberacao_recebiveis) == TRUE ~ 0,
                                                 TRUE ~ liberacao_recebiveis),
                 variacao_livre = entradas                +
                                  fornecedores            +
                                  pessoal                 +
                                  impostos                +
                                  movimentacao_financeira +
                                  resultado_financeiro    +
                                  outros                  +
                                  liberacao_vinculada     +
                                  liberacao_recebiveis,
                 variacao_bloq  = variacao_livre          -
                                  liberacao_vinculada     -
                                  liberacao_recebiveis,
                 caixa_livre = cx0 + cumsum(variacao_livre),
                 caixa_bloq  = cx0 + cumsum(variacao_bloq)) |> 
          pivot_longer(cols = c(- data,
                                - categoria),
                       names_to = 'variavel',
                       values_to = 'valor') |>
          mutate(categoria = case_when(data >= as.Date('2023-04-15') ~
                                         'Projetado',
                                       TRUE ~ categoria)) |> 
          mutate(variavel = case_when(variavel == 'caixa_bloq' ~
                                        'Caixa sem liberação',
                                      variavel == 'caixa_livre' ~
                                        'Caixa com liberação',
                                      variavel == 'liberacao_recebiveis' ~
                                        'Liberação de Recebíveis',
                                      variavel == 'liberacao_vinculada' ~
                                        'Liberação Vinculada',
                                      variavel == 'movimentacao financeira' ~
                                        'Movimentação Financeira',
                                      variavel == 'resultado_financeiro' ~
                                        'Resultado Financeiro',
                                      variavel == 'variacao_bloq' ~
                                        'Variação sem Liberação',
                                      variavel == 'variacao_livre' ~
                                        'Variação com Liberação', 
                                      TRUE ~ variavel)) |> 
            mutate(variavel = case_when(variavel == 'Caixa sem liberação' &
                                        categoria == 'Realizado' ~
                                          'Caixa realizado sem liberação',
                                        variavel == 'Caixa com liberação' &
                                          categoria == 'Realizado' ~
                                          'Caixa realizado com liberação',
                                        TRUE ~ variavel))

# Gráfico no Tempo ----

x <- dados |> 
  filter(variavel == 'Caixa sem liberação' |
         variavel == 'Caixa com liberação' |
         variavel == 'Caixa realizado com liberação') |> 
  ggplot() +
  aes(x = data, 
      y = valor,
      color = variavel) +
  geom_line(aes(linetype = categoria),
            size = .8) +
  scale_color_manual(values = c('Caixa realizado com liberação' = r2c_amarelo_1,
                                'Caixa sem liberação'           = r2c_vermelho_1,
                                'Caixa com liberação'           = r2c_azul_1)) +
  scale_linetype_manual(values = c('Realizado'           = 'solid',
                                   'Projetado' = 'longdash')) +
  geom_hline(yintercept = 0,
             size = .5,
             linetype = 'solid',
             colour = 'black') +
  geom_vline(xintercept = as.Date('2023-03-28'),
             size = 0.3,
             linetype = 'solid',
             colour = r2c_roxo_1) +
  scale_y_continuous(breaks = pretty_breaks(n = 5),
                     labels = comma,
                     limits = c(-250000, 250000)) +
  scale_x_date(breaks = seq(as.Date('2023-02-01'),
                            as.Date('2023-07-03'),
                            by = '15 days'),
               date_labels = '%d/%m',
               limits = c(as.Date('2023-01-31'),
                          as.Date('2023-07-04'))) +
  theme_minimal() +
  theme(panel.background = element_rect(fill     =  'transparent',
                                        colour   =  'transparent'),
        axis.line.y          =   element_blank(),
        axis.line.x          =   element_blank(),
        axis.ticks           =   element_blank(),
        panel.grid.major.x   = element_line(color = 'gray',
                                            size = 0.1),
        panel.grid.major.y   = element_line(color = 'gray',
                                            size = 0.1),
        legend.position      =   'none',
        axis.text.x          =   element_text(size = 7,
                                              face = 'bold',
                                              family = 'sans',
                                              colour = '#27474E'),
        axis.text.y          =   element_text(size = 6,
                                              face = 'bold',
                                              family = 'sans',
                                              colour = '#27474E'),
        axis.title.x         =   element_blank(),
        axis.title.y         =   element_blank(),
        rect                 =   element_rect(fill = "transparent"),
        plot.margin          =   unit(c(.1,.1,.1,.1),"cm"))

ggsave(x,
       filename = here('output',
                       glue('Posição de Caixa.png')),
       device   = 'png',
       scale    =  1,
       width    =  18,
       height   =  12,
       units    =  c('cm'),
       dpi      =  500,
       bg       = 'white')

# Decomposições ----

## Realizado ----

vars <- c('entradas',
          'fornecedores',
          'impostos',
          'movimentacao_financeira',
          'Resultado Financeiro',
          'outros',
          'pessoal',
          'Liberação de Recebíveis',
          'Liberação Vinculada')

tcf <- dados |> 
        filter(variavel %in% vars) |> 
        group_by(variavel) |> 
        summarise(valor = sum(valor)) |> 
        arrange(desc(valor)) |> 
        adorn_totals(where = 'row',
                     name = 'subtotal') |> 
        mutate(variavel = as.factor(variavel),
               id   = seq_along(variavel),
               type = case_when(valor >= 0 ~ 'in',
                                valor <  0 ~ 'out'),
               type = case_when(variavel == 'Total' ~
                                  'net',
                                TRUE ~ type)) |> 
        mutate(end = cumsum(valor),
               end = case_when(variavel == 'Total' ~ 0,
                               TRUE ~ end)) %>%
        mutate(start = lag(end, 1),
               start = case_when(is.na(start) == TRUE ~ 0,
                                 TRUE ~ start))

ttcf <- tcf %>%
  mutate(valor = format(round(valor * 1e-3, 1)), scientific = FALSE, big.mark = ',') %>%
  select(variavel, valor)

gcf <- ggplot(tcf) +
  aes(descricao, fill = type) +
  geom_rect(aes(x = variavel, xmin = id - .45, xmax = id + 0.45, ymin = end, ymax = start)) +
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

ggsave(filename =   here('output',
                         paste0(name,
                                ' - Decomposicao Total Variacao de Caixa - ',
                                '.png')),
       plot     =  gcf,
       device   =  'png',
       scale    =  1,
       width    =  13,
       height   =  4.5,
       units    =  c('cm'),
       dpi      =  500,
       bg       = 'transparent')
