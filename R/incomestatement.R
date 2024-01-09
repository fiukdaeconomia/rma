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
library(viridis)

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

## Gráficos Históricos ----

dates <- seq.Date(from = as.Date('2019-12-31', '%Y-%m-%d'),
                  to   = as.Date('2022-12-31', '%Y-%m-%d'),
                  by = 'year')

for (date in seq_along(dates)) {
  tis <- is |> 
    filter(data == as.Date(dates[date], '%Y-%m-%d')) %>%
    filter(item != 'Devoluções de Vendas',
           item != 'Impostos',
           item != 'Depreciação e Amortização') %>%
    select(item, valor) %>%
    mutate(item = as.character(item),
           item = case_when(item == 'Despesas Financeirasa' ~
                              'Despesas Financeiras',
                            TRUE ~ item),
           item = case_when(item == 'Despesas Financeiras' |
                              item == 'Receitas Financeiras' ~
                              'Resultado Financeiro',
                            item == 'Outras Despesas Operacionais Líquidas' ~
                              'Despesas Gerais e Administrativas',
                            item == 'IR e CSLL Diferidos' |
                              item == 'IR e CSLL Corrente' ~
                              'IR e CSLL',
                            TRUE ~ item)) %>%
    group_by(item) %>%
    summarise(valor = sum(valor)) %>%
    arrange(factor(item, levels = c('Receita Líquida',
                                    'Custo dos Produtos Vendidos',
                                    'Despesas Logísticas',
                                    'Despesas Comerciais',
                                    'Despesas Gerais e Administrativas',
                                    'Outras Despesas Operacionais Líquidas',
                                    'Resultado de Equivalência Patrimonial',
                                    'Resultado Financeiro',
                                    'IR e CSLL'))) %>%
    adorn_totals(name = 'Lucro Líquido', where = 'row', fill = 'Lucro Líquido') %>%
    select(item, valor) %>%
    mutate(valor = case_when(valor == 0 ~ 1e-9,
                             TRUE ~ valor)) %>%
    mutate(item = as.factor(item),
           id   = seq_along(item),
           type = case_when(valor >= 0 ~ 'in',
                            valor <  0 ~ 'out'),
           type = case_when(item == 'Lucro Líquido' ~
                              'net',
                            TRUE ~ type)) %>%
    mutate(end = cumsum(valor),
           end = case_when(item == 'Lucro Líquido' ~ 0,
                           TRUE ~ end)) %>%
    mutate(start = lag(end, 1),
           start = case_when(is.na(start) == TRUE ~ 0,
                             TRUE ~ start)) %>%
    mutate(item = fct_reorder(item, id)) %>%
    arrange(-row_number()) %>%
    mutate(id = seq_along(id)) %>%
    mutate(item = fct_reorder(item, id))
  
  ttis <- tis %>%
    mutate(valor = format(round(valor * 1e-3, 1)), scientific = FALSE, big.mark = ',') %>%
    select(item, valor)
  
  print(dates[date])
  print(ttis)
  
  gis <- ggplot(tis) +
    aes(item, fill = type) +
    geom_rect(aes(x = item, xmin = id - .45, xmax = id + 0.45, ymin = end, ymax = start)) +
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
                                  ' - DRE Histórica - ',
                                  year(as.Date(dates[date])),
                                  '.png')),
         plot     =  gis,
         device   =  'png',
         scale    =  1,
         width    =  3.5,
         height   =  6.5,
         units    =  c('cm'),
         dpi      =  500,
         bg       = 'transparent')
  
}

# Histórico Receitas e Lucro Líquido ----

ll <- is %>%
  filter(data > as.Date('2019-01-01', '%Y-%m-%d')) %>%
  mutate(data = year(data)) %>%
  group_by(data) %>%
  summarise(lucro = sum(valor))

gi <- is %>%
  filter(data > as.Date('2019-01-01', '%Y-%m-%d')) %>%
  mutate(data = year(data)) %>%
  filter(categoria == 'Receita Líquida') %>%
  select(data, valor) %>%
  rename(receita = 'valor') %>%
  left_join(ll, by = 'data') %>%
  pivot_longer(cols = -data)

## Gráfico ----

ghis <- gi %>%
  ggplot(aes(x = data, y = value, fill = name)) +
  geom_bar(position = position_dodge2(reverse = TRUE),
           stat = 'identity') +
  scale_fill_manual(values = c('lucro' = r2c_amarelo_1,
                               'receita' = r2c_azul_1)) +
  geom_hline(yintercept = 0, color = 'red', linetype = 'dashed', size = .3) +
  scale_x_continuous(breaks = pretty_breaks(n = 4)) +
  theme_minimal() +
  theme(panel.background = element_rect(fill     =  'transparent',
                                        colour   =  'transparent'),
        axis.line.y          =   element_blank(),
        axis.line.x          =   element_line(colour = 'black'),
        panel.grid.major     =   element_blank(),
        panel.grid.minor     =   element_blank(),
        legend.title         =   element_blank(),
        legend.text          =   element_text(size = 10),
        legend.position      =   c("none"),
        legend.key.size      =   unit(1,"cm"),
        axis.text.x          =   element_text(size = 8, color = 'black', face= 'italic'),
        axis.text.y          =   element_blank(),
        axis.title.x         =   element_blank(),
        axis.title.y         =   element_blank(),
        rect                 =   element_rect(fill = "transparent"),
        plot.margin          =   unit(c(.1,.1,.1,.1),"cm"))

### Armazenar o Gráfico ----

ggsave(filename =   here(output,
                         paste0(name,
                                ' - Historico Receitas e Lucros',
                                '.png')),
       plot     =  ghis,
       device   =  'png',
       scale    =  1,
       width    =  17.14,
       height   =  6.5,
       units    =  c('cm'),
       dpi      =  500,
       bg       = 'transparent')

# Decomposição do Lucro Líquido Agregado ----


## Dados ----

tis <- is %>%
  filter(data >= as.Date('2019-01-01', '%Y-%m-%d')) %>%
  filter(item != 'Devoluções de Vendas',
         item != 'Impostos',
         item != 'Depreciação e Amortização') %>%
  select(item, valor) %>%
  mutate(item = as.character(item),
         item = case_when(item == 'Despesas Financeirasa' ~
                            'Despesas Financeiras',
                          TRUE ~ item),
         item = case_when(item == 'Despesas Financeiras' |
                            item == 'Receitas Financeiras' ~
                            'Resultado Financeiro',
                          item == 'Outras Despesas Operacionais Líquidas' ~
                            'Despesas Gerais e Administrativas',
                          item == 'IR e CSLL Diferidos' |
                            item == 'IR e CSLL Corrente' ~
                            'IR e CSLL',
                          TRUE ~ item)) %>%
  group_by(item) %>%
  summarise(valor = sum(valor)) %>%
  arrange(factor(item, levels = c('Receita Líquida',
                                  'Custo dos Produtos Vendidos',
                                  'Despesas Logísticas',
                                  'Despesas Comerciais',
                                  'Despesas Gerais e Administrativas',
                                  'Outras Despesas Operacionais Líquidas',
                                  'Resultado de Equivalência Patrimonial',
                                  'Resultado Financeiro',
                                  'IR e CSLL'))) %>%
  adorn_totals(name = 'Lucro Líquido', where = 'row', fill = 'Lucro Líquido') %>%
  select(item, valor) %>%
  select(item, valor) %>%
  mutate(valor = case_when(valor == 0 ~ 1e-9,
                           TRUE ~ valor)) %>%
  mutate(item = as.factor(item),
         id   = seq_along(item),
         type = case_when(valor >= 0 ~ 'in',
                          valor <  0 ~ 'out'),
         type = case_when(item == 'Lucro Líquido' ~
                            'net',
                          TRUE ~ type)) %>%
  mutate(end = cumsum(valor),
         end = case_when(item == 'Lucro Líquido' ~ 0,
                         TRUE ~ end)) %>%
  mutate(start = lag(end, 1),
         start = case_when(is.na(start) == TRUE ~ 0,
                           TRUE ~ start)) %>%
  mutate(item = fct_reorder(item, id)) %>%
  mutate(id = seq_along(id)) %>%
  mutate(item = fct_reorder(item, id))


ttis <- tis %>%
  mutate(valor = format(round(valor * 1e-3, 1)), scientific = FALSE, big.mark = ',') %>%
  select(item, valor)
print(ttis)

## Gráfico ----

gis <- ggplot(tis) +
  aes(item, fill = type) +
  geom_rect(aes(x = item, xmin = id - .45, xmax = id + 0.45, ymin = end, ymax = start)) +
  geom_hline(yintercept = 0, color = 'black', size = .6) +
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

### Armazenar Gráfico ----

ggsave(filename =   here(output,
                         paste0(name,
                                ' - dre acumulada',
                                '.png')),
       plot     =  gis,
       device   =  'png',
       scale    =  1,
       width    =  15.0,
       height   =  5.5,
       units    =  c('cm'),
       dpi      =  500,
       bg       = 'transparent')

# Tabela Histórica DRE ----

tis <- is %>%
  select(data, item, valor) %>%
  mutate(data = paste0(year(data), 'Y')) %>%
  mutate(item = as.character(item),
         item = case_when(item == 'Despesas Financeirasa' ~
                            'Despesas Financeiras',
                          TRUE ~ item)) |>
  pivot_wider(id_cols = data,
              names_from = 'item',
              values_from = 'valor') %>%
  clean_names() %>%
  mutate(lucro_bruto = receita_liquida       +
           custo_dos_produtos_vendidos,
         lucro_antes_do_resultado_financeiro = lucro_bruto  +
           despesas_logisticas                       +
           despesas_comerciais                       +
           despesas_gerais_e_administrativas           +
           outras_despesas_operacionais_liquidas +
           resultado_de_equivalencia_patrimonial,
         lucro_antes_dos_impostos = lucro_antes_do_resultado_financeiro +
           despesas_financeiras +
           receitas_financeiras,
         lucro_liquido =
           lucro_antes_dos_impostos +
           ir_e_csll ) %>%
  pivot_longer(cols = -data) %>%
  pivot_wider(id_cols = name,
              names_from = 'data',
              values_from = 'value') %>%
  mutate(name = str_replace_all(name, '_', " "),
         name = str_to_title(name),
         name = str_replace_all(name, 'Liquido', 'Líquido'),
         name = str_replace_all(name, 'Liquida', 'Líquida'),
         name = str_replace_all(name, 'Liquidos', 'Líquidos'),
         name = str_replace_all(name, 'Liquidas', 'Líquidas'),
         name = str_replace_all(name, 'Servicos', 'Serviços'),
         name = str_replace_all(name, 'Equivalencia', 'Equivalência'),
         name = str_replace_all(name, 'Sc Ps', 'SCPs'),
         name = str_replace_all(name, 'A Scps', 'A SCPs'),
         name = str_replace_all(name, 'Ir E Csll', 'IR E CSLL')) %>%
  adorn_totals(where = 'col') %>%
  filter(Total != 0) %>%
  select(name, Total,
         `2019Y`, `2020Y`, `2021Y`, `2022Y`) %>%
  arrange(factor(name, levels = c(
    'Receita Líquida',
    'Custo Dos Produtos Vendidos',
    'Lucro Bruto',
    'Despesas Logisticas',
    'Despesas Comerciais',
    'Despesas Gerais E Administrativas',
    'Outras Despesas Operacionais Líquidas',
    'Resultado De Equivalência Patrimonial',
    'Lucro Antes Do Resultado Financeiro',
    'Despesas Financeiras',
    'Receitas Financeiras',
    'Lucro Antes Dos Impostos',
    'IR E CSLL',
    'Lucro Líquido'))) %>%
  mutate(across(-name, ~ round(.x * 1e-0,0))) %>%
  mutate(across(-name, ~ as.character(accounting(.x,
                                                 digits = 0L)))) %>%
  mutate(across(-name, ~ case_when(.x == '0.0' ~ '-',
                                   TRUE ~ .x)))
## Flextable DRE ----


flex_is <- tis %>%
  flextable() %>%
  
  border_remove() %>%
  
  # CABEÇALHO
  set_header_labels(name    = 'Item') %>%
  hline(i = 1,  j = 1:ncol(tis),
        part = 'header', fp_border(width = 2, color = r2c_amarelo_1)) %>%
  bold(i = 1,  j = 1, part = 'header', bold = TRUE) %>%
  fontsize(i = 1, j = 1, part = 'header', size = 12) %>%
  italic(i = 1, j = 1, part = 'header', italic = TRUE) %>%
  font(part = 'header', fontname = 'Cambria') %>%
  color(part = 'header', color = r2c_azul_1) %>%
  align(j = 1,  part = 'header',   align = 'left') %>%
  align(j = 2:ncol(tis),  part = 'header',   align = 'center') %>%
  valign(j = 2, part = 'header',   valign = 'bottom') %>%
  padding(part = 'header', padding.bottom = .5) %>%
  
  # LINHAS
  hline(i = nrow(tis) - 1 ,  j = 1:ncol(tis), part = 'body',
        fp_border(width = 1, color = r2c_azul_1)) %>%
  hline(i =          2, j = 1:ncol(tis), part = 'body',
        fp_border(width = 1, color = r2c_azul_1)) %>%
  hline(i =          8, j = 1:ncol(tis), part = 'body',
        fp_border(width = 1, color = r2c_azul_1)) %>%
  hline(i =          11, j = 1:ncol(tis), part = 'body',
        fp_border(width = 1, color = r2c_azul_1)) %>%
  hline(i = nrow(tis), j = 1:ncol(tis), part = 'body',
        fp_border(width = 1.8, color = r2c_azul_1)) %>%
  vline(i = 1:nrow(tis), j = 1:2, part = 'body',
        fp_border(width = 1.2, color = r2c_azul_1)) %>%
  vline(i = 1, j = 1:2, part = 'header',
        fp_border(width = 1.2, color = r2c_azul_1)) %>%
  
  # NEGRITOS
  bold(i =         3,  j = 1:ncol(tis), bold = TRUE) %>%
  bold(i =         9,  j = 1:ncol(tis), bold = TRUE) %>%
  bold(i =        12,  j = 1:ncol(tis), bold = TRUE) %>%
  bold(i = nrow(tis),  j = 1:ncol(tis), bold = TRUE) %>%
  
  # MARGENS
  padding(i =   1:2, j = 1, padding.left = 15) %>%
  padding(i =   4:8, j = 1, padding.left = 15) %>%
  padding(i =   10:11, j = 1, padding.left = 15) %>%
  padding(i =   13, j = 1, padding.left = 15) %>%
  
  # ALINHAMENTOS
  align(j = 1,  part = 'body',   align = 'left') %>%
  align(j = 2:ncol(tis),  part = 'body',   align = 'right') %>%
  width(j = 1,  width = 10.0) %>%
  width(j = 2:ncol(tis),  width = 1.5) %>%
  
  # FONTE
  font(part = 'body', fontname = 'Cambria') %>%
  fontsize(part = 'body', size = 11)

flex_is


### Armazenar Tabela ----

save_as_image(flex_is,
              here(output,
                   paste0(name,
                          ' - DRE Histórica',
                          '.png')),
              zoom = .8)

# Gráfico das Margens ----

## Dados ----

depr <- fs %>%
  filter(demonstrativo == 'Demonstrações dos Fluxos de Caixa',
         categoria     == 'Ajuste de Itens Não-Caixa',
         item          == 'Depreciação e Amortização') %>%
  group_by(data) %>%
  summarise(valor = sum(valor)) %>%
  mutate(item = 'Depreciação e Amortização') %>%
  select(data, item, valor)

mis <- is %>%
  select(data, item, valor) %>%
  mutate(item = as.character(item),
         item = case_when(item == 'Despesas Financeirasa' ~
                            'Despesas Financeiras',
                          TRUE ~ item)) |>
  bind_rows(depr) %>%
  pivot_wider(id_cols = data,
              names_from = 'item',
              values_from = 'valor') %>%
  clean_names() %>%
  mutate(lucro_bruto = receita_liquida       +
           custo_dos_produtos_vendidos,
         lucro_antes_do_resultado_financeiro = lucro_bruto  +
           despesas_logisticas                         +
           despesas_comerciais +
           despesas_gerais_e_administrativas           +
           outras_despesas_operacionais_liquidas       +
          resultado_de_equivalencia_patrimonial,
         lucro_antes_dos_impostos = lucro_antes_do_resultado_financeiro +
           despesas_financeiras +
           receitas_financeiras,
         lucro_liquido =
           lucro_antes_dos_impostos +
           ir_e_csll,  
         ebitda = lucro_antes_do_resultado_financeiro +
           depreciacao_e_amortizacao) %>%
  mutate(mg_bruta = lucro_bruto/receita_liquida * 100,
         mg_opera = lucro_antes_do_resultado_financeiro/receita_liquida * 100,
         mg_liqui = lucro_liquido/receita_liquida * 100,
         mg_ebitda = ebitda/receita_liquida * 100) %>%
  select(data, mg_bruta, mg_opera, mg_liqui, mg_ebitda) %>%
  pivot_longer(cols = -data)

## Gráfico ----

gmis <-   mis %>%
  mutate(data = year(data)) %>%
  ggplot(aes(x = data, y = value, color = name, label = round(value, 0))) +
  geom_hline(yintercept = 0, color = r2c_cinza_2) +
  geom_vline(xintercept = 2019,
             color = r2c_cinza_2,
             alpha = 0.5) +
  geom_vline(xintercept = 2020,
             color = r2c_cinza_2,
             alpha = 0.5) +
  geom_vline(xintercept = 2021,
             color = r2c_cinza_2,
             alpha = 0.5) +
  geom_line() +
  scale_color_manual(values = c('#F3954B', '#B961AA', '#1E178A', r2c_amarelo_1,
                                r2c_vermelho_1)) +
  scale_x_continuous(breaks = pretty_breaks(n = 3)) +
  geom_point(size = 5) +
  geom_text(color = "white", size = 2.3) +
  scale_fill_viridis() +
  theme_minimal() +
  theme(legend.position = 'none',
        panel.background = element_rect(fill     =  'transparent',
                                        colour   =  'transparent'),
        panel.grid.major     =   element_blank(),
        axis.text.x          =   element_text(size = 8, color = 'black', face = 'italic'),
        axis.text.y          =   element_text(size = 8, color = 'black', face = 'italic'),
        axis.title           =  element_blank(),
        panel.grid.minor     =   element_blank(),
        rect                 =   element_rect(fill = "transparent"),
        plot.margin          =   unit(c(.1,.1,.1,.1),"cm"))


### Armazenar o Gráfico ----

ggsave(filename =   here(output,
                         paste0(name,
                                ' - Margens DRE Históricas',
                                '.png')),
       plot     =  gmis,
       device   =  'png',
       scale    =  1,
       width    =  14.0,
       height   =  7.5,
       units    =  c('cm'),
       dpi      =  500,
       bg       = 'transparent')
