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

dmpl <- fs |>
  filter(demonstrativo == 'Demonstração das Mutações do Patrimônio Líquido')

bs <- fs |>
  filter(demonstrativo == 'Balanço Patrimonial')

### Selecionar PLs ----

pls <- bs |>
  mutate(categoria = as.character(categoria),
         classe    = as.character(classe)) |>
  filter(categoria == 'Passivo',
         classe    == 'Patrimônio Líquido',
         data      <= as.Date('2022-12-31')) |>
  group_by(data) |>
  summarise(EoP = sum(valor)) |>
  mutate(BoP = lag(EoP)) |>
  pivot_longer(cols = -data,
               names_to = 'item',
               values_to = 'valor')

pls$valor[2] <- 3684473

# Tabela do DMPL ----

tdmpl <- dmpl %>%
  select(data, item, valor) %>%
  filter(valor != 0,
         data >= as.Date('2017-12-31', '%Y-%m-%d')) %>%
  bind_rows(pls) %>%
  mutate(data = paste0(as.character(year(data)), 'Y')) %>%
  pivot_wider(id_cols = item,
              names_from = data,
              values_from = valor,
              values_fill = 0) %>%
  adorn_totals(where = 'col') %>%
  select(item, Total, 
         `2019Y`, `2020Y`, `2021Y`, `2022Y`) %>%
  arrange(factor(item, levels = c(
    'BoP',
    'Variação na Participação dos Minoritários',
    'Ajustes de Conversão de Balanço em Controlada',
    'Ajuste IRPJ/CSLL Diferidos de AAP',
    'Resultado Líquido do Exercício',
    'Dividendos Distribuídos',
    'EoP'))) %>%
  mutate(across(-item, ~ round(.x * 1e-3, 1))) %>%
  mutate(across(-item, ~ as.character(accounting(.x,
                                                 digits = 1L)))) %>%
  mutate(across(-item, ~ case_when(.x == '0.0' ~ '-',
                                   TRUE ~ .x)))



tdmpl$Total[1] <- ''
tdmpl$Total[nrow(tdmpl)] <- ''


## Gerar Tabela ----


flex_dmpl <- tdmpl %>%
  flextable() %>%
  
  border_remove() %>%
  
  # CABEÇALHO
  set_header_labels(item     = 'Item') %>%
  hline(i = 1,  j = 1:ncol(tdmpl), part = 'header', fp_border(width = 2,
                                                              color = r2c_amarelo_1)) %>%
  bold(i = 1,  j = 1, part = 'header', bold = TRUE) %>%
  fontsize(i = 1, j = 1, part = 'header', size = 12) %>%
  italic(i = 1, j = 1, part = 'header', italic = TRUE) %>%
  font(part = 'header', fontname = 'Cambria') %>%
  color(part = 'header', color = r2c_azul_1) %>%
  align(j = 1,  part = 'header',   align = 'left') %>%
  align(j = 2:ncol(tdmpl),  part = 'header',   align = 'center') %>%
  valign(j = 2, part = 'header',   valign = 'bottom') %>%
  padding(part = 'header', padding.bottom = .5) %>%
  
  # LINHAS
  hline(i = 1,  j = 1:ncol(tdmpl), part = 'body', fp_border(width = 1,
                                                            color = r2c_azul_1)) %>%
  
  hline(i = nrow(tdmpl) - 1 ,  j = 1:ncol(tdmpl), part = 'body', fp_border(width = 1,
                                                                           color = r2c_azul_1)) %>%
  hline(i = nrow(tdmpl), j = 1:ncol(tdmpl), part = 'body', fp_border(width = 1,
                                                                     color = r2c_azul_1)) %>%
  vline(j = 1:2, i = 1:nrow(tdmpl), part = 'body', fp_border(width = 1,
                                                             color = r2c_azul_1)) %>%
  vline(j = 1:2, i = 1, part = 'header', fp_border(width = 1,
                                                   color = r2c_azul_1)) %>%
  # NEGRITOS
  bold(i = 1,  j = 1:ncol(tdmpl), bold = TRUE) %>%
  bold(i = nrow(tdmpl) ,  j = 1:ncol(tdmpl), bold = TRUE) %>%
  bold(j = 2, bold = TRUE) %>%
  
  # MARGENS
  padding(i = 2:(nrow(tdmpl) - 1) , j = 1, padding.left = 15) %>%
  
  # ALINHAMENTOS
  align(j = 1,  part = 'body',   align = 'left') %>%
  align(j = 2:ncol(tdmpl),  part = 'body',   align = 'right') %>%
  width(j = 1,  width = 12.0) %>%
  width(j = 2:ncol(tdmpl),  width = 1.5) %>%
  
  # FONTE
  font(part = 'body', fontname = 'Cambria') %>%
  fontsize(part = 'body', size = 11)

#### Armazenar Tabela ----

save_as_image(flex_dmpl,
              here(output,
                   paste0(name,
                          ' - DMPL.png')),
              zoom = 1)


# Decomposição do DMPL ----

tdmpl <- dmpl %>%
  filter(data >= as.Date('2017-12-31', '%Y-%m-%d'),
         valor != 0) %>%
  group_by(item) %>%
  summarise(valor = sum(valor)) %>%
  group_by(item) %>%
  summarise(valor = sum(valor)) %>%
  arrange(desc(valor)) %>%
  adorn_totals(name = 'Variação Total', where = 'row', fill = 'Variação Total') %>%
  select(item, valor) %>%
  mutate(valor = case_when(valor == 0 ~ 1e-9,
                           TRUE ~ valor)) %>%
  mutate(item = as.factor(item),
         id   = seq_along(item),
         type = case_when(valor >= 0 ~ 'in',
                          valor <  0 ~ 'out'),
         type = case_when(item == 'Variação Total' ~
                            'net',
                          TRUE ~ type)) %>%
  mutate(end = cumsum(valor),
         end = case_when(item == 'Variação Total' ~ 0,
                         TRUE ~ end)) %>%
  mutate(start = lag(end, 1),
         start = case_when(is.na(start) == TRUE ~ 0,
                           TRUE ~ start)) %>%
  mutate(item = fct_reorder(item, id)) %>%
  mutate(id = seq_along(id)) %>%
  mutate(item = fct_reorder(item, id))

ttdmpl <- tdmpl %>%
  mutate(valor = format(round(valor * 1e-3, 1)), scientific = FALSE, big.mark = ',') %>%
  select(item, valor)
print(ttdmpl)

gdmpl <- ggplot(tdmpl) +
  aes(item, fill = type) +
  geom_rect(aes(x = item, xmin = id - .45, xmax = id + 0.45, ymin = end, ymax = start)) +
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
                                ' - DMPL Consolidado',
                                '.png')),
       plot     =  gdmpl,
       device   =  'png',
       scale    =  1,
       width    =  12.0,
       height   =  5.0,
       units    =  c('cm'),
       dpi      =  500,
       bg       = 'transparent')

# Gráfico do Histórico do PL ----

hpl <- bs %>%
  mutate(classe = as.character(classe)) |>
  filter(classe        == 'Patrimônio Líquido',
         data          <= as.Date('2022-12-31')) %>%
  group_by(data) %>%
  summarise(valor = sum(valor))


gpl <- hpl %>%
  ggplot() +
  aes(x = data,
      y = valor) +
  geom_line(colour = r2c_azul_1, size = 1.3) +
  geom_segment(aes(x = hpl$data[nrow(hpl)],
                   y = hpl$valor[nrow(hpl)],
                   xend = hpl$data[nrow(hpl)] - as.difftime(4*1, unit = 'weeks'),
                   yend = hpl$valor[nrow(hpl)] +  10e4),
               colour = r2c_amarelo_1) +
  geom_point(aes(x = hpl$data[nrow(hpl)],
                 y = hpl$valor[nrow(hpl)]),
             size = 3,
             colour = r2c_azul_1) +
  geom_label(aes(x = hpl$data[nrow(hpl)] - as.difftime(4*1, unit = 'weeks'),
                 y = hpl$valor[nrow(hpl)] + 10e4,
                 label = paste0('BRL ',
                                format(round(hpl$valor[nrow(hpl)] * 1e-3),
                                       big.mark = ','),
                                'MM')),
             fill = 'white',
             colour = r2c_azul_1,
             size = 3) +
  scale_x_date(breaks = '12 months',
               date_labels = '%Y') +
  scale_y_continuous(labels = unit_format(unit = "MM",
                                          scale = 1e-6),
                     breaks = pretty_breaks(n = 6)) +
  ylab('Patrimônio Líquido (BRL)') +
  theme_minimal() +
  theme(panel.background = element_rect(fill     =  'transparent',
                                        colour   =  'transparent'),
        axis.line.y          =   element_line(colour =  'black'),
        axis.line.x          =   element_line(colour =  'black'),
        panel.grid.major     =   element_blank(),
        panel.grid.minor     =   element_blank(),
        legend.title         =   element_blank(),
        legend.text          =   element_text(size = 10),
        legend.position      =   c("none"),
        legend.key.size      =   unit(1,"cm"),
        axis.text.x          =   element_blank(),
        axis.text.y          =   element_text(size = 8,
                                              colour = 'black'),
        axis.title.x         =   element_blank(),
        axis.title.y         =   element_blank(),
        rect                 =   element_rect(fill = "transparent"),
        plot.margin          =   unit(c(.1,.1,.1,.1),"cm")
  )

## Armazenar o Gráfico ----

ggsave(filename =   here(output,
                         paste0(name,
                                ' - Historico PL ',
                                '.png')),
       plot     =  gpl,
       device   =  'png',
       scale    =  1,
       width    =  20.5,
       height   =  6.0,
       units    =  c('cm'),
       dpi      =  500,
       bg       = 'transparent')
