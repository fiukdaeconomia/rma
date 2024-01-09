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
library(treemap)
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

bs <- fs |>
  filter(demonstrativo == 'Balanço Patrimonial')

# Gráficos Histórico Balanços ----

dates <- as.Date(unique(bs$data))

map(dates, function(date){
  
  ggsave(filename = here(output,
                         glue('balanço - {name} - {date}.png')),
         device   = 'png',
         scale    =  1,
         width    =  4.0,
         height   =  6.4,
         units    =  c('cm'),
         dpi      =  500,
         bg       = 'transparent',
         bs |>
           filter(data == date) |>
           mutate(valor = valor * 1e-6)  |>
           group_by(categoria, classe) |>
           summarise(valor = sum(valor)) |>
           ggplot() +
           aes(x = categoria, y = valor, fill = classe) +
           scale_fill_manual(values = c(r2c_amarelo_2,
                                        r2c_amarelo_1,
                                        r2c_roxo_1,
                                        r2c_roxo_2,
                                        r2c_azul_1)) +
           geom_col(alpha = 1, col = 'white') +
           geom_hline(yintercept = 0,
                      color = r2c_cinza_1,
                      linetype = 'dashed',
                      size = .6) +
           geom_segment(x = 1.5,
                        y = 0,
                        xend = 1.5,
                        yend = 13 ,
                        color = r2c_cinza_1,
                        size = .6) +
           scale_x_discrete(position = 'top') +
           scale_y_continuous(limits = c(0, 13)) +
           theme(legend.position = 'none',
                 axis.line.y  = element_blank(),
                 axis.line.x  = element_blank(),
                 axis.text.y  = element_blank(),
                 axis.text.x  = element_blank(),
                 axis.title   = element_blank(),
                 axis.ticks   = element_blank(),
                 panel.background = element_rect(fill     =  'transparent',
                                                 colour   =  'transparent'),
                 panel.grid.major     =   element_blank(),
                 panel.grid.minor     =   element_blank(),
                 rect                 =   element_rect(fill = "transparent"),
                 plot.margin          =   unit(c(.1,.1,.1,.1),"cm")))
})

# Gráficos de Liquidez ----

wbs <- bs |>
  group_by(data, classe) |>
  summarise(valor = sum(valor)) |>
  pivot_wider(names_from = 'classe',
              values_from = 'valor')  |>
  clean_names()

## Cálculo de Indicadores ----

liq <- bs |>
  filter(data >= as.Date('2019-01-01',
                         '%Y-%m-%d'))  |>
  filter(classe    == 'Ativo Circulante'     &
           item      == 'Disponibilidades'     |
           classe    == 'Ativo Circulante'     &
           item      == 'Estoques'             |
           classe    == 'Ativo Não Circulante' &
           item      == 'Contas a Receber') |>
  group_by(data, item) |>
  summarise(valor = sum(valor)) |>
  pivot_wider(names_from  = 'item',
              values_from = 'valor') |>
  clean_names() |>
  left_join(wbs, by = 'data') |>
  mutate(liq_cor = ativo_circulante / passivo_circulante,
         liq_sec = (ativo_circulante - 0) / passivo_circulante,
         liq_ime = disponibilidades / passivo_circulante,
         liq_ger = (ativo_circulante + 0) /
           (passivo_circulante + passivo_nao_circulante)) |>
  select(data, liq_cor, liq_sec, liq_ime, liq_ger) |>
  pivot_longer(cols = starts_with('liq'))

## Elaboração do Gráfico ----

gliq <- liq |>
  ggplot(aes(x = data, y = value, color = name)) +
  geom_vline(xintercept = as.Date('2019-12-31'),
             color = r2c_cinza_2,
             alpha = 0.5) +
  geom_vline(xintercept = as.Date('2020-12-31'),
             color = r2c_cinza_2,
             alpha = 0.5) +
  geom_vline(xintercept = as.Date('2021-12-31'),
             color = r2c_cinza_2,
             alpha = 0.5) +
  geom_vline(xintercept = as.Date('2022-12-31'),
             color = r2c_cinza_2,
             alpha = 0.5) +
  geom_line() +
  scale_color_manual(values = c('#F3954B',
                                '#B961AA',
                                '#1E178A',
                                r2c_amarelo_1)) +
  geom_point(size = 4) +
  scale_fill_viridis() +
  theme_minimal() +
  geom_hline(yintercept = 0, color = r2c_cinza_2) +
  theme(legend.position = 'right',
        panel.background = element_rect(fill     =  'transparent',
                                        colour   =  'transparent'),
        panel.grid.major     =   element_blank(),
        axis.text.x          =   element_blank(),
        axis.title           =  element_blank(),
        panel.grid.minor     =   element_blank(),
        rect                 =   element_rect(fill = "transparent"),
        plot.margin          =   unit(c(.1,.1,.1,.1),"cm"))

### Armazenamento do Gráfico ----

ggsave(filename = here(output,
                       glue('liquidez - {name}.png')),
       plot     =  gliq,
       device   =  'png',
       scale    =  1,
       width    =  14.14,
       height   =  6.5,
       units    =  c('cm'),
       dpi      =  500,
       bg       = 'transparent')

# Endividamento ----

## Base de Dados ----

edv <- bs |>
  filter(data > as.Date('2019-01-01'))  |>
  mutate(data = year(data)) |>
  filter(categoria   == 'Ativo'                        &
           classe    == 'Ativo Circulante'             &
           item      == 'Disponibilidades'             |
           categoria == 'Passivo'                      &
           item      == 'Empréstimos e Financiamentos')  |>
  mutate(item = as.character(item),
         item = case_when(classe == 'Passivo Circulante' ~
                            paste0(item, ' CP '),
                          TRUE ~ item))  |>
  select(data,
         item,
         valor) |>
  pivot_wider(id_cols     = 'data',
              names_from  = 'item',
              values_from = 'valor')  |>
  clean_names() |>
  mutate(edv_total = emprestimos_e_financiamentos_cp +
           emprestimos_e_financiamentos) |>
  mutate(edv_liq   = edv_total - disponibilidades) |>
  pivot_longer(cols = -data)

## Gráfico do Endividamento ----

gdiv <- edv |>
  filter(name == 'emprestimos_e_financiamentos' |
           name == 'emprestimos_e_financiamentos_cp') |>
  ggplot() +
  aes(x = data, y = value, fill = name) +
  geom_bar(position = 'stack', stat = 'identity') +
  scale_fill_manual(values = c(r2c_azul_1, r2c_amarelo_1)) +
  scale_y_continuous(labels = unit_format(unit = "MM", scale = 1e-3)) +
  scale_x_continuous(breaks = pretty_breaks(n = 4)) +
  theme_minimal() +
  theme(panel.background = element_rect(fill     =  'transparent',
                                        colour   =  'transparent'),
        axis.line.y          =   element_blank(),
        axis.line.x          =   element_line(colour =  'black'),
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

### Armazenamento do Gráfico ----

ggsave(filename = here(output,
                       glue('endividamento historico - {name}.png')),
       plot     =  gdiv,
       device   =  'png',
       scale    =  1,
       width    =  14.14,
       height   =  6.5,
       units    =  c('cm'),
       dpi      =  500,
       bg       = 'transparent')

# Tabelas do Balanço ----

## Ativos ----

### Ativo Circulante ----

itens_ac <- c('Disponibilidades',
              'Contas a Receber',
              'Estoques',
              'Tributos a Recuperar',
              'Adiantamentos')

ac <- bs |>
  mutate(data  = as.character(year(data))) |>
  filter(classe == 'Ativo Circulante',
         data      == '2022') |>
  mutate(item = as.character(item),
         item = if_else(item %in% itens_ac == TRUE,
                        item,
                        'Outros'))  |>
  group_by(data, item) |>
  summarise(valor = sum(valor))  |>
  arrange(factor(item, levels =  c('Disponibilidades',
                                   'Contas a Receber',
                                   'Estoques',
                                   'Tributos a Recuperar',
                                   'Adiantamentos',
                                   'Outros'))) |>
  pivot_wider(names_from = 'data', values_from = 'valor')  |>
  adorn_totals(where = 'row', name = 'Total Ativo Circulante')  |>
  mutate(`2022` = format(round(`2022`), big.mark = ','))

### Ativo Não Circulante ----

itens_anc <- c('Partes Relacionadas',
               'Ativo Fiscal Diferido',
               'Tributos a Recuperar',
               'Imobilizado',
               'Intangível')
anc <- bs |>
  mutate(data  = as.character(year(data))) |>
  filter(classe == 'Ativo Não Circulante',
         data      == '2022') |>
  mutate(item = as.character(item),
         item = if_else(item %in% itens_anc == TRUE,
                        item,
                        'Outros')) |>
  group_by(data, item) |>
  summarise(valor = sum(valor))  |>
  arrange(factor(item, levels = c('Partes Relacionadas',
                                  'Ativo Fiscal Diferido',
                                  'Tributos a Recuperar',
                                  'Imobilizado',
                                  'Intangível',
                                  'Outros')))  |>
  pivot_wider(names_from = 'data', values_from = 'valor')  |>
  adorn_totals(where = 'row', name = 'Total Ativo Não Circulante')  |>
  mutate(`2022` = format(round(`2022`), big.mark = ','))

### Ativo Total ----

at <- bs |>
  mutate(data  = as.character(year(data))) |>
  filter(categoria == 'Ativo',
         data      == '2022') |>
  group_by(data, categoria) |>
  summarise(valor = sum(valor))  |>
  pivot_wider(names_from = 'data', values_from = 'valor')  |>
  adorn_totals(where = 'row', name = 'Total Ativo') |>
  rename('item' = categoria) |>
  slice(2) |>
  mutate(`2022` = format(round(`2022`), big.mark = ','))

### Tabela Ativos ----

ativos <- rbind(c('Ativo Circulante',''),
                ac,
                c('', ''),
                c('Ativo Não Circulante', ''),
                anc,
                c('',''),
                c('',''),
                at)  |>
  mutate(`2022` = format(`2022`, big.mark = ','))

flex_ativos <- ativos |>
  flextable() |>
  
  border_remove() |>
  
  # CABEÇALHO
  set_header_labels(item     = 'ATIVOS',
                    `2021`     = 'BRL') |>
  hline(i = 1,  j = 1:2,
        part = 'header',
        fp_border(width = 2,
                  color = r2c_amarelo_1)) |>
  bold(i = 1,  j = 1, part = 'header', bold = TRUE) |>
  fontsize(i = 1, j = 1, part = 'header', size = 12) |>
  fontsize(i = 1, j = 2, part = 'header', size = 7) |>
  italic(i = 1, j = 2, part = 'header', italic = TRUE) |>
  font(part = 'header', fontname = 'Cambria') |>
  color(part = 'header', color = r2c_azul_1) |>
  align(j = 1,  part = 'header',   align = 'left') |>
  align(j = 2,  part = 'header',   align = 'right') |>
  valign(j = 2, part = 'header',   valign = 'bottom') |>
  padding(part = 'header', padding.bottom = .5) |>
  
  # LINHAS
  hline(i = 7,  j = 2, part = 'body', fp_border(width = 1,
                                                color = r2c_azul_1)) |>
  hline(i = 8,  j = 2, part = 'body', fp_border(width = 1,
                                                color = r2c_azul_1)) |>
  hline(i = 16,  j = 2, part = 'body', fp_border(width = 1,
                                                 color = r2c_azul_1)) |>
  hline(i = 17, j = 2, part = 'body', fp_border(width = 1,
                                                color = r2c_azul_1)) |>
  hline(i = 19, j = 2, part = 'body', fp_border(width = 2,
                                                color = r2c_azul_1)) |>
  hline(i = 20, j = 2, part = 'body', fp_border(width = 2,
                                                color = r2c_azul_1)) |>
  
  # NEGRITOS
  bold(i =  8,  j = 1:2, bold = TRUE) |>
  bold(i = 17,  j = 1:2, bold = TRUE) |>
  bold(i = 20,  j = 1:2, bold = TRUE) |>
  
  # MARGENS
  padding(i = 2:7, j = 1, padding.left = 15) |>
  padding(i = 11:16, j = 1, padding.left = 15) |>
  
  # ALINHAMENTOS
  align(j = 1,  part = 'body',   align = 'left') |>
  align(j = 2,  part = 'body',   align = 'right') |>
  width(j = 1,  width = 3.0) |>
  width(j = 2,  width = 1.0) |>
  
  # FONTE
  font(part = 'body', fontname = 'Cambria') |>
  fontsize(part = 'body', size = 11)

#### Armazenar Tabela ----

save_as_image(flex_ativos,
              here(output,
                   glue('Tabela Ativos - {name}.png')),
              zoom = 1)

## Passivos ----

### Passivo Circulante ----

itens_pc <- c('Fornecedores',
              'Empréstimos e Financiamentos',
              'Obrigações Tributárias')

pc <- bs |>
  mutate(data  = as.character(year(data))) |>
  filter(classe == 'Passivo Circulante',
         data      == '2022') |>
  mutate(item = as.character(item),
         item = if_else(item %in% itens_pc == TRUE,
                        item,
                        'Outros')) |>
  group_by(data, item) |>
  summarise(valor = sum(valor)) |>
  arrange(factor(item, levels = c('Fornecedores',
                                  'Empréstimos e Financiamentos',
                                  'Obrigações Tributárias',
                                  'Outros'))) |>
  pivot_wider(names_from = 'data', values_from = 'valor') |>
  adorn_totals(where = 'row', name = 'Total Passivo Circulante') |>
  mutate(`2022` = format(round(`2022`), big.mark = ','))

### Passivo Não Circulante ----

itens_pnc <- c('Partes Relacionadas',
               'Empréstimos e Financiamentos',
               'Obrigações Tributárias',
               'Passivo Fiscal Diferido')

pnc <- bs |>
  mutate(data  = as.character(year(data))) |>
  filter(classe == 'Passivo Não Circulante',
         data      == '2022') |>
  mutate(item = as.character(item),
         item = if_else(item %in% itens_pnc == TRUE,
                        item,
                        'Outros')) |>
  group_by(data, item) |>
  summarise(valor = sum(valor)) |>
  arrange(factor(item, levels = c('Partes Relacionadas',
                                  'Empréstimos e Financiamentos',
                                  'Obrigações Tributárias',
                                  'Passivo Fiscal Diferido',
                                  'Outros'))) |>
  pivot_wider(names_from = 'data', values_from = 'valor') |>
  adorn_totals(where = 'row', name = 'Total Passivo Não Circulante') |>
  mutate(`2022` = format(round(`2022`), big.mark = ','))

### Patrimônio Líquido ----

itens_pl <- c('Capital Social',
              'Resultados Acumulados')
pl <- bs |>
  mutate(data  = as.character(year(data))) |>
  filter(classe == 'Patrimônio Líquido',
         data      == '2022') |>
  mutate(item = as.character(item),
         item = if_else(item %in% itens_pl == TRUE,
                        item,
                        'Outros')) |>
  group_by(data, item) |>
  summarise(valor = sum(valor)) |>
  arrange(factor(item, levels = c('Capital Social',
                                  'Resultados Acumulados',
                                  'Outros'))) |>
  pivot_wider(names_from = 'data', values_from = 'valor') |>
  adorn_totals(where = 'row', name = 'Total Patrimônio Líquido') |>
  mutate(`2022` = format(round(`2022`), big.mark = ','))

### Passivo Total e PL Total

ps <- bs %>%
  mutate(data  = as.character(year(data))) %>%
  filter(categoria == 'Passivo',
         classe    != 'Patrimônio Líquido',
         data      == '2022') %>%
  group_by(data, categoria) %>%
  summarise(valor = sum(valor)) %>%
  pivot_wider(names_from = 'data', values_from = 'valor') %>%
  adorn_totals(where = 'row', name = 'Total Passivo') %>%
  rename('item' = categoria) %>%
  slice(2) %>%
  mutate(`2022` = format(round(`2022`), big.mark = ','))

pspl <- bs %>%
  mutate(data  = as.character(year(data))) %>%
  mutate(categoria = as.character(categoria),
         categoria = case_when(classe == 'Patrimônio Líquido' ~
                                 'Patrimônio Líquido',
                               TRUE ~ categoria)) |>
  filter(categoria == 'Passivo' |
           categoria == 'Patrimônio Líquido',
         data      == '2022') %>%
  group_by(data, categoria) %>%
  summarise(valor = sum(valor)) %>%
  pivot_wider(names_from = 'data', values_from = 'valor') %>%
  adorn_totals(where = 'row', name = 'Total Passivo e Patrimônio Líquido') %>%
  rename('item' = categoria) %>%
  slice(3) %>%
  mutate(`2022` = format(round(`2022`), big.mark = ','))

### Tabela Passivos ----


passivos <- rbind(c('Passivo Circulante',''),
                  pc,
                  c('', ''),
                  c('Passivo Não Circulante', ''),
                  pnc,
                  c('',''),
                  c('Patrimônio Líquido', ''),
                  pl,
                  c('',''),
                  ps,
                  pspl)  %>%
  mutate(`2022` = format(`2022`, big.mark = ','))

flex_passivos <- passivos %>%
  flextable() %>%
  
  border_remove() %>%
  
  # CABEÇALHO
  set_header_labels(item     = 'PASSIVOS',
                    `2021`     = 'BRL') %>%
  hline(i = 1,  j = 1:2, part = 'header', fp_border(width = 2,
                                                    color = r2c_amarelo_1)) %>%
  bold(i = 1,  j = 1, part = 'header', bold = TRUE) %>%
  fontsize(i = 1, j = 1, part = 'header', size = 12) %>%
  fontsize(i = 1, j = 2, part = 'header', size = 7) %>%
  italic(i = 1, j = 2, part = 'header', italic = TRUE) %>%
  font(part = 'header', fontname = 'Cambria') %>%
  color(part = 'header', color = r2c_azul_1) %>%
  align(j = 1,  part = 'header',   align = 'left') %>%
  align(j = 2,  part = 'header',   align = 'right') %>%
  valign(j = 2, part = 'header',   valign = 'bottom') %>%
  padding(part = 'header', padding.bottom = .5) %>%
  
  # LINHAS
  hline(i = 5,  j = 2, part = 'body', fp_border(width = 1,
                                                color = r2c_azul_1)) %>%
  hline(i = 6,  j = 2, part = 'body', fp_border(width = 1,
                                                color = r2c_azul_1)) %>%
  hline(i = 13,  j = 2, part = 'body', fp_border(width = 1,
                                                 color = r2c_azul_1)) %>%
  hline(i = 14, j = 2, part = 'body', fp_border(width = 1,
                                                color = r2c_azul_1)) %>%
  hline(i = 19,  j = 2, part = 'body', fp_border(width = 1,
                                                 color = r2c_azul_1)) %>%
  hline(i = 20, j = 2, part = 'body', fp_border(width = 1,
                                                color = r2c_azul_1)) %>%
  hline(i = 21, j = 2, part = 'body', fp_border(width = 2,
                                                color = r2c_azul_1)) %>%
  hline(i = 23, j = 2, part = 'body', fp_border(width = 2,
                                                color = r2c_azul_1)) %>%
  
  # NEGRITOS
  bold(i = 6,  j = 1:2, bold = TRUE) %>%
  bold(i = 14,  j = 1:2, bold = TRUE) %>%
  bold(i = 20,  j = 1:2, bold = TRUE) %>%
  bold(i = 21:23,  j = 1:2, bold = TRUE) %>%
  
  # MARGENS
  padding(i = 2:5, j = 1, padding.left = 15) %>%
  padding(i = 9:13, j = 1, padding.left = 15) %>%
  padding(i = 17:19, j = 1, padding.left = 15) %>%
  
  # ALINHAMENTOS
  align(j = 1,  part = 'body',   align = 'left') %>%
  align(j = 2,  part = 'body',   align = 'right') %>%
  width(j = 1,  width = 3.0) %>%
  width(j = 2,  width = 1.0) %>%
  
  # FONTE
  font(part = 'body', fontname = 'Cambria') %>%
  fontsize(part = 'body', size = 11)

#### Armazenar Tabela ----

save_as_image(flex_passivos,
              here(output,
                   glue('Tabela Passivos - {name}.png')),
              zoom = 1)

# Breakdown do Balanço Atual ----

## Ativo ----

tat <- bs %>%
  filter(categoria == 'Ativo',
         data   == as.Date('2022-12-31', '%Y-%m-%d'))

treemap(tat,
        index = c('classe','item'),
        vSize = 'valor',
        type = 'index',
        fontsize.labels  = c(1,10),
        fontcolor.labels = c("transparent","transparent"),
        fontface.labels  = c(2,1),
        align.labels     = list(c("center", "center"),
                                c("right", "bottom")),
        overlap.labels   = 0.5,
        inflate.labels   = F,
        border.col       = c("white","white"),
        border.lwds      = c(7,2),
        palette          = c(r2c_amarelo_2, r2c_amarelo_1),
        fontsize.title = 0)

## Passivo ----

tpas <- bs %>%
  filter(categoria == 'Passivo',
         classe != 'Patrimônio Líquido',
         data   == as.Date('2022-12-31', '%Y-%m-%d'))

treemap(tpas,
        index = c('classe','item'),
        vSize = 'valor',
        type = 'index',
        fontsize.labels  = c(1,10),
        fontcolor.labels = c("transparent","transparent"),
        fontface.labels  = c(2,1),
        align.labels     = list(c("center", "center"), c("right", "bottom")),
        overlap.labels   = 0.5,
        inflate.labels   = F,
        border.col       = c("white","white"),
        border.lwds      = c(7,2),
        palette          = c(r2c_roxo_1, r2c_roxo_2, r2c_azul_1),
        fontsize.title = 0)

