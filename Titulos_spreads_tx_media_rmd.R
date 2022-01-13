#Diretório raiz
#diretorio = input('Insira o diretório onde está a planilha fonte:') + '/'
#getwd()
#setwd("C:\\Users\\User\\Downloads")

#Importação de pacotes
library(dplyr)
library(BETS)
library(rio)
library(ggplot2)
library(ggrepel)

#Montando funções
para_data <- function(df, nomes_colunas){
  for(i in 1:length(nomes_colunas)){
    df[,nomes_colunas[i]] <- as.Date(df[,nomes_colunas[i]], format = "%d/%m/%Y")
  }
  return(df)
}

para_numero <- function(df, nomes_colunas){
  for(i in 1:length(nomes_colunas)){
    df[,nomes_colunas[i]] <- as.numeric(gsub(",", ".", df[,nomes_colunas[i]], fixed = T))
  }
  return(df)
}

#Importação de arquivos
#nome_arquivo_ltn <- as.character(readline("Insira nome do arquivo excel da LTN a ser importado:")) %>% paste(".csv", sep = "")
nome_arquivo_ltn <- "histórico leilões LTN.csv"
ltn <- import(nome_arquivo_ltn, skip = 7, encoding = "Latin-1")
ltn <- ltn[-1,]

#nome_arquivo_ntn_f <- as.character(readline("Insira nome do arquivo excel da NTN-F a ser importado:")) %>% paste(".csv", sep = "")
nome_arquivo_ntn_f <- "histórico leilões NTN-F.csv"
ntn_f <- import(nome_arquivo_ntn_f, skip = 7, encoding = "Latin-1")
ntn_f <- ntn_f[-1,]

meta_selic <- BETSget("432")
colnames(meta_selic) <- c("Data", "Meta Selic")

#Mudança de colunas para formato de data
nomes_colunas_para_data <- c("Data do leilão", "Data de liquidação", "Data de vencimento")
ltn <- para_data(ltn, nomes_colunas_para_data)
ntn_f <- para_data(ntn_f, nomes_colunas_para_data)

#Mudança de colunas para formato de numero
nomes_colunas_para_numero <- c("Taxa média", "Taxa de corte")
ltn <- para_numero(ltn, nomes_colunas_para_numero)
ntn_f <- para_numero(ntn_f, nomes_colunas_para_numero)

#Aplicando filtros
ltn <- filter(ltn, `Tipo de leilão` == "Venda",  Volta == "1.ª volta")
ntn_f <- filter(ntn_f, `Tipo de leilão` == "Venda",  Volta == "1.ª volta")

#Fazendo cálculos
datas_vencimento_ltn <- unique(ltn[,"Data de vencimento"])
datas_vencimento_ntn_f <- unique(ntn_f[,"Data de vencimento"])

ltn <- left_join(ltn, meta_selic, by = c("Data do leilão" = "Data"))
ntn_f <- left_join(ntn_f, meta_selic, by = c("Data do leilão" = "Data"))
#ltn$Spreads <- as.numeric(gsub(",", ".", ltn$`Taxa média`, fixed = TRUE)) - as.numeric(ltn$`Meta Selic`)
#ntn_f$Spreads <- as.numeric(gsub(",", ".", ntn_f$`Taxa média`, fixed = TRUE)) - as.numeric(ntn_f$`Meta Selic`)
ltn$Spreads <- ltn$`Taxa média` - ltn$`Meta Selic`
ntn_f$Spreads <- ntn_f$`Taxa média` - ntn_f$`Meta Selic`
lista_de_nomes_ltn <- vector()
lista_de_nomes_ntn_f <- vector()

#Exportação de arquivos
#LTN
for(i in 1:length(datas_vencimento_ltn)){
  filtrado <- filter(ltn, `Data de vencimento` == datas_vencimento_ltn[i]) %>% arrange(`Data do leilão`)
  nome_aba <- paste("LTN", datas_vencimento_ltn[i], sep = "-") %>% gsub(x = , pattern = "-", replacement = "_")
  lista_de_nomes_ltn[i] <- nome_aba
  assign(nome_aba, filtrado)
}

#NTN-F
for(i in 1:length(datas_vencimento_ntn_f)){
  filtrado <- filter(ntn_f, `Data de vencimento` == datas_vencimento_ntn_f[i])
  nome_aba <- paste("NTN_F", datas_vencimento_ntn_f[i], sep = "-") %>% gsub(x = , pattern = "-", replacement = "_")
  lista_de_nomes_ntn_f[i] <- nome_aba
  assign(nome_aba, filtrado)
}

#Gráficos

titulos_ltn <- ltn
titulos_ltn$`Data de vencimento` <- paste("LTN", titulos_ltn$`Data de vencimento`)
titulos_ntn_f <- ntn_f
titulos_ntn_f$`Data de vencimento` <- paste("NTN-F", titulos_ntn_f$`Data de vencimento`)
titulos <- rbind(titulos_ltn, titulos_ntn_f)

grafico_geral_tx <- function(df, titulo){
  ggplot(df, aes(x = `Data do leilão`, y = `Taxa média`, group = `Data de vencimento`, colour = factor(`Data de vencimento`))) +  
    geom_line(size = 1.5) + theme_minimal() + theme(legend.position = "none")+ scale_x_date(date_breaks = "1 month", date_labels = "%b/%Y", ) + 
    theme(axis.text.x = element_text(angle = 90, hjust = 1)) +
    labs(x = "Data do leilão", y = "% ao ano", title = titulo, colour = "Data de vencimento")
}

grafico_geral_spread <- function(df, titulo){
  ggplot(df, aes(x = `Data do leilão`, y = Spreads, group = `Data de vencimento`, colour = factor(`Data de vencimento`))) +  
    geom_line(size = 1.5) + theme_minimal() + theme(legend.position = "none")+ scale_x_date(date_breaks = "1 month", date_labels = "%b/%Y", ) + 
    theme(axis.text.x = element_text(angle = 90, hjust = 1)) +
    labs(x = "Data do leilão", y = "% ao ano", title = titulo, colour = "Data de vencimento")
}

graf_tx_ltn <- grafico_geral_tx(ltn, "Taxa média por LTN, data do leilão e data do vencimento")
graf_spread_ltn <- grafico_geral_spread(ltn, "Spreads sobre meta Selic por LTN, data do leilão e data do vencimento")
graf_tx_ntn_f <- grafico_geral_tx(ntn_f, "Taxa média por NTN-F, data do leilão e data do vencimento")
graf_spread_ntn_f <- grafico_geral_spread(ntn_f, "Spreads sobre meta Selic por NTN-F, data do leilão e data do vencimento")

graficos_data <- function(df, titulo){
  ggplot(df, aes(x = `Data do leilão`, y = Spreads, group = `Data de vencimento`,  label = round(Spreads, 2))) + 
    geom_line(size = 1.5) + theme_minimal() + theme(legend.position = "none")+ scale_x_date(date_breaks = "1 month", date_labels = "%b/%Y", ) + 
    theme(axis.text.x = element_text(angle = 90, hjust = 1)) + geom_label_repel() + scale_y_continuous(breaks = seq(-2,10,1)) +
    labs(x = "Data do leilão", y = "% ao ano", title = titulo, colour = "Data de vencimento")
}

lista_nomes_graficos_ltn <- NA

for(i in 1:length(lista_de_nomes_ltn)){
  grafico <- graficos_data(get(lista_de_nomes_ltn[i]), as.character(lista_de_nomes_ltn[i]))
  nome_grafico <- paste("grafico", lista_de_nomes_ltn[i], sep = "_")
  assign(nome_grafico, grafico)
  lista_nomes_graficos_ltn[i] <- nome_grafico
}

lista_nomes_graficos_ntn_f <- NA


for(i in 1:length(lista_de_nomes_ntn_f)){
  grafico <- graficos_data(get(lista_de_nomes_ntn_f[i]), as.character(lista_de_nomes_ntn_f[i]))
  nome_grafico <- paste("grafico", lista_de_nomes_ntn_f[i], sep = "_")
  assign(nome_grafico, grafico)
  lista_nomes_graficos_ntn_f[i] <- nome_grafico
}