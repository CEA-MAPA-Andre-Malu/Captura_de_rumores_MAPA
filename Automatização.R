install.packages(c("readxl", "tidyverse", "httr", "dplyr", "magrittr", "stringr"))

library(readxl)
library(tidyverse)
library(httr)
library(dplyr)
library(magrittr)
library(stringr)



#### MALU QUICES ----
rm_accent <- function(str,pattern="all") {
  # Rotinas e funções úteis V 1.0
  # rm.accent - REMOVE ACENTOS DE PALAVRAS
  # Função que tira todos os acentos e pontuações de um vetor de strings.
  # Parâmetros:
  # str - vetor de strings que terão seus acentos retirados.
  # patterns - vetor de strings com um ou mais elementos indicando quais acentos deverão ser retirados.
  #            Para indicar quais acentos deverão ser retirados, um vetor com os símbolos deverão ser passados.
  #            Exemplo: pattern = c("´", "^") retirará os acentos agudos e circunflexos apenas.
  #            Outras palavras aceitas: "all" (retira todos os acentos, que são "´", "`", "^", "~", "¨", "ç")
  if(!is.character(str))
    str <- as.character(str)
  
  pattern <- unique(pattern)
  
  if(any(pattern=="Ç"))
    pattern[pattern=="Ç"] <- "ç"
  
  symbols <- c(
    acute = "áéíóúÁÉÍÓÚýÝ",
    grave = "àèìòùÀÈÌÒÙ",
    circunflex = "âêîôûÂÊÎÔÛ",
    tilde = "ãõÃÕñÑ",
    umlaut = "äëïöüÄËÏÖÜÿ",
    cedil = "çÇ"
  )
  
  nudeSymbols <- c(
    acute = "aeiouAEIOUyY",
    grave = "aeiouAEIOU",
    circunflex = "aeiouAEIOU",
    tilde = "aoAOnN",
    umlaut = "aeiouAEIOUy",
    cedil = "cC"
  )
  
  accentTypes <- c("´","`","^","~","¨","ç")
  
  if(any(c("all","al","a","todos","t","to","tod","todo")%in%pattern)) # opcao retirar todos
    return(chartr(paste(symbols, collapse=""), paste(nudeSymbols, collapse=""), str))
  
  for(i in which(accentTypes%in%pattern))
    str <- chartr(symbols[i],nudeSymbols[i], str)
  
  return(str)
}

delete.na <- function(DF, n=5) {
  DF[!(c(rowSums(is.na(DF))) >= n),]
}


#PC NO MAPA
#TROCAR O SETWD PELO DIRETORIO CORRETO DO SEU COMPUTADOR
setwd("C:/Users/andre.vale/Downloads")
#TROCAR A DATA DA PLANILHA
Teste_Aut <- read_xlsx('Dia 12_07_2023.xlsx')
names(Teste_Aut)[20] <- "__PowerAppsId__"

Aut_GA <- Teste_Aut[which(Teste_Aut$`FERRAMENTA DE BUSCA` == "GOOGLE ALERTA"),]
Aut_PRO <- Teste_Aut[which(Teste_Aut$`FERRAMENTA DE BUSCA` == "PROMED" |Teste_Aut$`FERRAMENTA DE BUSCA` == "PROMED Plant"),]
Aut_WOAH <- Teste_Aut[which(Teste_Aut$`FERRAMENTA DE BUSCA` == "WOAH"),]
Aut_PRO$ASSUNTO <- Aut_PRO$ASSUNTO %>%
  str_remove(".+\\> ") %>%
  str_remove(" \\(.+|\\:|\\-.+") %>%
  str_trim()

codigo <- Teste_Aut$`__PowerAppsId__`[Teste_Aut$`FERRAMENTA DE BUSCA` == 'GOOGLE ALERTA']

newrowsga <- data.frame('1' = NA, '2' = NA, '3' = NA, '4' = NA, '5' = NA,
                        '6' = NA, '7' = NA, '8' = NA, '9' = NA, '10' = NA, '11' = NA,
                        '12' = NA, '13' = NA, '14' = NA, '15' = NA, '16' = NA, '17' = NA,
                        '18' = NA, '19' = NA, '20' = NA)


newrowspro <- data.frame(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA)
newrowswoah <- data.frame(matrix(ncol = 20, nrow = nrow(Aut_WOAH)))


names(newrowsga) <- names(Teste_Aut)
names(newrowspro) <- names(Teste_Aut)
names(newrowswoah) <- names(Teste_Aut)


for(j in 1:nrow(Aut_GA)){
  a <- as.character(Aut_GA[j, "TÍTULO"])
  a <- unlist(strsplit(a, split = '",'))
  v <- 1
  titulo <- c()
  link <- c()
  for (i in 1:(length(a))) {
    if(!is.na(a[i])){
      if (a[i] == '\"Sinalizar como irrelevante') {
        a <- a[-c((i+1):(i-4))]
      }
    }
  }
  for (i in 1:(length(a))) {
    if(!is.na(a[i])){
      if (a[i] == '\"NOTÍCIAS') {
        a <- a[-c(i:(i-5))]
      }
    }
  }
  for (i in 1:(length(a))) {
    if(!is.na(a[i])){
      if (a[i] == '\"Ver mais resultados'){
        a <- a[-c((i-1):(i+14))]
      }
    }
  }
  for (t in 2:length(a)) {
    b <- unlist(strsplit(a[t], ""))
    
    if ((a[t] != '"') & (a[t] != "") & (length(b) >= 6)){
      if (all(b[c(3:6)] == c("h", "t", "t", "p"))){
        if (nchar(a[t-1]) < 20) {
          a[t-1] <- substr(a[t-1], start = 2, stop = nchar(a[t-1]))
          a[t-2] <- substr(a[t-2], start = 2, stop = nchar(a[t-2]))
          titulo[v] <- paste(a[t-2], a[t-1])
          link[v] <- a[t]
          link[v] <- substr(link[v], start = 3, stop = nchar(link[v])-1)
          v <- v + 1
        }
        else{
          a[t-1] <- substr(a[t-1], start = 2, stop = nchar(a[t-1]))
          titulo[v] <- a[t-1]
          link[v] <- a[t]
          link[v] <- substr(link[v], start = 3, stop = nchar(link[v])-1)
          v = v + 1
        }
      }
    }
  }
  
  
  
  googlealerts <- data.frame('1' = NA, '2' = NA, '3' = NA, '4' = NA, '5' = NA,
                             '6' = NA, '7' = NA, '8' = NA, "9" = titulo, '10' = link, '11' = NA,
                             '12' = NA, '13' = NA, '14' = NA, '15' = NA, '16' = NA, '17' = NA,
                             '18' = NA, '19' = NA, '20')
  
  names(googlealerts) <- names(Teste_Aut)
  
  googlealerts[,c(1,8,11,20)] <- Aut_GA[j,c(1,8,11,20)]
  
  newrowsga <- rbind(googlealerts, newrowsga)
  
}


linha <- 1
for(j in 1:nrow(Aut_PRO)){
  if (!is.na(Aut_PRO[j,"__PowerAppsId__"])) {
    sepvec <- unlist(strsplit(as.character(Aut_PRO[j,"LINK"]),'",'))
    names(newrowspro) <- names(Teste_Aut)
    for(z in 1:length(sepvec)){
      if(grepl("Source",  sepvec[z] , fixed = TRUE)){
        for(l in z:length(sepvec)){
          if(grepl("http",  sepvec[l] , fixed = TRUE)){
            newrowspro[linha,] <- NA
            newrowspro[linha,"LINK"] <- gsub('"','', sepvec[l])
            newrowspro[linha,c(1,8,9,11,20)] <- Aut_PRO[j,c(1,8,9,11,20)]
            linha <- linha+1
            break
          }
        }
      }
    }
  }
}


linha <- 1
for(j in 1:nrow(Aut_WOAH)){
  newrowswoah$TÍTULO[j] <- str_trim(strsplit(strsplit(Aut_WOAH$TÍTULO[j], " /")[[1]][1], " — ")[[1]][2])
  newrowswoah$PAÍS[j] <- strtrim(Aut_WOAH$TÍTULO[j], 3)
  sepvec <- unlist(strsplit(as.character(Aut_WOAH[j,10]),'",'))
  for(z in 1:length(sepvec)){
    if(grepl("Consult the report",  sepvec[z] , fixed = TRUE)){
      newrowswoah[j,10] <- substr(sepvec[z+1],3,nchar(sepvec[z+1])-1)
    }
  }
  newrowswoah[j, c(1,8,20)] <- Aut_WOAH[j, c(1,8,20)]
}



for(j in 1:nrow(newrowspro)){
  newrowspro[j,9] <- unlist(strsplit(as.character(newrowspro[j,9]),'",'))[1]
  newrowspro[j,9] <- substr(newrowspro[j,9], start = 3, stop = nchar(newrowspro[j,9]))
}

for(j in 1:nrow(newrowsga)){
  newrowsga[j,11] <- toupper(rm_accent(gsub('"','',substr(newrowsga[j,11],19,nchar(newrowsga[j,11])))))
}

newrowsga[,10] <- gsub("^.*?&url=","", newrowsga[,"LINK"])
newrowsga[,10] <- gsub("&ct=.*","",newrowsga[,"LINK"])


noticias <- rbind(newrowsga, newrowspro, newrowswoah)
noticias <- noticias[,-c(13:length(noticias))]
noticias <- delete.na(noticias, n=12)
noticias$ASSUNTO <- trimws(noticias$ASSUNTO)
noticias$ASSUNTO <- toupper(noticias$ASSUNTO)

noticias <- transform(noticias, `RESULTADO DA AVALIAÇÃO`= ifelse(`ASSUNTO` == "COVID" |`ASSUNTO` == "DENGUE" | `ASSUNTO` == "MPOX"  , "DESCARTADO", `RESULTADO DA AVALIAÇÃO`))

form.ori <- c()
form.cont <- c()
form.se <- c()
form.est <- c()


for (i in 2:(nrow(noticias)+1)){
  form.ori <- c(form.ori, paste('=IF(ISBLANK($E', i, ');"";IF($E', i, '="BRASIL"; "NACIONAL"; "INTERNACIONAL"))', sep = ''))
  form.cont <- c(form.cont, paste('=IF(ISBLANK($E',i,');"";VLOOKUP($E',i,';REFERENCIA!$K:$L;2;0))', sep = ''))
  form.est <- c(form.est, paste('=IF(ISBLANK($G',i,');"";VLOOKUP($G',i,';REFERENCIA!$O:$P;2;0))', sep = ''))
  form.se <- c(form.se, paste('=IF(ISBLANK($A', i, ');"";WEEKNUM($A', i, '))', sep = ''))
}

noticias$ORIGEM <- form.ori
noticias$CONTINENTE <- form.cont
noticias$ESTADO <- form.est
noticias$SE <- form.se

#TROCAR A DATA DA PLANILHA
write.csv(noticias, "12_07_2023 nova.xlsx")
