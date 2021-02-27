#           ****************************                                 
#            ****** Cliente ********    
#           ****************************  

# ASSUNTO: SEPARACAO DA PLANILHA DE CONTROLE DE PRECOS, POR MODULOS.
# CONTRATO: xxxxxxx
 
# PASSO 1 - CARREGANDO PACOTES

  library(openxlsx)
  library(tidyverse)
  library(stringr)

# PASSO 2 - LEITURA DO ARQUIVO PRINCIPAL ('PLANILHA DE CONTROLE')
  
# E IMPORTANTE QUE A PLANILHA DE CONTROLE MANTENHA A MESMA SEQUENCIA DAS ABAS
  
# O "CHUNK" DO CODIGO A SEGUIR SELECIONA AS VARIAVEIS NECESSÁRIAS:
#         .ITEM xxxx
#         .CODIGO EXTERNO
#         .CODIGO BASE
#         .PRECO UNITARIO


  arquivo_da_pesquisa <- dir("planilha", full.names = T) %>% .[str_detect(., ".xlsx")]
  
  planilha_controle <- read.xlsx(arquivo_da_pesquisa, sheet = 2, startRow = 4) %>%
    select(c(1,2,9,38)) %>%
    rename('Item xxxx' = `Código.do.item.(xxxxx)`,
           Externo = Código.do.Item.xxxxx, 
           Cod.Base = Cód.Base, 
           preço = `Preço.Unitário.-.novo`) %>%
    mutate(., Externo = as.character(.$Externo))
  
# Passo 3 - Criando arquivos para fazer filtro
  
# As abas no arquivo planilha de controle, com os 3 modulos, foram criadas no proprio arquivo Excel
  
  filtros_lista <- list()
  
  for (i in 1:3) {
    
    filtros_lista[[i]] <- read.xlsx(arquivo_da_pesquisa, sheet = (i + 3)) %>%
      mutate(., Externo = as.character(.$Externo))
    
  }
  
  
  planilha_modulo <- list()
  
  for (i in 1:3) {
    
    planilha_modulo[[i]] <- right_join(planilha_controle, filtros_lista[[i]], by = "Externo") %>%
      na.omit()
    
  }

##############################################

#exportando para .xlsx
  
write.xlsx(planilha_modulo,"resultado/precos_controle.xlsx", sheetName = c("Mod1","Mod2","Mod3"))


if (length(dir("resultado/Comparar/", pattern = "precos_comparacao", all.files = FALSE)) == 0) { ## all.files - apenas os arquivos visíveis
  
  #escrever arquivos na pasta
  
  write.xlsx(planilha_modulo,"resultado/Comparar/precos_comparacao.xlsx", sheetName = c("Mod1","Mod2","Mod3"))
  
  #escrever arquivos na pasta de comparacao
  
  #write.xlsx(planilha_modulo,"resultado/Comparado/teste2.xlsx", sheetName = c("Mod1","Mod2","Mod3"))
  
  
} else {
  
  #criar um arquivo comparado
  
  comparar_lista <- list ()
  
  for (i in 1:3) {
    
  comparar_lista [[i]] <- read.xlsx("resultado/Comparar/precos_comparacao.xlsx", sheet = i)
    
  }
  
  write.xlsx(planilha_modulo,"resultado/Comparar/precos_comparacao.xlsx", sheetName = c("Mod1","Mod2","Mod3"))
  
  planilha_modulo2 <- list()
  
  for (i in 1:3) {
    
    planilha_modulo2[[i]] <- anti_join(planilha_modulo[[i]], comparar_lista[[i]], by = "Cod.Base") %>%
       na.omit()
    
  }
  
  write.xlsx(planilha_modulo2,"resultado/Comparado/precos_comparado.xlsx", sheetName = c("Mod1","Mod2","Mod3"))
  
}

  




