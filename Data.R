library(lubridate)
library(ggplot2)
library(dplyr)
library(reshape2)
library(stringr)
library(knitr)
library(kableExtra)
library(googledrive)
library(googlesheets4)
library(forcats)
library(readxl)
library(WriteXLS)
library(readxl)
library(openxlsx)

# Identificarse----------------------------------
#Correo_electrónico <- "sefa@oefa.pe"
#Identificarse en Google:
#drive_auth(email = Correo_electrónico)
#gs4_auth(token = drive_token())

# Importando y guardando data --------------------------------

# CONEXION A BASE DE DATOS
SEGUIMIENTO <- "https://docs.google.com/spreadsheets/d/1Um2C8IhkGTjUDLYalsp3o3y5kRSng15wWSbG13Jl5X0/edit#gid=1763680572"
DERIVACIONES<- as.data.frame(read_sheet(SEGUIMIENTO, sheet = "Derivaciones"))
write.xlsx(DERIVACIONES,file="DERIVACIONES.xlsx", sheetName="DERIVACIONES")

CALCULADORA <- "https://docs.google.com/spreadsheets/d/e/2PACX-1vR2hNks3EIqx1T_KCjatIxBTbu3L4p08IHUeIUIneV44NRDZmcmc4tNrYIUbA17Ls3OGLXRvnkDfZUS/pub?output=xlsx"
tp1 <- tempfile()
download.file(CALCULADORA, tp1, mode ="wb")
CALCULADORA <- as.data.frame(read_xlsx(tp1, sheet = "Calculadora"))
write.xlsx(CALCULADORA,file="CALCULADORA.xlsx", sheetName="CALCULADORA")

POI <- "https://docs.google.com/spreadsheets/d/e/2PACX-1vRTICHUqIMYIaWOWoojCeetIj1_aax-LZ2DW0SR2rpMq9ZsqvyBRfcYGPjzjJuUBHCjD0qf0HEjkWjh/pub?output=xlsx"
tp1 <- tempfile()
download.file(POI, tp1, mode ="wb")
POI <- as.data.frame(read_xlsx(tp1, sheet = "POI SINADA"))
write.xlsx(POI,file="POI.xlsx", sheetName="POI")


# GUARDANDO EN CARPETA

write.xlsx(CALCULADORA, ss= "DENUNCIASSINADA.xlsx", sheet= "CALCULADORA")

#Carpeta donde se guardará las hojas de cálculo
#Carpeta_denuncias <- "https://drive.google.com/drive/folders/1Cad-AaQCcpegJucfzxC1TrxnrPC2kIpu?usp=sharing"

# Descargando data en la carpeta drive
# crear objeto con el nombre del archivo:
Nombre_del_archivo1 <- "DENUNCIASSINADA"
Nombre_completo_del_archivo_Drive1 <- paste(Nombre_del_archivo1, now())
# Crear archivo en Drive
Archivo_en_Drive1 <- gs4_create(name = Nombre_completo_del_archivo_Drive1)
# Escribir la hoja de cálculo en el archivo creado en Drive:
sheet_write(SEGUIMIENTO, ss= Archivo_en_Drive1, sheet = 'Denuncias')
sheet_write(CALCULADORA, ss= Archivo_en_Drive1, sheet = 'calculadora')
sheet_write(POI, ss= Archivo_en_Drive1, sheet = 'POI')
# Borrar pestaña vacía
sheet_delete(ss= Archivo_en_Drive1, sheet = 'Hoja 1')
# Cambiar de lugar el archivo que hemos creado en Drive:
drive_mv(Archivo_en_Drive1, path = as_id(Carpeta_denuncias))

#EXPORTAR DATA--------------------------------------------------------

#Generar un objeto con la hoja de cálculo (Estamos convirtiendo el link de la base de datos en un objeto):
#CARPETA_FICHA_CIERRE <- "https://docs.google.com/spreadsheets/d/1X4HyiumqInd0rIO8dAdepDNJKmvUZ2EiOTI2STshdz4/edit?usp=sharing"

#Leer la hoja de cálculo y convertirla en un objeto (esto puede tardar unos momentos):


#DERIVACIONES <- read_sheet(CARPETA_FICHA_CIERRE, sheet = "Denuncias")
#CALCULADORA_RIESGO <- read_sheet(CARPETA_FICHA_CIERRE, sheet = "calculadora")
#POI2 <- read_sheet(CARPETA_FICHA_CIERRE, sheet = "POI")


CALCULADORA_FILTRADA <- select(
  CALCULADORA,
  DENUNCIA = 'CODIGO_SINADA_FINAL',
  DEPARTAMENTO = 'DEPARTAMAMENTO (automático)',
  PROVINCIA = 'PROVINCIA (automático)',
  DISTRITO = 'DISTRITO (automático)',
  COMPONENTE = 'COMPONENTE',
  AGENTE = 'AGENTE',
  ACTIVIDAD = 'ACTIVIDAD',
  EXTENSION = 'EXTENSIÓN',
  UBICACION = 'UBICACIÓN',
  OCURRENCIA = 'OCURRENCIA',
  Resultado = 'Resultado',
  Amerita_seguimiento = 'Amerita seguimiento',
  Observaciones = 'Observaciones(opcional)',
  Especialista = 'Especialista (obligatorio)',
  Para_emitir = 'GENERAR',
  HT_TRASLADO = 'HT DE TRASLADO (obligatorio)',
  Fecha_cierre = 'Fecha cierre'
)

LISTA <- filter(CALCULADORA_FILTRADA, CALCULADORA_FILTRADA$Para_emitir == "Si")

DATOS_DENUNCIA <- data.frame(
  "DENUNCIA" = LISTA $DENUNCIA,
  "ESPECIALISTA" = LISTA $Especialista)

DATOS_DENUNCIA2 <- distinct(DATOS_DENUNCIA)

#KAREM VIÑAS
denuncias_karem <- DATOS_DENUNCIA2 %>% filter(ESPECIALISTA == "Karem Viñas")
### EXPORTAR EN FORMATO EXCEL 
#install.packages("WriteXLS")
library(WriteXLS)
library(readxl)
library(openxlsx)
karem_denuncias <- as.data.frame(denuncias_karem)
write.xlsx(karem_denuncias,file="Karem_denuncias.xlsx", sheetName="DATA")

#MONICA ARCE
denuncias_monica <- DATOS_DENUNCIA2 %>% filter(ESPECIALISTA == "Mónica Arce")
### EXPORTAR EN FORMATO EXCEL 
#install.packages("WriteXLS")
library(WriteXLS)
library(readxl)
library(openxlsx)
monica_denuncias <- as.data.frame(denuncias_monica)
write.xlsx(monica_denuncias,file="Monica_denuncias.xlsx", sheetName="DATA")

#RAUL VARGAS
denuncias_raul <- DATOS_DENUNCIA2 %>% filter(ESPECIALISTA == "Raúl Vargas")
### EXPORTAR EN FORMATO EXCEL 
#install.packages("WriteXLS")
library(WriteXLS)
library(readxl)
library(openxlsx)
raul_denuncias <- as.data.frame(denuncias_raul)
write.xlsx(raul_denuncias,file="Raul_denuncias.xlsx", sheetName="DATA")

#INGRIT
denuncias_ingrit <- DATOS_DENUNCIA2 %>% filter(ESPECIALISTA == "Ingrit Curo")
### EXPORTAR EN FORMATO EXCEL 
#install.packages("WriteXLS")
library(WriteXLS)
library(readxl)
library(openxlsx)
ingrit_denuncias <- as.data.frame(denuncias_ingrit)
write.xlsx(ingrit_denuncias,file="Ingrit_denuncias.xlsx", sheetName="DATA")

#ANA
denuncias_ana <- DATOS_DENUNCIA2 %>% filter(ESPECIALISTA == "Ana Paula Saravia")
### EXPORTAR EN FORMATO EXCEL 
#install.packages("WriteXLS")
library(WriteXLS)
library(readxl)
library(openxlsx)
ana_denuncias <- as.data.frame(denuncias_ana)
write.xlsx(ana_denuncias,file="Ana_denuncias.xlsx", sheetName="DATA")

#NATHALI
denuncias_nathali <- DATOS_DENUNCIA2 %>% filter(ESPECIALISTA == "Nathali Bardalez")
### EXPORTAR EN FORMATO EXCEL 
#install.packages("WriteXLS")
library(WriteXLS)
library(readxl)
library(openxlsx)
nathali_denuncias <- as.data.frame(denuncias_nathali)
write.xlsx(nathali_denuncias,file="Nathali_denuncias.xlsx", sheetName="DATA")

#PAUL
denuncias_paul <- DATOS_DENUNCIA2 %>% filter(ESPECIALISTA == "Paul Díaz")
### EXPORTAR EN FORMATO EXCEL 
#install.packages("WriteXLS")
library(WriteXLS)
library(readxl)
library(openxlsx)
paul_denuncias <- as.data.frame(denuncias_paul)
write.xlsx(paul_denuncias,file="Paul_denuncias.xlsx", sheetName="DATA")

#ALBERT
denuncias_albert <- DATOS_DENUNCIA2 %>% filter(ESPECIALISTA == "Albert Vila")
### EXPORTAR EN FORMATO EXCEL 
#install.packages("WriteXLS")
library(WriteXLS)
library(readxl)
library(openxlsx)
albert_denuncias <- as.data.frame(denuncias_albert)
write.xlsx(albert_denuncias,file="Albert_denuncias.xlsx", sheetName="DATA")

#PETER
denuncias_peter <- DATOS_DENUNCIA2 %>% filter(ESPECIALISTA == "Peter Fernández")
### EXPORTAR EN FORMATO EXCEL 
#install.packages("WriteXLS")
library(WriteXLS)
library(readxl)
library(openxlsx)
peter_denuncias <- as.data.frame(denuncias_peter)
write.xlsx(peter_denuncias,file="Peter_denuncias.xlsx", sheetName="DATA")

#ERNESTO
denuncias_ernesto <- DATOS_DENUNCIA2 %>% filter(ESPECIALISTA == "Ernesto Salamanca")
### EXPORTAR EN FORMATO EXCEL 
#install.packages("WriteXLS")
library(WriteXLS)
library(readxl)
library(openxlsx)
ernesto_denuncias <- as.data.frame(denuncias_ernesto)
write.xlsx(ernesto_denuncias,file="Ernesto_denuncias.xlsx", sheetName="DATA")

#federico
denuncias_federico <- DATOS_DENUNCIA2 %>% filter(ESPECIALISTA == "Federico Murriel")
### EXPORTAR EN FORMATO EXCEL 
#install.packages("WriteXLS")
library(WriteXLS)
library(readxl)
library(openxlsx)
federico_denuncias <- as.data.frame(denuncias_federico)
write.xlsx(federico_denuncias,file="Federico_denuncias.xlsx", sheetName="DATA")

##dataaa
DERIVACIONES <- as.data.frame(DERIVACIONES)
write.xlsx(DERIVACIONES,file="DERIVACIONES.xlsx", sheetName="DERIVACIONES")

CALCULADORA <- as.data.frame(CALCULADORA)
write.xlsx(CALCULADORA,file="CALCULADORA_RIESGO.xlsx", sheetName="CALCULADORA_RIESGO")

POI <- as.data.frame(POI)
write.xlsx(POI,file="POI.xlsx", sheetName="POI")


