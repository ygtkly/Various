library(ggplot2)
library(tidyverse)
library(dplyr)
library(RCurl)
library(gdata)
library(knitr)
library(stringr)
library(xml2)
library(stringi)
library(gt)
library(writexl)
library(scales)

options(digits = 2)
options(scipen=999)


#IMPORTING THE DATA

ana.tablo <- read.xls('ana_tablo.XLSX')

satis.stok <- read.xls("satis_stok.xlsx")

maliyet <- read.xls('maliyet.XLSX')


# CURRENT STOCK AND SALES DATA CLEANING AND ARRANGING


options(digits = 2)
options(scipen=999)

str(satis.stok)
str(ana.tablo)

satis.stok <- satis.stok %>% mutate(mlz.kodu = A_URUN_KODU)

satis.stok <- satis.stok %>% mutate(mlz.no = MALZEME_NO)

satis.stok <- satis.stok %>% distinct(mlz.no, .keep_all = T)

satis.stok <- satis.stok %>% mutate(STOK_MIKTAR = str_remove_all(STOK_MIKTAR, ","))
satis.stok <- satis.stok %>% mutate(CS_SIP_TOPLAM = str_remove_all(CS_SIP_TOPLAM, ","))
satis.stok <- satis.stok %>% mutate(CS_SEVK_MIKTAR = str_remove_all(CS_SEVK_MIKTAR, ","))
satis.stok <- satis.stok %>% mutate(CS_KALAN_MIKTAR = str_remove_all(CS_KALAN_MIKTAR, ","))

satis.stok <- satis.stok %>% mutate(STOK_MIKTAR = as.numeric(STOK_MIKTAR))
satis.stok <- satis.stok %>% mutate(CS_SIP_TOPLAM = as.numeric(CS_SIP_TOPLAM))
satis.stok <- satis.stok %>% mutate(CS_SEVK_MIKTAR = as.numeric(CS_SEVK_MIKTAR))
satis.stok <- satis.stok %>% mutate(CS_KALAN_MIKTAR = as.numeric(CS_KALAN_MIKTAR))

satis.stok <- satis.stok %>% mutate(across(STOK_MIKTAR : CS_KALAN_MIKTAR, ~ ifelse(is.na(.x), 0, .x))) %>% 
  as.data.frame()

satis.stok <- satis.stok %>% 
  mutate(sevk.oncesi.adet = ifelse(B_DONUSUM_BIRIM_KODU == 'dz', STOK_MIKTAR*12, STOK_MIKTAR))

satis.stok <- satis.stok %>% 
  mutate(siparis.adet = ifelse(B_DONUSUM_BIRIM_KODU == 'dz', CS_SIP_TOPLAM*12, CS_SIP_TOPLAM))

satis.stok <- satis.stok %>% 
  mutate(sevk.edilen.adet = ifelse(B_DONUSUM_BIRIM_KODU == 'dz', CS_SEVK_MIKTAR*12, CS_SEVK_MIKTAR))

satis.stok <- satis.stok %>% 
  mutate(kalan.adet = ifelse(B_DONUSUM_BIRIM_KODU == 'dz', CS_KALAN_MIKTAR*12, CS_KALAN_MIKTAR))

satis.stok <- satis.stok %>% 
  mutate(mevcut.satilabilir.adet = sevk.oncesi.adet - kalan.adet)

satis.stok <- satis.stok %>% 
  select(mlz.no, sevk.oncesi.adet, siparis.adet, sevk.edilen.adet, kalan.adet, mevcut.satilabilir.adet)


satis.stok$mlz.kodu

str(ana.tablo)
ana.tablo$S_MALZEME_KODU



#COST DATA CLEANING AND ARRANGING


options(digits = 2)
options(scipen=999)

str(maliyet)

maliyet$A_DNM_SONU_ORT_BIR_FIY

maliyet <- maliyet %>% filter(B_MLZ_KODU %in% as.vector(ana.tablo$S_MALZEME_KODU))

maliyet <- maliyet %>% mutate(A_DNM_SONU_ORT_BIR_FIY = str_remove_all(A_DNM_SONU_ORT_BIR_FIY, ","))

maliyet$A_DNM_SONU_ORT_BIR_FIY <- as.numeric(maliyet$A_DNM_SONU_ORT_BIR_FIY)


maliyet <- maliyet %>% 
  mutate(adet.birim.maliyet = ifelse(A_BIRIM == 'dz', A_DNM_SONU_ORT_BIR_FIY/12, A_DNM_SONU_ORT_BIR_FIY))

maliyet <- maliyet %>% select(B_MLZ_KODU, adet.birim.maliyet)
names(maliyet) <- c('mlz.kodu', 'adet.birim.maliyet')

#CREATING THE MAIN TABLE USING THE PREVIOUS TABLE AND CURRENT STOCK, SALES & COST TABLES


options(digits = 2)
options(scipen=999)

company.guncel.stok.satis <- data.frame(mlz.turu = ana.tablo$S_MLZ_TURU)
company.guncel.stok.satis$mlz.kodu <- ana.tablo$S_MALZEME_KODU
company.guncel.stok.satis$mlz.no <- ana.tablo$MALZEME_NO
company.guncel.stok.satis$mlz.adi <- ana.tablo$S_MALZEME_ADI
company.guncel.stok.satis$kolisaj <- ana.tablo$KOLİSAJ
company.guncel.stok.satis$cilt <- ana.tablo$CİLT 
company.guncel.stok.satis$kapak <- ana.tablo$KAPAK 
company.guncel.stok.satis$seri <- ana.tablo$SERİ 
company.guncel.stok.satis$mensei <- ana.tablo$MENŞEİ 
company.guncel.stok.satis$temel.birim <- ana.tablo$CF_TEMEL_BIRIM 

company.guncel.stok.satis$market.satis.adet <- ana.tablo$X2020.MARKET.SATIŞ 
company.guncel.stok.satis$market.satis.adet.19 <- ana.tablo$X2019.MARKET.SATIŞ 
company.guncel.stok.satis$market.satis.adet.18 <- ana.tablo$X2018.MARKET.SATIŞ 
company.guncel.stok.satis$market.satis.adet.17 <- ana.tablo$X2017.MARKET.SATIŞ 

company.guncel.stok.satis$satis.adet.19 <- ana.tablo$X2019.TOPLAM.SATIŞ
company.guncel.stok.satis$satis.adet.18 <- ana.tablo$X2018.TOPLAM.SATIŞ
company.guncel.stok.satis$satis.adet.17 <- ana.tablo$X2017.TOPLAM.SATIŞ

company.guncel.stok.satis$ana.kategori <- ana.tablo$ANA.KATEGORI
company.guncel.stok.satis$alt.aksiyon <- ana.tablo$ALT.AKSIYON
company.guncel.stok.satis$comment <- ana.tablo$FREDY.BEY.KOMMENTLER

company.guncel.stok.satis <- company.guncel.stok.satis %>% left_join(satis.stok, by = 'mlz.no')

company.guncel.stok.satis <- company.guncel.stok.satis %>% left_join(maliyet, by = 'mlz.kodu')

company.guncel.stok.satis <- distinct(company.guncel.stok.satis)

str(company.guncel.stok.satis)

#REMOVA COMMAS & CHANGE TO NUMERIC

options(digits = 2)
options(scipen=999)

company.guncel.stok.satis <- company.guncel.stok.satis %>%
  mutate(market.satis.adet = str_remove_all(market.satis.adet, ","))

company.guncel.stok.satis <- company.guncel.stok.satis %>% 
  mutate(market.satis.adet = as.numeric(market.satis.adet))


company.guncel.stok.satis <- company.guncel.stok.satis %>%
  mutate(market.satis.adet.19 = str_remove_all(market.satis.adet.19, ","))

company.guncel.stok.satis <- company.guncel.stok.satis %>% 
  mutate(market.satis.adet.19 = as.numeric(market.satis.adet.19))


company.guncel.stok.satis <- company.guncel.stok.satis %>%
  mutate(market.satis.adet.18 = str_remove_all(market.satis.adet.18, ","))

company.guncel.stok.satis <- company.guncel.stok.satis %>% 
  mutate(market.satis.adet.18 = as.numeric(market.satis.adet.18))


company.guncel.stok.satis <- company.guncel.stok.satis %>%
  mutate(market.satis.adet.17 = str_remove_all(market.satis.adet.17, ","))

company.guncel.stok.satis <- company.guncel.stok.satis %>% 
  mutate(market.satis.adet.17 = as.numeric(market.satis.adet.17))


company.guncel.stok.satis <- company.guncel.stok.satis %>%
  mutate(satis.adet.19 = str_remove_all(satis.adet.19, ","))

company.guncel.stok.satis <- company.guncel.stok.satis %>% 
  mutate(satis.adet.19 = as.numeric(satis.adet.19))


company.guncel.stok.satis <- company.guncel.stok.satis %>%
  mutate(satis.adet.18 = str_remove_all(satis.adet.18, ","))

company.guncel.stok.satis <- company.guncel.stok.satis %>% 
  mutate(satis.adet.18 = as.numeric(satis.adet.18))


company.guncel.stok.satis <- company.guncel.stok.satis %>%
  mutate(satis.adet.17 = str_remove_all(satis.adet.17, ","))

company.guncel.stok.satis <- company.guncel.stok.satis %>% 
  mutate(satis.adet.17 = as.numeric(satis.adet.17))


company.guncel.stok.satis <- company.guncel.stok.satis %>% 
  dplyr::mutate(across(sevk.oncesi.adet : adet.birim.maliyet, ~ ifelse(is.na(.x), 0, .x))) %>% 
  as.data.frame()

company.guncel.stok.satis <- company.guncel.stok.satis %>% 
  mutate(satilabilir.stok.tl = adet.birim.maliyet*mevcut.satilabilir.adet)


company.guncel.stok.satis <- company.guncel.stok.satis %>% 
  mutate(sevk.oncesi.stok.tl = adet.birim.maliyet*sevk.oncesi.adet)

company.guncel.stok.satis <- company.guncel.stok.satis %>% 
  mutate(son.2.yil.ort.satis.adet = ((siparis.adet - market.satis.adet) + (satis.adet.19 - market.satis.adet.19))/2)

company.guncel.stok.satis <- company.guncel.stok.satis %>% 
  mutate(son.3.yil.ort.satis.adet = ((siparis.adet - market.satis.adet) + (satis.adet.19 - market.satis.adet.19) +
                                       (satis.adet.18 - market.satis.adet.18))/3)

company.guncel.stok.satis <- company.guncel.stok.satis %>% 
  mutate(son.4.yil.ort.satis.adet = ((siparis.adet - market.satis.adet) + (satis.adet.19 - market.satis.adet.19) +
                                       (satis.adet.18 - market.satis.adet.18) + (satis.adet.17 - market.satis.adet.17) )/4)

company.guncel.stok.satis <- company.guncel.stok.satis %>% 
  mutate(stok.yaslandirma.2yil = son.2.yil.ort.satis.adet/mevcut.satilabilir.adet)

company.guncel.stok.satis <- company.guncel.stok.satis %>% 
  mutate(stok.yaslandirma.3yil = son.3.yil.ort.satis.adet/mevcut.satilabilir.adet)

company.guncel.stok.satis <- company.guncel.stok.satis %>% 
  mutate(stok.yaslandirma.4yil = son.4.yil.ort.satis.adet/mevcut.satilabilir.adet)

#REVISING THE COMMENTS BASED ON NEW STOCK DATA

company.guncel.stok.satis$ana.kategori[company.guncel.stok.satis$mevcut.satilabilir.adet < 0 & company.guncel.stok.satis$satilabilir.stok.tl == 0] <- "SORUNSUZ - EKSI STOK - ADET" 


company.guncel.stok.satis$ana.kategori[company.guncel.stok.satis$satilabilir.stok.tl < 0] <- "SORUNSUZ - EKSI STOK - TL"

company.guncel.stok.satis$ana.kategori[company.guncel.stok.satis$mevcut.satilabilir.adet == 0] <- "SORUNSUZ - SIFIR STOK"

levels(company.guncel.stok.satis$ana.kategori)



# CREATING STOCK AND SALES REPORTS & SUMMARIES AND VISUALIZATIONS
options(digits = 2)
options(scipen=999)

satilabilir.stok.maliyet <- company.guncel.stok.satis %>% filter(satilabilir.stok.tl > 0) %>%
  summarize(sat.stok.maliyet = sum(satilabilir.stok.tl)) %>% .$sat.stok.maliyet

satilabilir.standart.stok.maliyet <- company.guncel.stok.satis %>% 
  filter(satilabilir.stok.tl > 0 & mlz.turu != "ÖZEL DEFTERLER") %>%
  summarize(sat.standart.stok.tl = sum(satilabilir.stok.tl)) %>% .$sat.standart.stok.tl

sevk.oncesi.stok.maliyet <- company.guncel.stok.satis %>% filter(sevk.oncesi.stok.tl > 0) %>%
  summarize(sevk.oncesi.stok.maliyet = sum(sevk.oncesi.stok.tl)) %>% .$sevk.oncesi.stok.maliyet

sevk.oncesi.stok.adet <- company.guncel.stok.satis %>% filter(sevk.oncesi.adet > 0) %>%
  summarize(sevk.oncesi.stok.adet = sum(sevk.oncesi.adet)) %>% .$sevk.oncesi.stok.adet

satilabilir.stok.adet <- company.guncel.stok.satis %>% filter(mevcut.satilabilir.adet > 0) %>%
  summarize(sat.stok.adet = sum(mevcut.satilabilir.adet)) %>% .$sat.stok.adet

satilabilir.standart.stok.adet <- company.guncel.stok.satis %>% 
  filter(mevcut.satilabilir.adet > 0 & mlz.turu != "ÖZEL DEFTERLER") %>%
  summarize(sat.standart.stok.adet = sum(mevcut.satilabilir.adet)) %>% .$sat.standart.stok.adet

toplam.tablo <- matrix(c(sevk.oncesi.stok.adet,
                         sevk.oncesi.stok.maliyet,
                         satilabilir.stok.adet, 
                         satilabilir.stok.maliyet, 
                         satilabilir.standart.stok.adet, 
                         satilabilir.standart.stok.maliyet), ncol = 2, byrow = T)


rownames(toplam.tablo) <- c('Sevk_Oncesi_Stok', 'Satilabilir_Stok', 'Standard_Urun_Stok')

colnames(toplam.tablo) <- c('ADET', 'TL')

toplam.tablo <- as.data.frame(toplam.tablo)

toplam.tablo <- format(x = toplam.tablo, big.mark = ",")

toplam.tablo <- toplam.tablo %>% gt(rownames_to_stub = T) 
toplam.tablo <- toplam.tablo %>% tab_header(title = '8 OCAK 2021 TOPLAM TABLOSU')

#HIGHLY PROBLEMATIC STOCK

cok.sorunlu.total <- company.guncel.stok.satis %>% filter(str_detect(ana.kategori, 'COK SORUNLU')) %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

cok.sorunlu.guzelyazi <- company.guncel.stok.satis %>% filter(ana.kategori == "COK SORUNLU - GUZEL YAZI") %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

cok.sorunlu.formali <- company.guncel.stok.satis %>% filter(ana.kategori == "COK SORUNLU - FORMALI" ) %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

cok.sorunlu.flora <- company.guncel.stok.satis %>% filter(ana.kategori == "COK SORUNLU - FLORA" ) %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

cok.sorunlu.duzkaye <- company.guncel.stok.satis %>% filter(ana.kategori == "COK SORUNLU - DUZ KAYE" ) %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

cok.sorunlu.capricorn <- company.guncel.stok.satis %>% filter(ana.kategori == "COK SORUNLU - CAPRICORN" ) %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

cok.sorunlu.lotus <- company.guncel.stok.satis %>% filter(ana.kategori == "COK SORUNLU - Lotus" ) %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

cok.sorunlu.infinity <- company.guncel.stok.satis %>% filter(ana.kategori == "COK SORUNLU - Infinity" ) %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

#MODERATE STOCK

orta.seviye.total <- company.guncel.stok.satis %>% filter(str_detect(ana.kategori, 'ORTA')) %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl),
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

orta.seviye.bim <- company.guncel.stok.satis %>% filter(ana.kategori == "ORTA SEVIYE - OZEL URETIM" & mensei =="BİM") %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

orta.seviye.migros <- company.guncel.stok.satis %>% filter(ana.kategori == "ORTA SEVIYE - OZEL URETIM" & mensei =="MİGROS") %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

orta.seviye.unicef <- company.guncel.stok.satis %>% filter(ana.kategori == "ORTA SEVIYE - OZEL URETIM" & mensei =="UNICEF") %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

orta.seviye.elisi <- company.guncel.stok.satis %>% filter(ana.kategori == "ORTA SEVIYE - ELISI KAGIDI") %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

orta.seviye.fon <- company.guncel.stok.satis %>% filter(ana.kategori == "ORTA SEVIYE - FON KARTONU") %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

orta.seviye.karina <- company.guncel.stok.satis %>% filter(ana.kategori == "ORTA SEVIYE - KARINA") %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

orta.seviye.simple <- company.guncel.stok.satis %>% filter(ana.kategori == "ORTA SEVIYE - SIMPLE") %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

orta.seviye.pruva <- company.guncel.stok.satis %>% filter(ana.kategori == "ORTA SEVIYE - PRUVA") %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)


#UNPROBLEMATIC STOCK

sorunsuz.total <- company.guncel.stok.satis %>% filter(str_detect(ana.kategori, 'SORUNSUZ')) %>% 
  filter(!ana.kategori %in% c("SORUNSUZ - EKSI STOK - TL", "SORUNSUZ - EKSI STOK - ADET")) %>%
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl),
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

sorunsuz.adel <- company.guncel.stok.satis %>% filter(ana.kategori == "SORUNSUZ - OZEL URETIM") %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

sorunsuz.web <- company.guncel.stok.satis %>% filter(ana.kategori == "SORUNSUZ - OZEL WEB URUN") %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

sorunsuz.bloknot <- company.guncel.stok.satis %>% filter(ana.kategori == "SORUNSUZ - BLOKNOT") %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

sorunsuz.diger <- company.guncel.stok.satis %>% filter(ana.kategori == "SORUNSUZ") %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

sorunsuz.sifirstok <- company.guncel.stok.satis %>% filter(ana.kategori == "SORUNSUZ - SIFIR STOK") %>% 
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)

#TOTAL 

genel.total <- company.guncel.stok.satis %>% 
  filter(!ana.kategori %in% c("SORUNSUZ - EKSI STOK - ADET", "SORUNSUZ - EKSI STOK - TL")) %>%
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl), 
            adet.yuzde = urun.adedi/satilabilir.stok.adet, tl.yuzde = maliyet.tl/satilabilir.stok.maliyet)


#NEGATIVE STOCK

eksi.stok.total <- company.guncel.stok.satis %>% filter(mevcut.satilabilir.adet < 0) %>%
  summarize( urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl),
             adet.yuzde = 0, tl.yuzde = 0)

eksi.stok.adet <- company.guncel.stok.satis %>% filter(mevcut.satilabilir.adet < 0 & satilabilir.stok.tl == 0) %>%
  summarize( urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl),
             adet.yuzde = 0, tl.yuzde = 0)

eksi.stok.tl <- company.guncel.stok.satis %>% filter(satilabilir.stok.tl < 0 ) %>%
  summarize(urun.cesidi = n(), urun.adedi = sum(mevcut.satilabilir.adet), maliyet.tl = sum(satilabilir.stok.tl),
            adet.yuzde = 0, tl.yuzde = 0)


levels(company.guncel.stok.satis$mensei)
levels(company.guncel.stok.satis$ana.kategori)


#SUMMARY TABLE CREATION

options(digits = 2)
options(scipen=999)

satirlar <- matrix(c(cok.sorunlu.total, cok.sorunlu.guzelyazi, cok.sorunlu.formali, cok.sorunlu.flora, 
                     cok.sorunlu.duzkaye, cok.sorunlu.capricorn, cok.sorunlu.lotus, cok.sorunlu.infinity,
                     orta.seviye.total, orta.seviye.bim, orta.seviye.migros, orta.seviye.unicef, 
                     orta.seviye.elisi, orta.seviye.fon, orta.seviye.karina, orta.seviye.simple, 
                     orta.seviye.pruva, 
                     sorunsuz.total, sorunsuz.adel, sorunsuz.web, sorunsuz.bloknot, sorunsuz.diger, 
                     sorunsuz.sifirstok, 
                     genel.total, 
                     eksi.stok.tl, eksi.stok.adet, eksi.stok.total), ncol = 5, byrow = T)


options(digits = 2)
options(scipen=999)

ocak2021.ozet <- as.data.frame(satirlar)

colnames(ocak2021.ozet) <- c('Urun_Cesidi', 'Urun_Adedi', 'Urun_Maliyeti', 'Adet_Yuzde', 'TL_Yuzde')

rownames(ocak2021.ozet) <- c('Cok_Sorunlu_Toplam', 'GuzelYazi', 'Formali', 'Flora',
                             'Duz_Kaye', 'Capricorn', 'Lotus', 'Infinity', 
                             'Orta_Seviye_Toplam', 'Bim', 'Migros', 'Unicef', 
                             'Elis_Kagidi', 'Fon_Kartonu', 'Karina', 'Simple',
                             'Pruva', 
                             'Sorunsuz_Toplam', 'Adel', 'Web', 'Bloknot', 'Diger',
                             'Sifir_Stok', 
                             'Genel_Toplam', 
                             'Eksi_Stok_TL', 'Eksi_Stok_Adet', 'Eksi_Stok_Toplam')



ocak2021.ozet$Adet_Yuzde <- as.numeric(ocak2021.ozet$Adet_Yuzde)
ocak2021.ozet$TL_Yuzde <- as.numeric(ocak2021.ozet$TL_Yuzde)

ozet.ocak.2021.plot.icin <- ocak2021.ozet

ocak.2021.ozet.excel <- ocak2021.ozet
ocak.2021.ozet.excel$Urun_Cesidi <- as.numeric(ocak.2021.ozet.excel$Urun_Cesidi)
ocak.2021.ozet.excel$Urun_Adedi <- as.numeric(ocak.2021.ozet.excel$Urun_Adedi)
ocak.2021.ozet.excel$Urun_Maliyeti <- as.numeric(ocak.2021.ozet.excel$Urun_Maliyeti)
ocak.2021.ozet.excel$Stok_Durumu <- rownames(ocak.2021.ozet.excel)


ocak2021.ozet$Adet_Yuzde <- percent(ocak2021.ozet$Adet_Yuzde, accuracy = 0.01)
ocak2021.ozet$TL_Yuzde <- percent(ocak2021.ozet$TL_Yuzde, accuracy = 0.01)

ocak2021.ozet <- format(x = ocak2021.ozet, big.mark = ",", digit = 2)


ocak.2021.ozet.table <- gt(ocak2021.ozet, rownames_to_stub = T)

ocak.2021.ozet.table <- ocak.2021.ozet.table %>% tab_style(style = cell_fill(color = 'lightpink'), locations = cells_body(rows = 1:8)) %>%
  tab_style(style = cell_fill(color = 'lightblue'), locations = cells_body(rows = 9:17)) %>%
  tab_style(style = cell_fill(color = 'lightgreen'), locations = cells_body(rows = 18:23)) %>%
  tab_style(style = cell_fill(color = 'gold'), locations = cells_body(rows = 24)) %>%
  tab_style(style = cell_fill(color = 'bisque'), locations = cells_body(rows = 25:27)) %>%
  tab_style(style = cell_fill(color = 'gold'), locations = cells_body(rows = c(1, 9, 18)))


ocak.2021.ozet.table <- ocak.2021.ozet.table %>% tab_row_group(group = 'COK SORUNLU', rows = 1:8) %>%
  tab_row_group(group = 'ORTA SEVIYE', rows = 9:17) %>%
  tab_row_group(group = 'SORUNSUZ', rows = 18:23)

ocak.2021.ozet.table <- tab_header(ocak.2021.ozet.table, title = 'OCAK 2021 OZET')


#IMPORTING OLD SUMMARY TABLES 

ozet.tablo.eski <- read.xls("ozet_tablo.xlsx")

remove.index <- c(1, 9, 18, 24:27)

ozet.ocak.2021.plot.icin <- ozet.ocak.2021.plot.icin[-remove.index, ]

as.numeric(ozet.ocak.2021.plot.icin$Urun_Cesidi)

ozet.ocak.2021.plot.icin2 <- data.frame(tarih = rep("ocak.2021", 20) ,
                                        durum.main = ozet.tablo.eski$durum.main[1:20],
                                        durum.sub = ozet.tablo.eski$durum.sub[1:20],
                                        urun.cesidi = as.numeric(ozet.ocak.2021.plot.icin$Urun_Cesidi), 
                                        urun.adedi = as.numeric(ozet.ocak.2021.plot.icin$Urun_Adedi),
                                        maliyet = as.numeric(ozet.ocak.2021.plot.icin$Urun_Maliyeti),
                                        adet.yuzde = ozet.ocak.2021.plot.icin$Adet_Yuzde,
                                        tl.yuzde = ozet.ocak.2021.plot.icin$TL_Yuzde)


full.ozet.tablo <- rbind(ozet.tablo.eski, ozet.ocak.2021.plot.icin2)

full.ozet.tablo$tarih <- factor(full.ozet.tablo$tarih, 
                                levels = c('temmuz.2020', 'agustos.2020', 'eylul.2020', 
                                           'ekim.2020', 'kasim.2020', 'ocak.2021'))

levels(full.ozet.tablo$tarih)

#VISUALIZATION

adet.by.tarih <- full.ozet.tablo %>% group_by(tarih, durum.main) %>% summarize(toplam.adet = sum(urun.adedi), 
                                                                               toplam.maliyet = sum(maliyet)) %>%
  ggplot(aes(tarih, toplam.adet, color = durum.main, group = durum.main)) + geom_point() + geom_line() +
  facet_grid(durum.main~., scales = 'free_y') +
  theme(axis.text.x = element_text(angle = 90, hjust = 1, size = 12)) + ggtitle("Tarih Vs Stok Adet")


maliyet.by.tarih <- full.ozet.tablo %>% group_by(tarih, durum.main) %>% summarize(toplam.adet = sum(urun.adedi), 
                                                                                  toplam.maliyet = sum(maliyet)) %>%
  ggplot(aes(tarih, toplam.maliyet, color = durum.main, group = durum.main)) + geom_point() + geom_line() +
  facet_grid(durum.main~., scales = 'free_y') +
  theme(axis.text.x = element_text(angle = 90, hjust = 1, size = 12)) + ggtitle("Tarih Vs Stok Maliyet")

#EXPORT TABLES AND PLOTS

str(company.guncel.stok.satis)

company.guncel.stok.satis$mevcut.satilabilir.koli <- company.guncel.stok.satis$mevcut.satilabilir.adet/company.guncel.stok.satis$kolisaj

bitmis.urun.satis.stok.ocak.2021 <- company.guncel.stok.satis[, c("mlz.turu", "mlz.kodu", "mlz.no", "mlz.adi", "kolisaj",
                                                                   "cilt", "kapak", "seri", "mensei", "temel.birim",
                                                                   "sevk.oncesi.adet", "siparis.adet",
                                                                   "market.satis.adet","sevk.edilen.adet",
                                                                   "kalan.adet", "mevcut.satilabilir.adet", "mevcut.satilabilir.koli",
                                                                   "adet.birim.maliyet", "sevk.oncesi.stok.tl", "satilabilir.stok.tl",
                                                                   "satis.adet.19", "market.satis.adet.19", "satis.adet.18", "market.satis.adet.18",
                                                                   "satis.adet.17", "market.satis.adet.17", 
                                                                   "son.2.yil.ort.satis.adet", "son.3.yil.ort.satis.adet", "son.4.yil.ort.satis.adet", 
                                                                   "stok.yaslandirma.2yil", "stok.yaslandirma.3yil", "stok.yaslandirma.4yil", 
                                                                   "ana.kategori", "alt.aksiyon", "comment")]

write_xlsx(bitmis.urun.satis.stok.ocak.2021, 'Bitmis.Urun.Satis.Stok.Ocak.2021')

write_xlsx(full.ozet.tablo, 'ozet_tablo_tidy')

write_xlsx(ocak.2021.ozet.excel, 'ocak_21_ozet')

webshot::install_phantomjs()

ocak.2021.ozet.table %>% gtsave("ocak.2021.ozet.tablo.png")

toplam.tablo %>% gtsave("toplam_tablo_ocak2021.png")
