# 🏛️ KMS Aktivátor Magyar Közigazgatási EDU

Ez a program egy grafikus felhasználói felülettel (GUI) rendelkező KMS (Key Management Service) aktivátor, amelyet a Microsoft Windows és Office szoftverek **jogtiszta aktiválására** hoztak létre, a **Magyar Közigazgatási EDU hálózati** címet használva.

---

## Cél és Működés

A KMS egy licenckezelési módszer, amelyet nagyvállalatok és oktatási intézmények használnak. Ahelyett, hogy minden gép egyedi kulccsal aktiválódna, a hálózatban lévő számítógépek egy központi KMS szerverhez kapcsolódnak a licenc megújítása céljából.

### Főbb Funkciók:
1. **GVLK kulcs telepítése:** Beírja a szükséges licenckulcsot (pl. Win 11 Enterprise, Office 2021).
2. **KMS szerver beállítása:** Beállítja a hálózat KMS szerverének címét (`kms.edu.hu`).
3. **Azonnali aktiválás:** Futtatja az aktiválási parancsot.
4. **Státusz ellenőrzés és kulcs eltávolítás.**

## ⚠️ Fontos Figyelmeztetés!

A program alacsony szintű rendszerparancsokat futtat, ezért **Rendszergazdai Jogosultság** szükséges a futtatásához. Mind az indító `.py` szkriptet, mind a belőle készült `.exe` fájlt **"Futtatás rendszergazdaként"** opcióval kell indítani.

---

## 💻 Használat

1.  **Válaszd ki a Terméket:** Használd a TabView füleket (`Windows`, `Server`, `Office`) a megfelelő termék kiválasztásához.
2.  **Aktiválás:** Kattints a kívánt verziójú szoftverhez tartozó **"Aktiválás"** gombra.
3.  **Visszajelzés:** Kövesd a folyamatot a "Visszajelzési Folyamat / Kimenet" dobozban.

---

## Készítők

**Készítette:** Csipai Gergő 2025 (Eredeti: Ocskó Zsolt & Lovász Ádám)
