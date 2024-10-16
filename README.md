# TAB skriptas

## Kaip įsikelti kodą į savo kompiuterį naudojantis "Visual studio code"
Pirma, susikurkite aplanką, kuriame norite laikyti šį kodą. 
Antra, atsidarykite aplanką per "Visual studio code" programą.

Naudokite šią komandą "Visual studio Code" programoje:
```
git clone https://gitlab.com/api-duomenu-importavimas/TAB.git
```
## Norint įkelti pakeitimus į šią repozitoriją
Pirma, įsijungę Visual studio code naudokite šias komandas:
```
git config --global user.name "įveskite savo vardą"
git config --global user.email "įveskite savo el. paštą"
```
Antra, "path/to/TAB" repozitoriją reikia pakeisti tiksliu aplanko keliu. 
Tai padarę, galite įkelti pakeitimus į repozitoriją:
```
cd path/to/TAB
git remote add origin https://gitlab.com/api-duomenu-importavimas/TAB.git
git branch -M main
git push -uf origin main
```

Jeigu jus ištiks tokia klaida: 
```
fatal: unable to access 'https://gitlab.com/api-duomenu-importavimas/TAB.git/': SSL certificate problem: self-signed certificate in certificate chain
```
Naudokite šią komandą:
```
git -c http.sslVerify=false push -uf origin main
```

