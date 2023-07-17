# Proiect Gestiunea contului bancar

### Descrierea proiectului

Programul primește drept bază de date un document Excel. Fiecare foaie conține 
numele unei persoane, IBAN-ul contului bancar și un istoric de tranzacții. 
Pornind de la aceste date, utilizatorul este capabil:
* să afișeze datele contului (nume, IBAN)
* să afișeze extrasul de cont complet sau pe o anumită perioadă de timp
* să afle soldul curent 
* să efectueze noi tranzacții (depunere, extragere numerar sau altele),  

La final, programul generează un nou fișier Excel conținând atât tranzacțiile vechi, cât și pe cele noi. 

#### Biblioteci utilizate: _tabulate_, _re_, _openpyxl_, _datetime_.